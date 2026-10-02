"""공급계약서 자동작성 — 사내망용 작은 웹서버.

  python server.py            # 사내망 공개 (접속키 필요)
  python server.py --local    # 이 PC 에서만

이 PC 에서는 http://localhost:8790/ 로, 다른 PC 에서는 http://<이 PC IP>:8790/?k=<접속키> 로 들어온다.
접속키는 app\\_key.txt 에 있다.
"""
import json
import logging
import os
import random
import re
import shutil
import socket
import string
import sys
import threading
import time
import traceback
import urllib.parse
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

APP = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(APP), 'docs', 'py'))   # 엔진은 웹 버전과 같은 코드를 쓴다

from engine import PlanError, analyze, build   # noqa: E402
from report import make_report                  # noqa: E402

PORT = 8790
JOBS = os.path.join(APP, '_jobs')
KEYFILE = os.path.join(APP, '_key.txt')
URLFILE = os.path.join(APP, '_url.txt')
LOGFILE = os.path.join(APP, '_server.log')
MAX_UPLOAD = 60 * 1024 * 1024
KEEP_DAYS = 3

os.makedirs(JOBS, exist_ok=True)
logging.basicConfig(filename=LOGFILE, level=logging.INFO, encoding='utf-8',
                    format='%(asctime)s %(levelname)s %(message)s')
log = logging.getLogger('server')


def get_key():
    if os.path.exists(KEYFILE):
        k = open(KEYFILE, encoding='utf-8').read().strip()
        if k:
            return k
    k = ''.join(random.choice('abcdefghijkmnpqrstuvwxyz23456789') for _ in range(8))
    open(KEYFILE, 'w', encoding='utf-8').write(k)
    return k


KEY = get_key()


def lan_ip():
    ips = []
    try:
        ips = socket.gethostbyname_ex(socket.gethostname())[2]
    except Exception:
        pass
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('10.255.255.255', 1))
        ips.insert(0, s.getsockname()[0])
        s.close()
    except Exception:
        pass
    for ip in ips:
        if re.match(r'^(10\.|192\.168\.|172\.(1[6-9]|2\d|3[01])\.)', ip):
            return ip
    return ips[0] if ips else '127.0.0.1'


def cleanup_jobs():
    now = time.time()
    for d in os.listdir(JOBS):
        p = os.path.join(JOBS, d)
        try:
            if os.path.isdir(p) and now - os.path.getmtime(p) > KEEP_DAYS * 86400:
                shutil.rmtree(p, ignore_errors=True)
        except Exception:
            pass


def job_dir(job):
    if not re.fullmatch(r'[0-9a-f]{12}', job or ''):
        raise PlanError('작업 번호가 올바르지 않습니다.')
    d = os.path.join(JOBS, job)
    if not os.path.isdir(d):
        raise PlanError('작업이 만료되었습니다. 파일을 다시 올려 주세요.')
    return d


def rd(path, mode='rb'):
    with open(path, mode, **({} if 'b' in mode else {'encoding': 'utf-8'})) as f:
        return f.read()


def wr(path, data):
    with open(path, 'wb' if isinstance(data, bytes) else 'w', **({} if isinstance(data, bytes) else {'encoding': 'utf-8'})) as f:
        f.write(data)


def _dec(b):
    for enc in ('utf-8', 'cp949'):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            pass
    return b.decode('utf-8', 'replace')


def parse_multipart(ctype, raw):
    m = re.search(r'boundary="?([^";]+)"?', ctype)
    if not m:
        raise PlanError('업로드 형식이 올바르지 않습니다.')
    sep = b'--' + m.group(1).encode('latin-1')
    files, form = {}, {}
    for part in raw.split(sep):
        if not part or part.startswith(b'--'):
            continue
        part = part[2:] if part.startswith(b'\r\n') else part
        head, _, data = part.partition(b'\r\n\r\n')
        if data.endswith(b'\r\n'):
            data = data[:-2]
        hd = _dec(head)
        nm = re.search(r'\bname="([^"]*)"', hd)
        fn = re.search(r"filename\*=UTF-8''([^;\r\n]+)", hd, re.I)
        fname = urllib.parse.unquote(fn.group(1)) if fn else None
        if fname is None:
            fn = re.search(r'filename="([^"]*)"', hd)
            fname = fn.group(1) if fn else None
        if not nm:
            continue
        if fname is not None:
            files[nm.group(1)] = (fname.replace('\\', '/').split('/')[-1], data)
        else:
            form[nm.group(1)] = _dec(data)
    return files, form


def public_plan(plan, job, names):
    out = dict(plan)
    out['job'] = job
    out['names'] = names
    out['default_name'] = default_output_name(names)
    return out


def default_output_name(names):
    today = datetime.now().strftime('%y%m%d')
    m = re.match(r'^\s*(\([^)]*\))', names.get('notice', ''))
    if m:
        return f'{m.group(1)} 공급계약서(안)_{today}.docx'
    base = os.path.splitext(names.get('contract', '공급계약서'))[0]
    base = re.sub(r'_\d{6}$', '', base)
    return f'{base}_{today}.docx'


def safe_filename(name):
    name = re.sub(r'[\\/:*?"<>|]', '_', name).strip() or '공급계약서.docx'
    if not name.lower().endswith('.docx'):
        name += '.docx'
    return name


class Handler(BaseHTTPRequestHandler):
    server_version = 'ContractBuilder/1.0'

    def log_message(self, fmt, *args):
        log.info('%s %s', self.client_address[0], fmt % args)

    # ---------------------------------------------------------------- 공통
    def is_local(self):
        return self.client_address[0] in ('127.0.0.1', '::1')

    def authorized(self):
        if self.is_local():
            return True
        q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        if q.get('k', [''])[0] == KEY:
            return True
        if self.headers.get('X-Key') == KEY:
            return True
        cookie = self.headers.get('Cookie', '')
        return f'k={KEY}' in cookie

    def send(self, code, body, ctype='application/json; charset=utf-8', extra=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False).encode('utf-8')
        elif isinstance(body, str):
            body = body.encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def fail(self, msg, code=400):
        self.send(code, {'ok': False, 'error': msg})

    def body(self):
        n = int(self.headers.get('Content-Length') or 0)
        if n > MAX_UPLOAD:
            raise PlanError('파일이 너무 큽니다 (60MB 이하).')
        return self.rfile.read(n) if n else b''

    # ---------------------------------------------------------------- GET
    def do_GET(self):
        url = urllib.parse.urlparse(self.path)
        if not self.authorized():
            return self.send(403, '<meta charset="utf-8"><p style="font:16px sans-serif;padding:40px">'
                                  '접속키가 필요합니다. 담당자에게 받은 주소(…?k=…)로 들어와 주세요.</p>',
                             'text/html; charset=utf-8')
        if url.path in ('/', '/index.html'):
            extra = {}
            q = urllib.parse.parse_qs(url.query)
            if q.get('k', [''])[0] == KEY:
                extra['Set-Cookie'] = f'k={KEY}; Path=/; Max-Age=31536000; SameSite=Lax'
            return self.send(200, rd(os.path.join(APP, 'static', 'index.html')), 'text/html; charset=utf-8', extra)
        if url.path == '/api/ping':
            return self.send(200, {'ok': True, 'app': 'contract-builder', 'local': self.is_local()})
        if url.path == '/api/info':
            info = {'ok': True, 'local': self.is_local()}
            if self.is_local():
                info['share_url'] = f'http://{lan_ip()}:{PORT}/?k={KEY}'
            return self.send(200, info)
        if url.path == '/api/download':
            q = urllib.parse.parse_qs(url.query)
            try:
                d = job_dir(q.get('job', [''])[0])
            except PlanError as e:
                return self.fail(str(e), 404)
            kind = q.get('f', ['docx'])[0]
            meta = json.loads(rd(os.path.join(d, 'result.json'), 'r') if os.path.exists(os.path.join(d, 'result.json')) else '{}')
            if kind == 'xlsx':
                path, name = os.path.join(d, 'report.xlsx'), meta.get('report_name', '변경내역.xlsx')
                ctype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            else:
                path, name = os.path.join(d, 'result.docx'), meta.get('name', '공급계약서.docx')
                ctype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            if not os.path.exists(path):
                return self.fail('아직 만든 파일이 없습니다.', 404)
            disp = "attachment; filename*=UTF-8''" + urllib.parse.quote(name)
            return self.send(200, rd(path), ctype, {'Content-Disposition': disp})
        return self.fail('없는 주소입니다.', 404)

    # ---------------------------------------------------------------- POST
    def do_POST(self):
        url = urllib.parse.urlparse(self.path)
        if not self.authorized():
            return self.fail('접속키가 필요합니다.', 403)
        try:
            if url.path == '/api/analyze':
                return self.api_analyze()
            if url.path == '/api/reanalyze':
                return self.api_reanalyze()
            if url.path == '/api/build':
                return self.api_build()
            if url.path == '/api/shutdown' and self.is_local():
                self.send(200, {'ok': True})
                threading.Thread(target=self.server.shutdown, daemon=True).start()
                return
            return self.fail('없는 주소입니다.', 404)
        except PlanError as e:
            return self.fail(str(e))
        except Exception as e:
            log.error('error %s\n%s', e, traceback.format_exc())
            return self.fail(f'처리 중 오류가 났습니다: {e}', 500)

    def api_analyze(self):
        ctype = self.headers.get('Content-Type', '')
        if 'multipart/form-data' not in ctype:
            raise PlanError('파일을 올려 주세요.')
        raw = self.body()
        files, form = parse_multipart(ctype, raw)
        if 'contract' not in files or 'notice' not in files:
            raise PlanError('샘플 계약서와 모집공고 파일을 모두 올려 주세요.')
        cname, cdata = files['contract']
        nname, ndata = files['notice']
        if not cname.lower().endswith('.docx'):
            raise PlanError('샘플 계약서는 Word(.docx) 파일이어야 합니다.')
        if not nname.lower().endswith(('.hwp', '.hwpx')):
            raise PlanError('모집공고는 한글(.hwp) 파일이어야 합니다.')
        cleanup_jobs()
        job = os.urandom(6).hex()
        d = os.path.join(JOBS, job)
        os.makedirs(d)
        wr(os.path.join(d, 'contract.docx'), cdata)
        wr(os.path.join(d, 'notice.hwp'), ndata)
        names = {'contract': os.path.basename(cname), 'notice': os.path.basename(nname)}
        wr(os.path.join(d, 'names.json'), json.dumps(names, ensure_ascii=False))
        opts = {'strike': form.get('strike', '1') != '0'}
        t = time.time()
        try:
            plan = analyze(cdata, ndata, opts)
        except PlanError:
            shutil.rmtree(d, ignore_errors=True)
            raise
        wr(os.path.join(d, 'plan.json'), json.dumps(plan, ensure_ascii=False))
        log.info('analyze %s %s + %s (%.1fs) %s', job, names['contract'], names['notice'], time.time() - t, plan['summary'])
        self.send(200, {'ok': True, 'plan': public_plan(plan, job, names)})

    def api_reanalyze(self):
        req = json.loads(self.body() or b'{}')
        d = job_dir(req.get('job'))
        opts = {'strike': bool(req.get('strike', True))}
        plan = analyze(rd(os.path.join(d, 'contract.docx')), rd(os.path.join(d, 'notice.hwp')), opts)
        wr(os.path.join(d, 'plan.json'), json.dumps(plan, ensure_ascii=False))
        names = json.loads(rd(os.path.join(d, 'names.json'), 'r'))
        self.send(200, {'ok': True, 'plan': public_plan(plan, req['job'], names)})

    def api_build(self):
        req = json.loads(self.body() or b'{}')
        d = job_dir(req.get('job'))
        plan = json.loads(rd(os.path.join(d, 'plan.json'), 'r'))
        names = json.loads(rd(os.path.join(d, 'names.json'), 'r'))
        decisions = {k: bool(v) for k, v in (req.get('decisions') or {}).items()}
        fields = {}
        for f in plan['fields']:
            v = (req.get('fields') or {}).get(f['key'])
            if v is None:
                fields[f['key']] = f['new'] if f['apply'] else None
            else:
                fields[f['key']] = v.get('value') if v.get('apply') else None
        rules = req.get('rules')
        if rules is None:
            rules = plan['rules']
        cb, nb = rd(os.path.join(d, 'contract.docx')), rd(os.path.join(d, 'notice.hwp'))
        docx, blog = build(cb, nb, plan, decisions, fields, rules)
        name = safe_filename(req.get('filename') or default_output_name(names))
        wr(os.path.join(d, 'result.docx'), docx)
        wr(os.path.join(d, 'report.xlsx'), make_report(plan, decisions, fields, rules, blog, names))
        meta = {'name': name, 'report_name': os.path.splitext(name)[0] + '_변경내역.xlsx'}
        wr(os.path.join(d, 'result.json'), json.dumps(meta, ensure_ascii=False))
        log.info('build %s -> %s (%d steps)', req.get('job'), name, len(blog))
        self.send(200, {'ok': True, 'name': name, 'report_name': meta['report_name'], 'steps': len(blog)})


def main():
    local_only = '--local' in sys.argv
    host = '127.0.0.1' if local_only else '0.0.0.0'
    try:
        srv = ThreadingHTTPServer((host, PORT), Handler)
    except OSError:
        print(f'{PORT} 번 포트를 이미 쓰고 있습니다. 서버가 이미 켜져 있을 수 있습니다.')
        log.warning('port %s busy', PORT)
        sys.exit(2)
    srv.daemon_threads = True
    lines = [f'http://localhost:{PORT}/', '' if local_only else f'http://{lan_ip()}:{PORT}/?k={KEY}']
    wr(URLFILE, '\n'.join(lines))
    log.info('start %s', lines)
    try:
        print('공급계약서 자동작성 서버가 켜졌습니다.')
        print('  이 PC  :', lines[0])
        if lines[1]:
            print('  다른 PC:', lines[1])
    except Exception:
        pass
    try:
        srv.serve_forever()
    finally:
        try:
            os.remove(URLFILE)
        except OSError:
            pass
        log.info('stop')


if __name__ == '__main__':
    main()
