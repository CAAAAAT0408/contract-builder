"""브라우저(Pyodide)용 진입점. worker.js 가 부른다.

파일은 브라우저 메모리 안에서만 다루며 어디로도 보내지 않는다.
"""
import json
import os
import re
from datetime import datetime

from engine import PlanError, analyze as _analyze, build as _build
from report import make_report

STATE = {}


def _bytes(x):
    if hasattr(x, 'to_py'):
        x = x.to_py()
    return bytes(x)


def default_output_name(names):
    today = datetime.now().strftime('%y%m%d')
    m = re.match(r'^\s*(\([^)]*\))', names.get('notice', ''))
    if m:
        return f'{m.group(1)} 공급계약서(안)_{today}.docx'
    base = os.path.splitext(names.get('contract', '공급계약서'))[0]
    base = re.sub(r'_\d{6}$', '', base)
    return f'{base}_{today}.docx'


def safe_filename(name):
    name = re.sub(r'[\\/:*?"<>|]', '_', name or '').strip() or '공급계약서.docx'
    if not name.lower().endswith('.docx'):
        name += '.docx'
    return name


def _public(plan):
    out = dict(plan)
    out['names'] = STATE['names']
    out['default_name'] = default_output_name(STATE['names'])
    return json.dumps({'ok': True, 'plan': out}, ensure_ascii=False)


def analyze(cbytes, nbytes, cname, nname, strike):
    c, n = _bytes(cbytes), _bytes(nbytes)
    if not str(cname).lower().endswith('.docx'):
        return json.dumps({'ok': False, 'error': '샘플 계약서는 Word(.docx) 파일이어야 합니다.'}, ensure_ascii=False)
    if not str(nname).lower().endswith(('.hwp', '.hwpx')):
        return json.dumps({'ok': False, 'error': '모집공고는 한글(.hwp) 파일이어야 합니다.'}, ensure_ascii=False)
    try:
        plan = _analyze(c, n, {'strike': bool(strike)})
    except PlanError as e:
        return json.dumps({'ok': False, 'error': str(e)}, ensure_ascii=False)
    STATE.clear()
    STATE.update(c=c, n=n, plan=plan, names={'contract': str(cname), 'notice': str(nname)})
    return _public(plan)


def reanalyze(strike):
    if 'c' not in STATE:
        return json.dumps({'ok': False, 'error': '파일을 다시 넣어 주세요.'}, ensure_ascii=False)
    try:
        STATE['plan'] = _analyze(STATE['c'], STATE['n'], {'strike': bool(strike)})
    except PlanError as e:
        return json.dumps({'ok': False, 'error': str(e)}, ensure_ascii=False)
    return _public(STATE['plan'])


def build(req_json):
    if 'plan' not in STATE:
        return json.dumps({'ok': False, 'error': '먼저 분석을 해 주세요.'}, ensure_ascii=False)
    req = json.loads(req_json)
    plan = STATE['plan']
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
    try:
        docx, log = _build(STATE['c'], STATE['n'], plan, decisions, fields, rules)
    except PlanError as e:
        return json.dumps({'ok': False, 'error': str(e)}, ensure_ascii=False)
    name = safe_filename(req.get('filename') or default_output_name(STATE['names']))
    STATE['docx'] = docx
    STATE['xlsx'] = make_report(plan, decisions, fields, rules, log, STATE['names'])
    report_name = os.path.splitext(name)[0] + '_변경내역.xlsx'
    return json.dumps({'ok': True, 'name': name, 'report_name': report_name, 'steps': len(log)}, ensure_ascii=False)


def result(kind):
    return STATE.get('xlsx' if kind == 'xlsx' else 'docx', b'')


# ================================================================== 추가선택품목 계약서 (엑셀)
OPT = {}


def opt_default_name(names):
    today = datetime.now().strftime('%y%m%d')
    m = re.match(r'^\s*(\([^)]*\))', names.get('notice', ''))
    base = os.path.splitext(names.get('sample', '추가선택품목 계약서'))[0]
    core = re.sub(r'^\s*\([^)]*\)\s*', '', base)
    core = re.sub(r'_\d{6}.*$', '', core).strip() or '추가선택품목(유상옵션) 계약서(안)'
    return f'{m.group(1) + " " if m else ""}{core}_{today}.xlsx'


def opt_analyze(xbytes, nbytes, xname, nname):
    from options import OptionError, analyze_options
    from xlsxio import XlsxError
    x, n = _bytes(xbytes), _bytes(nbytes)
    if not str(xname).lower().endswith('.xlsx'):
        return json.dumps({'ok': False, 'error': '샘플 계약서는 엑셀(.xlsx) 파일이어야 합니다. .xls 이면 Excel 에서 .xlsx 로 저장해 주세요.'}, ensure_ascii=False)
    if not str(nname).lower().endswith(('.hwp', '.hwpx')):
        return json.dumps({'ok': False, 'error': '모집공고는 한글(.hwp) 파일이어야 합니다.'}, ensure_ascii=False)
    try:
        plan = analyze_options(x, n)
    except (OptionError, XlsxError) as e:
        return json.dumps({'ok': False, 'error': str(e)}, ensure_ascii=False)
    names = {'sample': str(xname), 'notice': str(nname)}
    OPT.clear()
    OPT.update(x=x, plan=plan, names=names)
    out = dict(plan)
    out['names'] = names
    out['default_name'] = opt_default_name(names)
    return json.dumps({'ok': True, 'plan': out}, ensure_ascii=False, default=list)


def opt_build(req_json):
    from options import OptionError, build_options
    from report import make_option_report
    from xlsxio import XlsxError
    if 'plan' not in OPT:
        return json.dumps({'ok': False, 'error': '먼저 분석을 해 주세요.'}, ensure_ascii=False)
    req = json.loads(req_json)
    plan = OPT['plan']
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
    balcony = {}
    for k, v in (req.get('balcony') or {}).items():
        val = re.sub(r'[^\d]', '', str(v.get('value') or ''))
        balcony[k] = {'value': int(val) if val else None, 'code': (v.get('code') or '').strip()}
    try:
        data, log = build_options(OPT['x'], plan, decisions, fields, rules, balcony, req.get('new_sheets') or {},
                                  bool(req.get('highlight')))
    except (OptionError, XlsxError) as e:
        return json.dumps({'ok': False, 'error': str(e)}, ensure_ascii=False)
    name = re.sub(r'[\\/:*?"<>|]', '_', req.get('filename') or opt_default_name(OPT['names'])).strip()
    if not name.lower().endswith('.xlsx'):
        name += '.xlsx'
    OPT['xlsx'] = data
    OPT['report'] = make_option_report(plan, decisions, log, OPT['names'])
    return json.dumps({'ok': True, 'name': name, 'report_name': os.path.splitext(name)[0] + '_변경내역.xlsx',
                       'steps': len(log)}, ensure_ascii=False)


def opt_result(kind):
    return OPT.get('report' if kind == 'report' else 'xlsx', b'')
