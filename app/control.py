"""서버 켜기/끄기/상태/자동시작. 폴더의 .bat 파일들이 이 스크립트를 부른다.

  python control.py open        서버가 꺼져 있으면 창 없이 켜고 브라우저를 연다
  python control.py start | stop | status
  python control.py autostart-on | autostart-off
"""
import json
import os
import subprocess
import sys
import time
import urllib.request
import webbrowser

APP = os.path.dirname(os.path.abspath(__file__))
PORT = 8790
URLFILE = os.path.join(APP, '_url.txt')
LOGFILE = os.path.join(APP, '_server.log')
STARTUP = os.path.join(os.environ.get('APPDATA', ''), r'Microsoft\Windows\Start Menu\Programs\Startup')
VBS = os.path.join(STARTUP, '공급계약서 자동작성 서버.vbs')

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass


def ping():
    try:
        with urllib.request.urlopen(f'http://127.0.0.1:{PORT}/api/ping', timeout=2) as r:
            return json.loads(r.read()).get('app') == 'contract-builder'
    except Exception:
        return False


def pythonw():
    exe = os.path.join(os.path.dirname(sys.executable), 'pythonw.exe')
    return exe if os.path.exists(exe) else sys.executable


def start():
    if ping():
        return 'already'
    flags = 0x00000008 | 0x00000200 | 0x08000000   # DETACHED_PROCESS | NEW_PROCESS_GROUP | NO_WINDOW
    subprocess.Popen([pythonw(), os.path.join(APP, 'server.py')], cwd=APP, creationflags=flags,
                     stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, close_fds=True)
    for _ in range(60):
        time.sleep(0.25)
        if ping():
            return 'started'
    return 'failed'


def stop():
    if not ping():
        return False
    try:
        req = urllib.request.Request(f'http://127.0.0.1:{PORT}/api/shutdown', data=b'{}', method='POST')
        urllib.request.urlopen(req, timeout=3).read()
    except Exception:
        pass
    for _ in range(20):
        time.sleep(0.25)
        if not ping():
            return True
    return False


def urls():
    if os.path.exists(URLFILE):
        lines = open(URLFILE, encoding='utf-8').read().splitlines() + ['', '']
        return lines[0], lines[1]
    return f'http://localhost:{PORT}/', ''


def show_start(r):
    if r == 'already':
        print('  서버는 이미 켜져 있습니다.')
    elif r == 'started':
        print('  서버를 켰습니다.')
    else:
        print('  서버가 켜지지 않았습니다. 로그를 확인하세요:', LOGFILE)
        if os.path.exists(LOGFILE):
            for line in open(LOGFILE, encoding='utf-8', errors='replace').read().splitlines()[-12:]:
                print('    ' + line)


def show_urls():
    local, lan = urls()
    print('  이 PC  :', local)
    if lan:
        print('  다른 PC:', lan, ' ← 팀원에게 이 주소를 알려 주세요')


def main():
    cmd = sys.argv[1] if len(sys.argv) > 1 else 'status'
    print()
    if cmd == 'open':
        r = start()
        show_start(r)
        if r != 'failed':
            show_urls()
            webbrowser.open(f'http://localhost:{PORT}/')
        return 0 if r != 'failed' else 1
    if cmd == 'start':
        r = start(); show_start(r)
        if r != 'failed':
            show_urls()
        return 0 if r != 'failed' else 1
    if cmd == 'stop':
        print('  서버를 껐습니다.' if stop() else '  서버가 켜져 있지 않습니다.')
        return 0
    if cmd == 'status':
        if ping():
            print('  서버가 켜져 있습니다.')
            show_urls()
        else:
            print('  서버가 꺼져 있습니다. "페이지 열기.bat" 을 실행하면 켜집니다.')
        print('  자동시작:', '등록됨' if os.path.exists(VBS) else '등록 안 됨')
        return 0
    if cmd == 'autostart-on':
        os.makedirs(STARTUP, exist_ok=True)
        vbs = ('Set sh = CreateObject("WScript.Shell")\r\n'
               f'sh.CurrentDirectory = "{APP}"\r\n'
               f'sh.Run """{pythonw()}"" ""{os.path.join(APP, "server.py")}""", 0, False\r\n')
        with open(VBS, 'w', encoding='utf-16') as f:
            f.write(vbs)
        print('  로그인할 때마다 서버가 저절로 켜지도록 등록했습니다.')
        show_start(start())
        show_urls()
        return 0
    if cmd == 'autostart-off':
        if os.path.exists(VBS):
            os.remove(VBS)
        print('  자동시작을 해제했습니다.')
        print('  서버를 껐습니다.' if stop() else '  (서버는 이미 꺼져 있었습니다.)')
        return 0
    print('알 수 없는 명령:', cmd)
    return 1


if __name__ == '__main__':
    sys.exit(main())
