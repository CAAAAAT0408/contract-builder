# -*- coding: utf-8 -*-
"""사이트 버전·수정 이력 관리 — 최초 v1.0, 수정 1회 = +0.1. 자동화 모음 페이지가 version.json 을 읽어 버전과 수정사항을 보여준다.

  python bump_version.py <페이지키> "수정 내용" ["수정 내용 2" ...]   버전 0.1 올리고 수정 이력에 적음(제목 <title>·<h1> 의 _v1.N 도 갱신)
  python bump_version.py <페이지키> --set 1.0 "내용"                  버전을 직접 정함
  python bump_version.py --show                                      현재 버전·최근 수정사항 보기

version.json: {"pages": {"<페이지키>": {"name", "version": "1.0", "files": [...],
                                        "changelog": [{"version": "1.1", "date": "YYYY-MM-DD", "changes": ["..."]}, ...]}}}
"""
import datetime
import io
import json
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
VFILE = os.path.join(HERE, "version.json")
TAG = r"_v\d+(?:\.\d+)?"
SPAN = '<span class="ver" style="font-size:.6em;font-weight:normal;opacity:.75;margin-left:6px">_v{v}</span>'


def stamp(path, ver):
    with io.open(path, encoding="utf-8", newline="") as f:
        s = f.read()
    orig = s

    def title_sub(m):
        inner = re.sub(r"\s*" + TAG + r"\s*$", "", m.group(1).strip())
        return f"<title>{inner}_v{ver}</title>"

    s, n1 = re.subn(r"<title>(.*?)</title>", title_sub, s, count=1, flags=re.S)

    def h1_sub(m):
        inner = re.sub(r'\s*<span class="ver"[^>]*>.*?</span>\s*$', "", m.group(2), flags=re.S)
        inner = re.sub(r"\s*" + TAG + r"\s*$", "", inner)
        return f"<h1{m.group(1)}>{inner}{SPAN.format(v=ver)}</h1>"

    s, n2 = re.subn(r"<h1([^>]*)>(.*?)</h1>", h1_sub, s, count=1, flags=re.S)
    if s != orig:
        with io.open(path, "w", encoding="utf-8", newline="") as f:
            f.write(s)
    return n1, n2


def main():
    args = list(sys.argv[1:])
    cfg = json.load(io.open(VFILE, encoding="utf-8"))
    pages = cfg["pages"]
    if not args or "--show" in args:
        for k, p in pages.items():
            last = (p.get("changelog") or [{}])[0]
            print(f"{k:10s} v{p['version']}  {p['name']}  — 최근 {last.get('date', '-')}: {' / '.join(last.get('changes') or [])}")
        return
    key = args.pop(0)
    if key not in pages:
        sys.exit(f"페이지키가 없습니다: {key} (가능: {', '.join(pages)})")
    p = pages[key]
    setv = None
    if "--set" in args:
        i = args.index("--set")
        setv = args[i + 1]
        del args[i:i + 2]
    changes = [a for a in args if a.strip()]
    if not changes:
        sys.exit("수정 내용을 적어 주세요.  예) python bump_version.py budget \"세대수 칸 추가\"")
    new = setv if setv else f"{float(p['version']) + 0.1:.1f}"
    for rel in p["files"]:
        n1, n2 = stamp(os.path.join(HERE, rel), new)
        print(f"  {rel}: title {'O' if n1 else 'X'} h1 {'O' if n2 else 'X'}")
    p["version"] = new
    p.setdefault("changelog", [])
    today = datetime.date.today().isoformat()
    if p["changelog"] and p["changelog"][0].get("version") == new:
        p["changelog"][0]["date"] = today
        p["changelog"][0]["changes"] = changes
    else:
        p["changelog"].insert(0, dict(version=new, date=today, changes=changes))
    print(f"{key}: v{new}  {p['name']}  — {' / '.join(changes)}")
    with io.open(VFILE, "w", encoding="utf-8", newline="\n") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=1)
        f.write("\n")


if __name__ == "__main__":
    main()
