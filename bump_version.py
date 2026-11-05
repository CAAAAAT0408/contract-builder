# -*- coding: utf-8 -*-
"""사이트 제목의 버전 표기(_v1.NN) 관리 — 수정 1회 = +0.01, 최초 버전 v1.00.

  python bump_version.py                 모든 페이지 0.01 올림
  python bump_version.py <페이지키>       그 페이지만 0.01 올림
  python bump_version.py <페이지키> --set 1.05
  python bump_version.py --show          현재 버전만 보기

version.json: {"pages": {"<페이지키>": {"name": "...", "version": "1.06", "files": ["docs/index.html", ...]}}}
각 파일의 <title>…</title> 끝과 첫 <h1>…</h1> 끝에 '_v1.NN' 을 붙인다(이미 있으면 바꿈). 줄바꿈 방식은 그대로 둔다.
"""
import io
import json
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
VFILE = os.path.join(HERE, "version.json")
TAG = r"_v\d+\.\d+"
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
    args = [a for a in sys.argv[1:]]
    cfg = json.load(io.open(VFILE, encoding="utf-8"))
    pages = cfg["pages"]
    if "--show" in args:
        for k, p in pages.items():
            print(f"{k:10s} v{p['version']}  {p['name']}  ({', '.join(p['files'])})")
        return
    setv = None
    if "--set" in args:
        i = args.index("--set")
        setv = args[i + 1]
        del args[i:i + 2]
    keys = args or list(pages)
    for k in keys:
        p = pages[k]
        new = setv if setv else f"{float(p['version']) + 0.01:.2f}"
        for rel in p["files"]:
            n1, n2 = stamp(os.path.join(HERE, rel), new)
            print(f"  {rel}: title {'O' if n1 else 'X'} h1 {'O' if n2 else 'X'}")
        p["version"] = new
        print(f"{k}: v{new}  {p['name']}")
    with io.open(VFILE, "w", encoding="utf-8", newline="\n") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=1)
        f.write("\n")


if __name__ == "__main__":
    main()
