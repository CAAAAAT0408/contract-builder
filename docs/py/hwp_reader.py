"""한글(.hwp, HWP 5.0) 파일에서 문단 텍스트를 읽는다.

한컴오피스 없이 파일을 직접 해석한다. 각 문단은
  level : 레코드 깊이. 본문 표 안 = 4, 표 안의 표 = 6 처럼 표가 중첩될수록 커진다.
  text  : 문단 글자. 문단 머리 기호(■, - 등)는 앞에 붙이고, 줄바꿈(Shift+Enter)은 '⏎' 로 둔다.
"""
import io
import struct
import zlib

import olefile


class HwpError(Exception):
    pass


# 윙딩 등 기호 글꼴의 사적 영역 문자를 일반 기호로 바꾼다
BULLET_MAP = {'': '■', '': '■', '': '■', '': '●', '': '·', '': '·', '': '▶'}


def _records(data):
    i, n = 0, len(data)
    while i + 4 <= n:
        h = struct.unpack_from('<I', data, i); h = h[0]; i += 4
        tag, level, size = h & 0x3FF, (h >> 10) & 0x3FF, h >> 20
        if size == 0xFFF:
            size = struct.unpack_from('<I', data, i)[0]; i += 4
        yield tag, level, data[i:i + size]
        i += size


EXTENDED = (1, 2, 3, 11, 12, 14, 15, 16, 17, 18, 21, 22, 23)     # 표·그림 등 하위 레코드(CTRL_HEADER)가 딸린 컨트롤
INLINE = (4, 5, 6, 7, 8, 19, 20)


def _deleted(rec):
    """PARA_RANGE_TAG 에서 '변경 내용 추적'으로 지운 구간 [(시작, 끝)] (글자 단위).
    태그 상위 8비트가 종류: 0x11 지움, 0x10 넣음, 0x12·0x13 서식/문단 변경, 0x02 형광펜."""
    out = []
    for k in range(0, len(rec) - 11, 12):
        s, e, t = struct.unpack_from('<III', rec, k)
        if t >> 24 == 0x11:
            out.append((s, e))
    return out


def _para_text(rec, dele=()):
    """문단 글자와, 지운 구간(dele) 안에 든 확장 컨트롤의 순번 set 을 돌려준다. 지운 글자는 뺀다."""
    out, j, gone, n = [], 0, set(), 0
    while j + 1 < len(rec):
        c = struct.unpack_from('<H', rec, j)[0]
        cut = any(s <= j // 2 < e for s, e in dele)
        if c in EXTENDED or c in INLINE:
            if c in EXTENDED:
                if cut:
                    gone.add(n)
                n += 1
            j += 16; continue          # 확장/인라인 컨트롤 (8 글자 분량)
        j += 2
        if cut: continue
        if c == 9: out.append('\t')
        elif c == 10: out.append('⏎')
        elif c == 24: out.append('-')
        elif c in (30, 31): out.append(' ')
        elif c >= 32: out.append(chr(c))
    return ''.join(out), gone


def _body(data):
    """본문 레코드를 돌며 변경 추적으로 지운 글자·표를 뺀다.
    (tag, level, rec) 를 내되 PARA_TEXT(67) 는 rec 대신 해석한 글자(str)를 낸다.
    지운 문단(글자가 모두 지워짐)의 PARA_TEXT 는 None 이다."""
    recs = list(_records(data))
    ctrls = {}          # 문단 하위 level → [지운 컨트롤 순번, 지금까지 나온 컨트롤 수]
    skip = None
    for i, (tag, lvl, rec) in enumerate(recs):
        if skip is not None:
            if lvl > skip:
                continue
            skip = None
        if tag == 66:
            ctrls[lvl + 1] = [set(), 0]
        elif tag == 67:
            dele, k = [], i + 1
            while k < len(recs) and recs[k][0] in (68, 69, 70) and recs[k][1] == lvl:
                if recs[k][0] == 70:
                    dele = _deleted(recs[k][2])
                k += 1
            text, gone = _para_text(rec, dele)
            if lvl in ctrls:
                ctrls[lvl][0] = gone
            if dele and not text.replace('\r', '').strip():
                text = None
            yield tag, lvl, text
            continue
        elif tag == 71 and lvl in ctrls:
            info = ctrls[lvl]
            info[1] += 1
            if info[1] - 1 in info[0]:
                skip = lvl
                continue
        yield tag, lvl, rec


def read_hwp(data):
    """data: 파일 바이트. [{'level': int, 'text': str}, ...] 를 돌려준다."""
    if data[:4] == b'PK\x03\x04':
        raise HwpError('HWPX 형식입니다. 한글에서 "다른 이름으로 저장 → 한글 문서(*.hwp)" 로 저장한 뒤 넣어 주세요.')
    try:
        ole = olefile.OleFileIO(io.BytesIO(data))
    except Exception:
        raise HwpError('한글(.hwp) 파일이 아니거나 손상된 파일입니다.')
    hdr = ole.openstream('FileHeader').read()
    if not hdr.startswith(b'HWP Document File'):
        raise HwpError('한글(.hwp) 파일이 아닙니다.')
    flags = struct.unpack_from('<I', hdr, 36)[0]
    if flags & 0x2:
        raise HwpError('암호가 걸린 한글 문서입니다. 암호를 해제한 뒤 넣어 주세요.')
    if flags & 0x4:
        raise HwpError('배포용(보호) 한글 문서라 읽을 수 없습니다. 원본 .hwp 를 넣어 주세요.')
    compressed = bool(flags & 0x1)

    def stream(name):
        raw = ole.openstream(name).read()
        return zlib.decompress(raw, -15) if compressed else raw

    bullets, pshapes = [], []
    for tag, _lvl, rec in _records(stream('DocInfo')):
        if tag == 24:      # BULLET
            ch = rec[12:14].decode('utf-16le', errors='replace') if len(rec) >= 14 else '■'
            bullets.append(BULLET_MAP.get(ch, ch))
        elif tag == 25:    # PARA_SHAPE
            attr = struct.unpack_from('<I', rec, 0)[0]
            head = struct.unpack_from('<H', rec, 30)[0] if len(rec) >= 32 else 0
            pshapes.append(((attr >> 23) & 3, head))

    sections = sorted((e for e in ole.listdir() if e[0] == 'BodyText' and e[1].startswith('Section')),
                      key=lambda e: int(e[1][7:]))
    paras = []
    for sec in sections:
        for tag, lvl, rec in _body(stream('/'.join(sec))):
            if tag == 66:   # PARA_HEADER
                ps = struct.unpack_from('<H', rec, 8)[0] if len(rec) >= 10 else 0
                htype, hid = pshapes[ps] if ps < len(pshapes) else (0, 0)
                pre = ''
                if htype == 3 and 0 < hid <= len(bullets):
                    pre = bullets[hid - 1] + ' '
                paras.append({'level': lvl, 'text': pre, '_pre': pre})
            elif tag == 67 and paras:   # PARA_TEXT
                if rec is None:         # 변경 추적으로 통째 지운 문단
                    paras.pop()
                    continue
                paras[-1]['text'] = paras[-1]['_pre'] + rec.replace('\r', '')
    for p in paras:
        p.pop('_pre', None)
        for k, v in BULLET_MAP.items():
            if k in p['text']:
                p['text'] = p['text'].replace(k, v)
    return paras


def read_hwp_tables(data):
    """공고 안의 표를 칸 좌표까지 읽는다.
    [{'before': 표 바로 앞 문단, 'rows': n, 'cols': m,
      'cells': [{'r','c','rs','cs','t'}]}]   (t: 칸 글자, 줄바꿈은 '⏎')"""
    try:
        ole = olefile.OleFileIO(io.BytesIO(data))
    except Exception:
        raise HwpError('한글(.hwp) 파일이 아니거나 손상된 파일입니다.')
    flags = struct.unpack_from('<I', ole.openstream('FileHeader').read(), 36)[0]
    if flags & 0x6:
        raise HwpError('암호가 걸렸거나 배포용(보호) 한글 문서라 읽을 수 없습니다.')
    compressed = bool(flags & 0x1)
    sections = sorted((e for e in ole.listdir() if e[0] == 'BodyText' and e[1].startswith('Section')),
                      key=lambda e: int(e[1][7:]))
    tables, stack, last = [], [], ''
    for sec in sections:
        raw = ole.openstream('/'.join(sec)).read()
        body = zlib.decompress(raw, -15) if compressed else raw
        for tag, lvl, rec in _body(body):
            while stack and lvl <= stack[-1]['_lvl']:
                stack.pop()
            if tag == 71 and rec[:4] == b' lbt':            # 표 컨트롤
                t = {'_lvl': lvl, 'cells': [], 'before': last, 'rows': 0, 'cols': 0}
                tables.append(t)
                stack.append(t)
            elif tag == 77 and stack and len(rec) >= 8:      # TABLE: 행·열 수
                stack[-1]['rows'], stack[-1]['cols'] = struct.unpack_from('<HH', rec, 4)
            elif tag == 72 and stack and lvl == stack[-1]['_lvl'] + 1 and len(rec) >= 16:   # 셀
                col, row, cs, rs = struct.unpack_from('<HHHH', rec, 8)
                stack[-1]['cells'].append({'r': row, 'c': col, 'rs': rs, 'cs': cs, 't': []})
            elif tag == 67 and rec is not None:
                txt = rec.replace('\r', '')
                for k, v in BULLET_MAP.items():
                    txt = txt.replace(k, v)
                if stack and stack[-1]['cells']:
                    stack[-1]['cells'][-1]['t'].append(txt)
                if txt.strip():
                    last = txt.strip()
    for t in tables:
        t.pop('_lvl', None)
        for c in t['cells']:
            c['t'] = '⏎'.join(x for x in c['t'] if x.strip())
    return tables


def table_grid(t):
    """병합을 풀어 [행][열] = (글자, 원래칸번호) 격자로 만든다."""
    g = [[('', -1)] * t['cols'] for _ in range(t['rows'])]
    for k, c in enumerate(t['cells']):
        for r in range(c['r'], min(c['r'] + max(c['rs'], 1), t['rows'])):
            for cc in range(c['c'], min(c['c'] + max(c['cs'], 1), t['cols'])):
                g[r][cc] = (c['t'], k)
    return g
