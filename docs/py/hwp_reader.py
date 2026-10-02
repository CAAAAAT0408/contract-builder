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


def _para_text(rec):
    out, j = [], 0
    while j + 1 < len(rec):
        c = struct.unpack_from('<H', rec, j)[0]
        if c in (1, 2, 3, 4, 5, 6, 7, 8, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23):
            j += 16; continue          # 확장/인라인 컨트롤 (8 글자 분량)
        if c == 9: out.append('\t')
        elif c == 10: out.append('⏎')
        elif c == 24: out.append('-')
        elif c in (30, 31): out.append(' ')
        elif c >= 32: out.append(chr(c))
        j += 2
    return ''.join(out)


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
        for tag, lvl, rec in _records(stream('/'.join(sec))):
            if tag == 66:   # PARA_HEADER
                ps = struct.unpack_from('<H', rec, 8)[0] if len(rec) >= 10 else 0
                htype, hid = pshapes[ps] if ps < len(pshapes) else (0, 0)
                pre = ''
                if htype == 3 and 0 < hid <= len(bullets):
                    pre = bullets[hid - 1] + ' '
                paras.append({'level': lvl, 'text': pre, '_pre': pre})
            elif tag == 67 and paras:   # PARA_TEXT
                paras[-1]['text'] = paras[-1]['_pre'] + _para_text(rec).replace('\r', '')
    for p in paras:
        p.pop('_pre', None)
        for k, v in BULLET_MAP.items():
            if k in p['text']:
                p['text'] = p['text'].replace(k, v)
    return paras
