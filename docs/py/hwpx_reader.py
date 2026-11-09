"""한글 HWPX(.hwpx, OWPML) 파일에서 문단 텍스트와 표를 읽는다.

hwp_reader.read_hwp / read_hwp_tables 와 같은 모양으로 돌려주므로 notice.py·options.py·engine.py 는 그대로 쓴다.
  문단 : {'level': int, 'text': str}   본문 = 0, 표 안 = 2, 표 안의 표 = 4 … (hwp_reader 와 같은 체계)
  표   : {'before', 'rows', 'cols', 'cells': [{'r','c','rs','cs','t'}]}
HWPX 는 zip 안의 XML(Contents/header.xml, Contents/section0.xml …)이므로 표준 라이브러리로만 읽는다.
한컴오피스 없이 파일을 직접 해석한다.
"""
import io
import re
import zipfile
import xml.etree.ElementTree as ET

from hwp_reader import BULLET_MAP, HwpError

STEP = 2             # 표·글상자 한 겹당 level 증가폭 (hwp 는 본문 0, 표 안 2, 표 안의 표 4)
SKIP_MEMO = True     # 메모(검토 의견) 글자는 공고 본문이 아니므로 뺀다
MEMO_FIELD_TYPES = ('MEMO',)


def is_hwpx(data):
    return data[:4] == b'PK\x03\x04'


def _local(tag):
    return tag.rsplit('}', 1)[-1]


def _iter(el, name):
    """이름공간을 무시하고 자손 요소를 찾는다 (ElementTree.iter 는 '{*}' 를 지원하지 않는다)."""
    return (e for e in el.iter() if _local(e.tag) == name)


def _child(el, name):
    return next((c for c in el if _local(c.tag) == name), None)


def _open(data):
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
        names = set(z.namelist())
    except Exception:
        raise HwpError('한글(.hwpx) 파일이 아니거나 손상된 파일입니다.')
    if 'Contents/header.xml' not in names:
        raise HwpError('한글(.hwpx) 파일이 아닙니다.')
    try:
        header = ET.fromstring(z.read('Contents/header.xml'))
    except Exception:
        if 'META-INF/manifest.xml' in names and b'encrypt' in z.read('META-INF/manifest.xml').lower():
            raise HwpError('암호가 걸린 한글 문서입니다. 암호를 해제한 뒤 넣어 주세요.')
        raise HwpError('한글(.hwpx) 파일이 손상되어 읽을 수 없습니다.')
    return z, header


def _bullet_heads(header):
    """paraPr id → 글머리표 접두어('■ ' 등). hwp_reader 처럼 글머리표(BULLET)만 붙이고 번호 매기기는 안 붙인다."""
    bullets = {}
    for b in _iter(header, 'bullet'):
        ch = b.get('char') or '■'
        bullets[b.get('id')] = BULLET_MAP.get(ch, ch)
    heads = {}
    for pp in _iter(header, 'paraPr'):
        h = _child(pp, 'heading')
        if h is not None and h.get('type') == 'BULLET' and h.get('idRef') in bullets:
            heads[pp.get('id')] = bullets[h.get('idRef')] + ' '
    return heads


def _sections(z):
    secs = [n for n in z.namelist() if re.fullmatch(r'Contents/section\d+\.xml', n)]
    return sorted(secs, key=lambda n: int(re.findall(r'\d+', n)[0]))


def _t_text(t, st):
    """<hp:t> 의 글자. 줄바꿈→'⏎', 탭, 특수 빈칸, 하이픈을 바꾸고 변경 추적으로 지운 구간은 뺀다."""
    out = []
    if not st['del']:
        out.append(t.text or '')
    for c in t:
        n = _local(c.tag)
        if n == 'deleteBegin':
            st['del'] += 1
            st['touched'] = True
        elif n == 'deleteEnd':
            st['del'] = max(0, st['del'] - 1)
        elif not st['del']:
            if n == 'lineBreak':
                out.append('⏎')
            elif n == 'tab':
                out.append('\t')
            elif n in ('nbSpace', 'fwSpace'):
                out.append(' ')
            elif n == 'hyphen':
                out.append('-')
        if not st['del']:
            out.append(c.tail or '')
    return ''.join(out)


def _inner(el):
    """el 안에서 가장 바깥의 표(tbl)·문단 목록(subList: 글상자, 머리말·꼬리말, 각주, 메모 등)을 문서 순서대로."""
    out = []
    for c in el:
        n = _local(c.tag)
        if n in ('tbl', 'subList'):
            out.append(c)
        elif n == 'fieldBegin' and SKIP_MEMO and c.get('type') in MEMO_FIELD_TYPES:
            continue
        else:
            out.extend(_inner(c))
    return out


def _walk(container, level, cell, st):
    """container 바로 아래의 <hp:p> 들을 차례로 읽는다.
    hwp 와 같이 문단 글자를 먼저 내고, 그 문단에 든 표·글상자의 문단들을 뒤에 낸다."""
    for p in container:
        if _local(p.tag) != 'p':
            continue
        st['touched'] = st['del'] > 0
        text, nested = [], []
        for run in p:
            if _local(run.tag) != 'run':
                continue
            for x in run:
                n = _local(x.tag)
                if n == 't':
                    text.append(_t_text(x, st))
                elif n == 'deleteBegin':
                    st['del'] += 1
                    st['touched'] = True
                elif n == 'deleteEnd':
                    st['del'] = max(0, st['del'] - 1)
                elif n == 'secPr':
                    continue
                elif not st['del']:                      # 지운 구간 안의 표·글상자는 뺀다
                    if n == 'tbl':
                        nested.append(x)
                    else:
                        nested.extend(_inner(x))
        txt = ''.join(text)
        for k, v in BULLET_MAP.items():
            txt = txt.replace(k, v)
        if not (st['touched'] and not txt.strip()):      # 변경 추적으로 통째 지운 문단은 뺀다
            st['paras'].append({'level': level, 'text': st['heads'].get(p.get('paraPrIDRef'), '') + txt})
            if cell is not None:
                cell['t'].append(txt)
            if txt.strip():
                st['last'] = txt.strip()
        for el in nested:
            if _local(el.tag) == 'tbl':
                _table(el, level, st)
            else:
                _walk(el, level + STEP, None, st)


def _table(tbl, level, st):
    t = {'before': st['last'], 'rows': int(tbl.get('rowCnt') or 0), 'cols': int(tbl.get('colCnt') or 0), 'cells': []}
    st['tables'].append(t)
    for tr in tbl:
        if _local(tr.tag) != 'tr':
            continue
        for tc in tr:
            if _local(tc.tag) != 'tc':
                continue
            addr, span = _child(tc, 'cellAddr'), _child(tc, 'cellSpan')
            cell = {'r': int(addr.get('rowAddr', 0)) if addr is not None else 0,
                    'c': int(addr.get('colAddr', 0)) if addr is not None else 0,
                    'rs': int(span.get('rowSpan', 1)) if span is not None else 1,
                    'cs': int(span.get('colSpan', 1)) if span is not None else 1, 't': []}
            t['cells'].append(cell)
            sl = _child(tc, 'subList')
            if sl is not None:
                _walk(sl, level + STEP, cell, st)


_CACHE = {}     # 같은 파일을 read_hwpx → read_hwpx_tables 로 연달아 읽을 때 두 번 풀지 않도록 (한 개만 기억)


def _parse(data):
    key = (len(data), hash(data))
    if _CACHE.get('key') == key:
        return _CACHE['val']
    z, header = _open(data)
    st = {'heads': _bullet_heads(header), 'paras': [], 'tables': [], 'last': '', 'del': 0, 'touched': False}
    for name in _sections(z):
        _walk(ET.fromstring(z.read(name)), 0, None, st)
    for t in st['tables']:
        for c in t['cells']:
            c['t'] = '⏎'.join(x for x in c['t'] if x.strip())
    _CACHE.clear()
    _CACHE.update(key=key, val=(st['paras'], st['tables']))
    return _CACHE['val']


def read_hwpx(data):
    """data: 파일 바이트. [{'level': int, 'text': str}, ...] 를 돌려준다."""
    return _parse(data)[0]


def read_hwpx_tables(data):
    """공고 안의 표를 칸 좌표까지 읽는다. hwp_reader.read_hwp_tables 와 같은 모양."""
    return _parse(data)[1]
