"""샘플 공급계약서(.docx) 읽기·고치기. 서식(글꼴, 색, 표)은 그대로 두고 글자만 바꾼다."""
import copy
import io
import re
import zipfile
from collections import Counter

from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
TERMINAL = ('.', '함', '음', '람', ')', '임', '됨', '요')


def w(tag):
    return '{%s}%s' % (W, tag)


class ContractError(Exception):
    pass


def nospace(s):
    return re.sub(r'\s+', '', s)


def ptext(el):
    return ''.join(t.text or '' for t in el.iter(w('t')))


def has_drawing(p):
    for e in p.iter():
        tag = e.tag if isinstance(e.tag, str) else ''
        if tag.endswith('}drawing') or tag.endswith('}AlternateContent') or tag.endswith('}pict'):
            return True
    return False


# ------------------------------------------------------------------ 글자 편집 도구
RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline',
             'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
             'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText',
             'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath']


def clean_rpr(rpr, hl):
    rpr = copy.deepcopy(rpr) if rpr is not None else etree.Element(w('rPr'))
    for tag in ('strike', 'dstrike', 'highlight', 'rPrChange', 'ins', 'del'):
        for e in rpr.findall(w(tag)):
            rpr.remove(e)
    if hl:
        h = etree.Element(w('highlight'))
        h.set(w('val'), 'yellow')
        after = RPR_ORDER[RPR_ORDER.index('highlight') + 1:]
        for i, c in enumerate(rpr):
            if etree.QName(c).localname in after:
                rpr.insert(i, h)
                break
        else:
            rpr.append(h)
    return rpr


def make_run(rpr, text, hl):
    r = etree.Element(w('r'))
    r.append(clean_rpr(rpr, hl))
    for k, part in enumerate(text.split('⏎')):
        if k:
            r.append(etree.Element(w('br')))
        if part:
            t = etree.SubElement(r, w('t'))
            t.text = part
            t.set(XML_SPACE, 'preserve')
    return r


def make_para(ppr_src, segments, rpr):
    """ppr_src 문단의 단락서식을 빌려 새 문단을 만든다. segments = [(글자, 형광펜여부), ...]"""
    p = etree.Element(w('p'))
    ppr = ppr_src.find(w('pPr')) if ppr_src is not None else None
    if ppr is not None:
        ppr = copy.deepcopy(ppr)
        for e in list(ppr.iter(w('ins'), w('del'), w('rPrChange'), w('pPrChange'), w('sectPr'))):
            e.getparent().remove(e)
        p.append(ppr)
    for text, hl in segments:
        if text:
            p.append(make_run(rpr, text, hl))
    return p


def explode_runs(p):
    """한 run 안에 글자·탭·줄바꿈이 섞여 있으면 하나씩 나눈다 (글자 위치 계산을 쉽게 하려고)."""
    for r in list(p.iter(w('r'))):
        content = [c for c in r if c.tag != w('rPr')]
        if len(content) <= 1:
            continue
        rpr = r.find(w('rPr'))
        anchor = r
        for c in content:
            nr = etree.Element(w('r'))
            if rpr is not None:
                nr.append(copy.deepcopy(rpr))
            nr.append(c)
            anchor.addnext(nr)
            anchor = nr
        r.getparent().remove(r)


def edit_range(p, i, j, new, hl=True):
    """문단 글자(ptext 기준) i:j 를 new 로 바꾼다. 바뀐 글자에는 형광펜을 칠한다."""
    explode_runs(p)
    ts = list(p.iter(w('t')))
    pos, first = 0, True
    if i == j:   # 끼워넣기: i 위치를 포함하는 글자 뒤에
        for t in ts:
            s = t.text or ''
            a, b = pos, pos + len(s)
            pos = b
            if a <= i <= b and (i < b or t is ts[-1] or b == a):
                r = t.getparent()
                pre, post = s[:i - a], s[i - a:]
                t.text = pre
                t.set(XML_SPACE, 'preserve')
                nr = make_run(r.find(w('rPr')), new, hl)
                r.addnext(nr)
                if post:
                    pr = copy.deepcopy(r)
                    pt = pr.find(w('t'))
                    pt.text = post
                    pt.set(XML_SPACE, 'preserve')
                    nr.addnext(pr)
                if not pre:
                    r.getparent().remove(r)
                return
        # 빈 문단
        ppr = p.find(w('pPr'))
        rpr = ppr.find(w('rPr')) if ppr is not None else None
        p.append(make_run(rpr, new, hl))
        return
    for t in ts:
        s = t.text or ''
        a, b = pos, pos + len(s)
        pos = b
        if b <= i or a >= j or a == b:
            continue
        r = t.getparent()
        pre = s[:max(0, i - a)]
        post = s[j - a:] if j < b else ''
        if first:
            first = False
            t.text = pre
            t.set(XML_SPACE, 'preserve')
            anchor = r
            if new:
                nr = make_run(r.find(w('rPr')), new, hl)
                anchor.addnext(nr)
                anchor = nr
            if post:
                pr = copy.deepcopy(r)
                pt = pr.find(w('t'))
                pt.text = post
                pt.set(XML_SPACE, 'preserve')
                anchor.addnext(pr)
            if not pre:
                r.getparent().remove(r)
        else:
            if post:
                t.text = post
                t.set(XML_SPACE, 'preserve')
            else:
                r.getparent().remove(r)


def replace_all(p, old, new, hl=True):
    """문단 안의 old 를 모두 new 로. 바꾼 횟수를 돌려준다."""
    n = 0
    while True:
        txt = ptext(p)
        k = txt.find(old)
        if k < 0 or n > 50:
            return n
        edit_range(p, k, k + len(old), new, hl)
        n += 1
        if old in new:   # 무한반복 방지
            return n


# ------------------------------------------------------------------ 계약서 본체
class Unit:
    """계약서 20조 안의 '항목' 하나 (■ 문단 + 엔터로 끊긴 이어지는 문단들)."""
    def __init__(self, kind, paras, section, major, dong):
        self.kind, self.paras = kind, paras
        self.section, self.major, self.dong = section, major, dong

    @property
    def text(self):
        out = ''
        for k, p in enumerate(self.paras):
            t = ptext(p)
            out = t if k == 0 else out + t.lstrip()
        return out

    def merge_paras(self):
        """엔터로 끊긴 문단들을 첫 문단 하나로 합친다 (글자는 그대로)."""
        p0 = self.paras[0]
        for q in self.paras[1:]:
            explode_runs(q)
            first = True
            for r in [r for r in q if r.tag == w('r')]:
                t = r.find(w('t'))
                if first and t is not None:
                    t.text = (t.text or '').lstrip()
                    t.set(XML_SPACE, 'preserve')
                    if not t.text:
                        continue
                    first = False
                p0.append(r)
            q.getparent().remove(q)
        self.paras = [p0]
        return p0


class Contract:
    def __init__(self, data):
        try:
            self.zin = zipfile.ZipFile(io.BytesIO(data))
            xml = self.zin.read('word/document.xml')
        except Exception:
            raise ContractError('Word(.docx) 파일이 아닙니다. .doc 이면 Word 에서 .docx 로 저장해 주세요.')
        self.tree = etree.fromstring(xml)
        self.body = self.tree.find(w('body'))

    def save(self):
        out = io.BytesIO()
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
            for info in self.zin.infolist():
                if info.filename == 'word/document.xml':
                    z.writestr(info, etree.tostring(self.tree, xml_declaration=True, encoding='UTF-8', standalone=True))
                else:
                    z.writestr(info, self.zin.read(info.filename))
        return out.getvalue()

    # -------------------------------------------------------------- 변경추적 수락
    def accept_revisions(self):
        root = self.tree
        n = 0
        for tag in ('moveFrom', 'del'):
            for e in list(root.iter(w(tag))):
                if e.getparent().tag in (w('rPr'), w('pPr')):
                    continue
                e.getparent().remove(e); n += 1
        for tag in ('moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd',
                    'rPrChange', 'pPrChange'):
            for e in list(root.iter(w(tag))):
                e.getparent().remove(e); n += 1
        for tag in ('ins', 'moveTo'):
            for e in list(root.iter(w(tag))):
                par = e.getparent()
                if par.tag == w('rPr'):
                    par.remove(e); continue
                idx = par.index(e)
                for c in list(e):
                    par.insert(idx, c); idx += 1
                par.remove(e); n += 1
        # 문단 기호가 삭제된 문단: 다음 문단과 합친다
        for e in list(root.iter(w('del'))):
            rpr = e.getparent()
            ppr = rpr.getparent() if rpr is not None else None
            p = ppr.getparent() if ppr is not None else None
            rpr.remove(e)
            if p is None or p.tag != w('p') or ppr.tag != w('pPr'):
                continue
            nxt = p.getnext()
            if nxt is not None and nxt.tag == w('p'):
                runs = [c for c in p if c.tag == w('r')]
                if ptext(p).strip():
                    first = nxt.find(w('pPr'))
                    pos = 1 if first is not None else 0
                    for r in runs:
                        nxt.insert(pos, r); pos += 1
                p.getparent().remove(p)
            n += 1
        return n

    # -------------------------------------------------------------- 1페이지
    def tables(self):
        return [c for c in self.body if c.tag == w('tbl')]

    @staticmethod
    def cells(tr):
        return tr.findall(w('tc'))

    def property_cell(self):
        for tbl in self.tables():
            for tr in tbl.findall(w('tr')):
                tcs = self.cells(tr)
                if len(tcs) >= 2 and nospace(ptext(tcs[0])) == '재산의표시':
                    return tcs[1]
        return None

    def movein_cell(self):
        for tbl in self.tables():
            for tr in tbl.findall(w('tr')):
                tcs = self.cells(tr)
                if len(tcs) >= 2 and nospace(ptext(tcs[0])).startswith('입주예정'):
                    return tcs[1]
        return None

    @staticmethod
    def _grid(tr):
        out, col = [], 0
        for tc in tr.findall(w('tc')):
            span = 1
            tcpr = tc.find(w('tcPr'))
            if tcpr is not None and tcpr.find(w('gridSpan')) is not None:
                span = int(tcpr.find(w('gridSpan')).get(w('val')))
            out.append((col, tc))
            col += span
        return out

    def payment_slots(self):
        """공급대금 표에서 '약정일자' 칸들. [(라벨, 칸, 현재글자), ...] 계약체결시 칸은 빼고 돌려준다."""
        for tbl in self.tables():
            rows = tbl.findall(w('tr'))
            for hi, tr in enumerate(rows):
                grid = self._grid(tr)
                col = next((c for c, tc in grid if nospace(ptext(tc)) == '약정일자'), None)
                if col is None:
                    continue
                slots, group = [], ''
                for tr2 in rows[hi + 1:]:
                    g = self._grid(tr2)
                    labels = [nospace(ptext(tc)) for c, tc in g if c < col]
                    if labels and labels[0] and '회차' not in labels[0]:
                        group = labels[0]
                    turn = next((x for x in labels if '회차' in x), '')
                    cell = next((tc for c, tc in g if c == col), None)
                    if not turn or cell is None:
                        continue
                    cur = ptext(cell).strip()
                    if '계약' in nospace(cur):
                        continue
                    slots.append((f'{group} {turn}'.strip(), cell, cur))
                return slots
        return []

    def account_para(self):
        for p in self.body.iter(w('p')):
            t = ptext(p)
            if '납부계좌' in t or '예금주' in t:
                m = re.search(r'([가-힣A-Za-z]*(?:은행|농협|수협|금고|신협))\s*([\d]{2,}[\d\-]+\d)', t)
                if m:
                    h = re.search(r'예금주\s*[:：]\s*([^\],]+)', t)
                    return p, m.group(1), m.group(2), (h.group(1).strip() if h else '')
        return None

    def old_identity(self):
        cell = self.property_cell()
        if cell is None:
            return {}
        t = ptext(cell)
        out = {'property_text': re.sub(r'\s{2,}', '  ', t.strip())}
        m = re.search(r'([A-Za-z]{1,4}-?\d{1,3})\s*(BL|블록|블럭)', t)
        if m:
            out['block'] = m.group(1)
            rest = t[m.end():]
            name = re.split(r'\s{2,}|동\s', rest.strip())[0].strip()
            out['name'] = name
        return out

    # -------------------------------------------------------------- 20조 영역
    def region_bounds(self):
        kids = list(self.body)
        start = None
        for i, c in enumerate(kids):
            if c.tag == w('p') and re.match(r'^제\d+조\(기타사항\)', nospace(ptext(c))):
                start = i
        if start is None:
            raise ContractError("계약서에서 '제○조 (기타사항)' 조항을 찾지 못했습니다.")
        end = len(kids)
        for i in range(start + 1, len(kids)):
            c = kids[i]
            if c.tag == w('tbl') and '승계' in ptext(c):
                end = i; break
            if c.tag == w('sectPr'):
                end = i; break
        return start, end

    def apply_strikes(self):
        """20조 안의 취소선 글자(검토자가 '삭제' 표시한 것)를 실제로 지운다."""
        start, end = self.region_bounds()
        kids = list(self.body)[start + 1:end]
        removed_paras, snippets = [], []
        for p in kids:
            if p.tag != w('p') or has_drawing(p):
                continue
            explode_runs(p)
            runs = [r for r in p if r.tag == w('r') and r.find(w('t')) is not None]
            if not runs:
                continue

            def struck(r):
                s = r.find(w('rPr'))
                s = s.find(w('strike')) if s is not None else None
                return s is not None and s.get(w('val'), 'true') not in ('0', 'false')
            flags = [struck(r) for r in runs]
            if not any(flags):
                continue
            full = ptext(p)
            if all(flags) or not ''.join(r.find(w('t')).text or '' for r, f in zip(runs, flags) if not f).strip():
                removed_paras.append(full)
                p.getparent().remove(p)
                continue
            pos = 0
            for r, f in zip(runs, flags):
                s = r.find(w('t')).text or ''
                if f:
                    left = full[max(0, pos - 2):pos]
                    snippets.append({'left': left, 'text': s, 'para': full})
                    p.remove(r)
                pos += len(s)
        # '(삭제)' 같은 표시용 글자와 공백뿐인 조각은 규칙에서 뺀다
        rules = []
        for s in snippets:
            core = re.sub(r'\(?\s*삭제\s*\)?', '', s['text']).strip()
            if len(core) >= 2 and s['left']:
                rules.append({'left': s['left'], 'text': core, 'para': s['para']})
        return removed_paras, rules

    def region_units(self):
        start, end = self.region_bounds()
        kids = list(self.body)[start + 1:end]
        units, blocks = [], []    # blocks: 영역 안 모든 요소를 순서대로 (kind, element, unit)
        section, major, dong = '', '', None
        prev_unit = None
        for c in kids:
            if c.tag == w('tbl'):
                blocks.append(('table', c, None)); prev_unit = None; continue
            if c.tag != w('p'):
                continue
            if has_drawing(c):
                blocks.append(('fixed', c, None)); prev_unit = None; continue
            t = ptext(c)
            s = t.strip()
            if not s:
                blocks.append(('blank', c, None)); prev_unit = None; continue
            if s.startswith('◎'):
                major, section, dong = s, '', None
                blocks.append(('major', c, None)); prev_unit = None; continue
            if re.fullmatch(r'<[^<>]{1,30}>', s):
                section, dong = s, None
                blocks.append(('header', c, None)); prev_unit = None; continue
            if re.fullmatch(r'\d{3,4}\s*동', s):
                dong = nospace(s)
                blocks.append(('dong', c, None)); prev_unit = None; continue
            if s.startswith('■') or s.startswith('※'):
                kind = 'item'
            elif t[:1].isspace():
                if s.startswith('-') or re.match(r'\d\)', s):
                    kind = 'sub'
                elif s.startswith('(') and prev_unit is not None and prev_unit.text.rstrip().endswith(TERMINAL):
                    kind = 'sub'
                elif prev_unit is not None:
                    prev_unit.paras.append(c)
                    blocks.append(('cont', c, prev_unit))
                    continue
                else:
                    kind = 'item'
            elif s.startswith('-'):
                kind = 'item'
            else:
                blocks.append(('other', c, None)); prev_unit = None; continue
            u = Unit(kind, [c], section, major, dong)
            units.append(u)
            blocks.append((kind, c, u))
            prev_unit = u
        return units, blocks

    def plain_rpr(self, units):
        """20조 본문에서 가장 흔한(=검정 보통) 글자서식."""
        cnt, keep = Counter(), {}
        for u in units:
            for p in u.paras:
                for r in p.iter(w('r')):
                    rp = r.find(w('rPr'))
                    if rp is None:
                        continue
                    t = ''.join(x.text or '' for x in r.findall(w('t')))
                    rp2 = clean_rpr(rp, False)
                    for tag in ('b', 'bCs', 'u'):
                        for e in rp2.findall(w(tag)):
                            rp2.remove(e)
                    col = rp2.find(w('color'))
                    if col is not None and col.get(w('val')) not in ('000000', 'auto'):
                        continue
                    key = etree.tostring(rp2)
                    cnt[key] += len(t)
                    keep[key] = rp2
        if not cnt:
            return None
        return keep[cnt.most_common(1)[0][0]]
