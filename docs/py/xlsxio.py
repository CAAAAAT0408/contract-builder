"""엑셀(.xlsx)을 내부 XML 그대로 고친다. 서식·인쇄설정·셀 안 글자서식(리치텍스트)을 건드리지 않는다.

지원: 셀 글자/숫자 읽기·쓰기, 문자열 바꾸기(리치텍스트 유지), 시트 복사(인쇄영역·그림·메모 포함)·삭제,
      행 넣기·지우기(수식·병합·인쇄영역·페이지 나누기·그림·메모 위치를 같이 옮김),
      간단한 수식(사칙연산·셀참조, 공유수식 포함)의 표시값 재계산, 바뀐 칸 노란색 표시.
"""
import copy
import io
import posixpath
import re
import uuid
import zipfile

from lxml import etree

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PNS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CNS = 'http://schemas.openxmlformats.org/package/2006/content-types'
XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
WS_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
WS_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


def q(tag, ns=NS):
    return '{%s}%s' % (ns, tag)


class XlsxError(Exception):
    pass


def col2num(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def num2col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def split_ref(ref):
    m = re.match(r'^\$?([A-Z]+)\$?(\d+)$', ref)
    return m.group(1), int(m.group(2))


def cell_key(ref):
    c, r = split_ref(ref)
    return r, col2num(c)


def si_text(si):
    """공유문자열 항목의 보이는 글자 (발음표기 rPh 제외)"""
    out = []
    for t in si.iter(q('t')):
        if t.getparent().tag == q('rPh') or (t.getparent().getparent() is not None and t.getparent().getparent().tag == q('rPh')):
            continue
        out.append(t.text or '')
    return ''.join(out)


class Sheet:
    def __init__(self, book, name, path):
        self.book, self.name, self.path = book, name, path
        self.root = etree.fromstring(book.files[path])
        self.data = self.root.find(q('sheetData'))
        self._index()

    def _index(self):
        self.cells = {}
        self.rows = {}
        for row in self.data.findall(q('row')):
            self.rows[int(row.get('r'))] = row
            for c in row.findall(q('c')):
                self.cells[c.get('r')] = c

    def merged(self):
        mc = self.root.find(q('mergeCells'))
        out = []
        if mc is not None:
            for m in mc.findall(q('mergeCell')):
                a, _, b = m.get('ref').partition(':')
                out.append((cell_key(a), cell_key(b or a)))
        return out

    # ---------------------------------------------------------------- 읽기
    def value(self, ref):
        """숫자 → float, 글자 → str, 없으면 None. 수식이면 표시값."""
        c = self.cells.get(ref)
        if c is None:
            return None
        t = c.get('t')
        v = c.find(q('v'))
        if t == 's':
            return si_text(self.book.sst[int(v.text)]) if v is not None else None
        if t == 'inlineStr':
            is_ = c.find(q('is'))
            return si_text(is_) if is_ is not None else ''
        if t in ('str', 'e'):
            return v.text if v is not None else None
        if t == 'b':
            return bool(int(v.text)) if v is not None else None
        if v is None or v.text is None:
            return None
        try:
            f = float(v.text)
            return int(f) if f.is_integer() else f
        except ValueError:
            return v.text

    def text(self, ref):
        v = self.value(ref)
        return '' if v is None else str(v)

    def formula(self, ref):
        c = self.cells.get(ref)
        f = c.find(q('f')) if c is not None else None
        return f

    def max_row(self):
        return max(self.rows) if self.rows else 0

    # ---------------------------------------------------------------- 쓰기
    def _cell(self, ref):
        if ref in self.cells:
            return self.cells[ref]
        col, r = split_ref(ref)
        row = self.rows.get(r)
        if row is None:
            row = etree.Element(q('row'))
            row.set('r', str(r))
            after = [k for k in self.rows if k < r]
            if after:
                self.rows[max(after)].addnext(row)
            else:
                self.data.insert(0, row)
            self.rows[r] = row
        c = etree.Element(q('c'))
        c.set('r', ref)
        n = col2num(col)
        prev = None
        for x in row.findall(q('c')):
            if cell_key(x.get('r'))[1] < n:
                prev = x
        if prev is not None:
            prev.addnext(c)
        else:
            row.insert(0, c)
        self.cells[ref] = c
        return c

    def _clear(self, c):
        for e in list(c):
            c.remove(e)
        for a in ('t',):
            c.attrib.pop(a, None)

    def set_number(self, ref, num):
        c = self._cell(ref)
        self._clear(c)
        v = etree.SubElement(c, q('v'))
        v.text = repr(float(num)).rstrip('0').rstrip('.') if float(num) != int(num) else str(int(num))
        self.book.touched.add(self.path)

    def set_text(self, ref, text):
        c = self._cell(ref)
        self._clear(c)
        c.set('t', 's')
        v = etree.SubElement(c, q('v'))
        v.text = str(self.book.add_string(text))
        self.book.touched.add(self.path)

    def replace_text(self, ref, old, new):
        """셀 글자 안의 old 를 new 로 (리치텍스트 서식 유지). 바꿨으면 True."""
        c = self.cells.get(ref)
        if c is None or c.get('t') != 's':
            if c is not None and c.get('t') == 'inlineStr':
                si = c.find(q('is'))
                if old in si_text(si):
                    _replace_runs(si, old, new)
                    self.book.touched.add(self.path)
                    return True
            return False
        v = c.find(q('v'))
        idx = int(v.text)
        if old not in si_text(self.book.sst[idx]):
            return False
        v.text = str(self.book.derive_string(idx, old, new))
        self.book.touched.add(self.path)
        return True

    def highlight(self, ref):
        c = self._cell(ref)
        c.set('s', str(self.book.yellow_style(int(c.get('s') or 0))))
        self.book.touched.add(self.path)

    def shrink(self, ref):
        """'셀에 맞춤' 켜기 (자릿수가 늘어난 금액이 ##### 로 보이지 않게)"""
        c = self.cells.get(ref)
        if c is not None:
            c.set('s', str(self.book.shrink_style(int(c.get('s') or 0))))
            self.book.touched.add(self.path)

    def clear(self, ref):
        """값·수식만 지우고 서식은 둔다."""
        c = self.cells.get(ref)
        if c is not None:
            self._clear(c)
            self.book.touched.add(self.path)

    def style(self, ref):
        c = self.cells.get(ref)
        return int(c.get('s') or 0) if c is not None else None

    def set_style(self, ref, s):
        self._cell(ref).set('s', str(s))

    # ---------------------------------------------------------------- 병합
    def unmerge(self, ref):
        mc = self.root.find(q('mergeCells'))
        if mc is None:
            return
        for m in list(mc):
            if m.get('ref') == ref:
                mc.remove(m)
        mc.set('count', str(len(mc)))

    def merge(self, ref):
        if ':' not in ref:
            return
        mc = self.root.find(q('mergeCells'))
        if mc is None:
            mc = etree.Element(q('mergeCells'))
            # mergeCells 는 sheetData / sheetCalcPr / sheetProtection / protectedRanges / scenarios / autoFilter / sortState
            #   / dataConsolidate / customSheetViews 다음 자리
            anchor = self.data
            for tag in ('sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState',
                        'dataConsolidate', 'customSheetViews'):
                e = self.root.find(q(tag))
                if e is not None:
                    anchor = e
            anchor.addnext(mc)
        m = etree.SubElement(mc, q('mergeCell'))
        m.set('ref', ref)
        mc.set('count', str(len(mc)))
        self.book.touched.add(self.path)

    # ---------------------------------------------------------------- 행 넣기·지우기
    def unshare(self):
        """공유수식을 칸마다 따로 쓴 수식으로 푼다 (행을 옮기거나 복사해도 안전하게)."""
        masters = {}
        for ref, c in self.cells.items():
            f = c.find(q('f'))
            if f is not None and f.get('t') == 'shared' and f.text:
                masters[f.get('si')] = (ref, f.text)
        for ref, c in self.cells.items():
            f = c.find(q('f'))
            if f is None or f.get('t') != 'shared':
                continue
            if not f.text:
                base = masters.get(f.get('si'))
                if not base:
                    continue
                f.text = _shift(base[1], cell_key(ref)[0] - cell_key(base[0])[0], cell_key(ref)[1] - cell_key(base[0])[1])
            for a in ('t', 'si', 'ref'):
                f.attrib.pop(a, None)
        self.book.touched.add(self.path)

    def insert_rows(self, at, n):
        """at 행 앞에 빈 행 n 개를 넣는다 (at 부터 아래는 n 줄씩 밀림)."""
        if n > 0:
            self._move_rows(RowMap('ins', at, n))

    def delete_rows(self, at, n):
        """at 행부터 n 개를 지운다."""
        if n > 0:
            self._move_rows(RowMap('del', at, n))

    def row_copy(self, r):
        """r 행 XML 사본 (나중에 put_row 로 다른 자리에 붙인다). 공유수식은 미리 unshare() 로 풀어 둘 것."""
        row = self.rows.get(r)
        return (r, copy.deepcopy(row)) if row is not None else (r, None)

    def put_row(self, snap, dst):
        """row_copy 로 떠 둔 행을 dst 자리에 놓는다 (있던 행은 바꿈). 수식은 행 차이만큼 옮긴다."""
        src, el = snap
        old = self.rows.get(dst)
        if el is None:
            if old is not None:
                self.data.remove(old)
            self._index()
            return
        row = copy.deepcopy(el)
        row.set('r', str(dst))
        for c in row.findall(q('c')):
            col, _ = split_ref(c.get('r'))
            c.set('r', f'{col}{dst}')
            f = c.find(q('f'))
            if f is not None and f.text and dst != src:
                f.text = _shift(f.text, dst - src, 0)
        if old is not None:
            old.addprevious(row)
            self.data.remove(old)
        else:
            after = [k for k in self.rows if k < dst]
            if after:
                self.rows[max(after)].addnext(row)
            else:
                self.data.insert(0, row)
        self._index()
        self.book.touched.add(self.path)

    def _move_rows(self, rm):
        self.unshare()
        # 1) 행·칸 번호
        for row in list(self.data.findall(q('row'))):
            r = int(row.get('r'))
            nr = rm.one(r)
            if nr is None:
                self.data.remove(row)
                continue
            if nr != r:
                row.set('r', str(nr))
                for c in row.findall(q('c')):
                    col, _ = split_ref(c.get('r'))
                    c.set('r', f'{col}{nr}')
        # 2) 수식: 이 시트 전부 + 이 시트를 가리키는 다른 시트 수식
        for sh in self.book.all_sheets():
            local = sh is self
            hit = False
            for c in sh.data.iter(q('c')):
                f = c.find(q('f'))
                if f is None:
                    continue
                if f.text and (local or self.name in f.text):
                    nt = _map_refs(f.text, rm, self.name, local)
                    if nt != f.text:
                        f.text = nt
                        hit = True
                if local and f.get('ref'):
                    v = _map_sqref(f.get('ref'), rm)
                    if v:
                        f.set('ref', v)
            if hit:
                self.book.touched.add(sh.path)
        # 3) 병합
        mc = self.root.find(q('mergeCells'))
        if mc is not None:
            for m in list(mc):
                v = _map_sqref(m.get('ref'), rm)
                if not v or ':' not in v or v.split(':')[0] == v.split(':')[1]:
                    mc.remove(m)
                else:
                    m.set('ref', v)
            mc.set('count', str(len(mc)))
            if not len(mc):
                self.root.remove(mc)
        # 4) 범위를 적어 두는 그 밖의 자리
        for tag, attr, drop in (('dimension', 'ref', False), ('selection', 'sqref', False),
                                ('conditionalFormatting', 'sqref', True), ('dataValidation', 'sqref', True),
                                ('hyperlink', 'ref', True), ('ignoredError', 'sqref', True),
                                ('protectedRange', 'sqref', True), ('autoFilter', 'ref', True)):
            for e in list(self.root.iter(q(tag))):
                v = e.get(attr)
                if not v:
                    continue
                nv = _map_sqref(v, rm)
                if nv:
                    e.set(attr, nv)
                elif drop:
                    e.getparent().remove(e)
                else:
                    e.set(attr, f'A{rm.start}')
        for e in self.root.iter(q('selection'), q('pane'), q('sheetView')):
            for attr in ('activeCell', 'topLeftCell'):
                v = e.get(attr)
                if v and re.fullmatch(r'[A-Z]+\d+', v):
                    nv = _map_sqref(v, rm)
                    e.set(attr, nv or f'{split_ref(v)[0]}{rm.start}')
        for e in self.root.iter(q('formula'), q('formula1'), q('formula2')):
            if e.text:
                e.text = _map_refs(e.text, rm, self.name, True)
        rb = self.root.find(q('rowBreaks'))
        if rb is not None:
            for b in list(rb):
                nr = rm.one(int(b.get('id')))
                if nr is None:
                    nr = rm.start - 1
                if nr < 1 or any(int(x.get('id')) == nr for x in rb if x is not b):
                    rb.remove(b)
                else:
                    b.set('id', str(nr))
            rb.set('count', str(len(rb)))
            if rb.get('manualBreakCount') is not None:
                rb.set('manualBreakCount', str(sum(1 for x in rb if x.get('man') == '1')))
        # 5) 인쇄영역 등 이름 정의
        self.book.map_defined_names(self.name, rm)
        # 6) 그림·메모
        self._move_objects(rm)
        self._index()
        self.book.touched.add(self.path)

    def rel_parts(self):
        """[(관계 종류, 부품 경로)] — drawing, comments, vmlDrawing, printerSettings …"""
        return self.book.part_rels(self.path)

    def _move_objects(self, rm):
        files = self.book.files
        for kind, part in self.rel_parts():
            if part not in files:
                continue
            if kind == 'drawing':
                root = etree.fromstring(files[part])
                for tag in ('from', 'to'):
                    for e in root.iter('{%s}%s' % (XDR, tag)):
                        re_ = e.find('{%s}row' % XDR)
                        if re_ is None or not (re_.text or '').strip().isdigit():
                            continue
                        r = int(re_.text) + 1
                        nr = rm.one(r)
                        re_.text = str((nr if nr is not None else rm.start) - 1)
                files[part] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
            elif kind in ('comments', 'threadedComment'):
                root = etree.fromstring(files[part])
                for e in list(root.iter()):
                    if not isinstance(e.tag, str) or e.get('ref') is None:
                        continue
                    nv = _map_sqref(e.get('ref'), rm)
                    if nv:
                        e.set('ref', nv)
                    else:
                        e.getparent().remove(e)
                files[part] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
            elif kind == 'vmlDrawing':
                txt = files[part].decode('utf-8', 'replace')

                def shape(m):
                    blk = m.group(0)
                    rm_ = re.search(r'<x:Row>(\d+)</x:Row>', blk)
                    if not rm_:
                        return blk
                    r = int(rm_.group(1)) + 1
                    nr = rm.one(r)
                    if nr is None:
                        return ''
                    d = nr - r
                    if not d:
                        return blk
                    blk = blk.replace(rm_.group(0), f'<x:Row>{nr - 1}</x:Row>')

                    def anchor(a):
                        v = [x.strip() for x in a.group(1).split(',')]
                        if len(v) == 8:
                            v[2], v[6] = str(int(v[2]) + d), str(int(v[6]) + d)
                        return '<x:Anchor>' + ', '.join(v) + '</x:Anchor>'
                    return re.sub(r'<x:Anchor>([^<]*)</x:Anchor>', anchor, blk)
                txt = re.sub(r'<v:shape\b.*?</v:shape>', shape, txt, flags=re.S)
                files[part] = txt.encode('utf-8')

    # ---------------------------------------------------------------- 수식 표시값 다시 계산
    def recalc(self):
        """사칙연산·셀참조만 있는 수식의 표시값(<v>)을 다시 계산한다. 못 푸는 수식은 값을 지워 Excel 이 계산하게 둔다."""
        shared = {}
        formulas = {}
        for ref, c in self.cells.items():
            f = c.find(q('f'))
            if f is None:
                continue
            if f.get('t') == 'shared' and f.text:
                shared[f.get('si')] = (ref, f.text)
        for ref, c in self.cells.items():
            f = c.find(q('f'))
            if f is None:
                continue
            if f.get('t') == 'shared' and not f.text:
                base = shared.get(f.get('si'))
                if not base:
                    continue
                formulas[ref] = _shift(base[1], cell_key(ref)[0] - cell_key(base[0])[0], cell_key(ref)[1] - cell_key(base[0])[1])
            elif f.get('t') in (None, 'shared', 'normal'):
                formulas[ref] = f.text or ''
        vals = {}
        for _ in range(12):
            changed = False
            for ref, expr in formulas.items():
                res = _eval(expr, lambda r: vals[r] if r in vals else self.value(r))
                if res is _FAIL:
                    continue
                if vals.get(ref, object()) != res:
                    vals[ref] = res
                    changed = True
            if not changed:
                break
        for ref, expr in formulas.items():
            c = self.cells[ref]
            v = c.find(q('v'))
            res = vals.get(ref, _FAIL)
            if res is _FAIL:
                if v is not None:
                    c.remove(v)
                    c.attrib.pop('t', None)
                continue
            if v is None:
                v = etree.SubElement(c, q('v'))
            if isinstance(res, str):
                c.set('t', 'str')
                v.text = res
            else:
                c.attrib.pop('t', None)
                v.text = str(int(res)) if float(res).is_integer() else repr(res)
        self.book.touched.add(self.path)

    def tostring(self):
        return etree.tostring(self.root, xml_declaration=True, encoding='UTF-8', standalone=True)


_FAIL = object()


def _shift(expr, dr, dc):
    def rep(m):
        cabs, col, rabs, row = m.group(1), m.group(2), m.group(3), m.group(4)
        if not cabs:
            col = num2col(col2num(col) + dc)
        if not rabs:
            row = str(int(row) + dr)
        return f'{cabs}{col}{rabs}{row}'
    return re.sub(r'(?<![A-Za-z_])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(])', rep, expr)


class RowMap:
    """행 넣기('ins': start 부터 n 줄 밀림)·지우기('del': start..start+n-1 사라짐) 때 행 번호를 옮기는 규칙."""

    def __init__(self, kind, start, n):
        self.kind, self.start, self.n = kind, start, n
        self.end = start + n - 1

    def one(self, r):
        if self.kind == 'ins':
            return r + self.n if r >= self.start else r
        if r < self.start:
            return r
        return None if r <= self.end else r - self.n

    def rng(self, r1, r2):
        if self.kind == 'ins':
            return self.one(r1), self.one(r2)
        if r1 >= self.start and r2 <= self.end:
            return None
        a = r1 if r1 < self.start else (self.start if r1 <= self.end else r1 - self.n)
        b = r2 if r2 < self.start else (self.start - 1 if r2 <= self.end else r2 - self.n)
        return (a, b) if a <= b else None


_SHEET_PREFIX = r"((?:'(?:[^']|'')+'|[^\W\d][\w.]*)!)"
_REF = re.compile(r"(?<![\w.$'!])" + _SHEET_PREFIX + r"?(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?(?![\w(!])")
_ROWS = re.compile(_SHEET_PREFIX + r"(\$?)(\d+):(\$?)(\d+)(?![\w])")


def _sheet_of(prefix):
    nm = prefix[:-1]
    return nm[1:-1].replace("''", "'") if nm.startswith("'") else nm


def _map_refs(expr, rm, sheet_name, local):
    """수식 글자 안의 셀 참조를 행 이동 규칙 rm 대로 고친다.
    local=True 면 시트 이름 없는 참조도 이 시트(sheet_name)를 가리키는 것으로 본다. 사라진 칸은 #REF!"""
    def one(m):
        pre = m.group(1)
        if pre:
            if _sheet_of(pre) != sheet_name:
                return m.group(0)
        elif not local:
            return m.group(0)
        pre = pre or ''
        if m.group(6):
            res = rm.rng(int(m.group(5)), int(m.group(9)))
            if res is None:
                return pre + '#REF!'
            return f'{pre}{m.group(2)}{m.group(3)}{m.group(4)}{res[0]}:{m.group(6)}{m.group(7)}{m.group(8)}{res[1]}'
        res = rm.one(int(m.group(5)))
        return pre + '#REF!' if res is None else f'{pre}{m.group(2)}{m.group(3)}{m.group(4)}{res}'

    def rows(m):
        if _sheet_of(m.group(1)) != sheet_name:
            return m.group(0)
        res = rm.rng(int(m.group(3)), int(m.group(5)))
        if res is None:
            return m.group(1) + '#REF!'
        return f'{m.group(1)}{m.group(2)}{res[0]}:{m.group(4)}{res[1]}'
    parts = expr.split('"')
    for i in range(0, len(parts), 2):
        parts[i] = _ROWS.sub(rows, _REF.sub(one, parts[i]))
    return '"'.join(parts)


def _map_sqref(s, rm):
    """'A1 B2:C5' 같은 칸 범위 목록을 옮긴다. 다 사라지면 None."""
    out = []
    for part in s.split():
        a, _, b = part.partition(':')
        if not re.fullmatch(r'\$?[A-Z]+\$?\d+', a) or (b and not re.fullmatch(r'\$?[A-Z]+\$?\d+', b)):
            out.append(part)        # 열 전체(A:A) 등
            continue
        ca, ra = split_ref(a)
        if b:
            cb, rb = split_ref(b)
            res = rm.rng(ra, rb)
            if res:
                out.append(f'{ca}{res[0]}:{cb}{res[1]}')
        else:
            r = rm.one(ra)
            if r is not None:
                out.append(f'{ca}{r}')
    return ' '.join(out) or None


def rels_path(path):
    d, f = posixpath.split(path)
    return f'{d}/_rels/{f}.rels'


def _eval(expr, get):
    expr = expr.strip()
    if re.fullmatch(r'\$?[A-Z]{1,3}\$?\d+', expr):
        v = get(expr.replace('$', ''))
        return '' if v is None else v
    if not re.fullmatch(r'[\sA-Z$\d.+\-*/()]+', expr):
        return _FAIL

    def rep(m):
        v = get(m.group(0).replace('$', ''))
        if v is None or v == '':
            return '0'
        if isinstance(v, str):
            raise ValueError
        return repr(float(v))
    try:
        py = re.sub(r'\$?[A-Z]{1,3}\$?\d+', rep, expr)
        res = eval(py, {'__builtins__': {}}, {})
        return round(res, 6)
    except Exception:
        return _FAIL


def _replace_runs(si, old, new):
    """공유문자열 항목 안의 글자를 바꾼다. 여러 run 에 걸치면 첫 run 에 모은다."""
    ts = [t for t in si.iter(q('t')) if not any(a.tag == q('rPh') for a in t.iterancestors())]
    start = 0
    while True:
        full = ''.join(t.text or '' for t in ts)
        i = full.find(old, start)
        if i < 0:
            return
        start = i + len(new)
        j = i + len(old)
        pos, first = 0, True
        for t in ts:
            s = t.text or ''
            a, b = pos, pos + len(s)
            pos = b
            if b <= i or a >= j or a == b:
                continue
            pre = s[:max(0, i - a)]
            post = s[j - a:] if j < b else ''
            t.text = (pre + new + post) if first else post
            t.set(XML_SPACE, 'preserve')
            first = False
        # 글자가 다 빠진 서식 조각(run)은 지운다 (빈 <t> 가 든 run 이 있으면 Excel 이 파일을 열지 못한다)
        for t in list(ts):
            r = t.getparent()
            if not t.text and r.tag == q('r') and len(si.findall(q('r'))) > 1:
                si.remove(r)
                ts.remove(t)
        if old in new:
            return


class Book:
    def __init__(self, data):
        try:
            self.zin = zipfile.ZipFile(io.BytesIO(data))
        except Exception:
            raise XlsxError('엑셀(.xlsx) 파일이 아닙니다. .xls 이면 Excel 에서 .xlsx 로 저장해 주세요.')
        self.order = [i.filename for i in self.zin.infolist()]
        self.files = {n: self.zin.read(n) for n in self.order}
        if 'xl/workbook.xml' not in self.files:
            raise XlsxError('엑셀(.xlsx) 파일이 아닙니다.')
        self.wb = etree.fromstring(self.files['xl/workbook.xml'])
        self.rels = etree.fromstring(self.files['xl/_rels/workbook.xml.rels'])
        self.ct = etree.fromstring(self.files['[Content_Types].xml'])
        sst_path = self._rel_target('sharedStrings')
        self.sst_path = sst_path
        if sst_path and sst_path in self.files:
            self.sst_root = etree.fromstring(self.files[sst_path])
        else:
            self.sst_root = etree.Element(q('sst'))
            self.sst_path = None
        self.sst = self.sst_root.findall(q('si'))
        self.touched = set()
        self._sheets = {}
        self._derived = {}
        self._styles = None
        self._yellow = {}
        self._structure_changed = False

    def _rel_target(self, kind):
        for r in self.rels:
            if r.get('Type', '').endswith('/' + kind):
                t = r.get('Target')
                return t.lstrip('/') if t.startswith('/') else 'xl/' + t
        return None

    def sheet_list(self):
        rid2t = {r.get('Id'): r.get('Target') for r in self.rels}
        out = []
        for s in self.wb.find(q('sheets')):
            t = rid2t[s.get(q('id', RNS))]
            path = t.lstrip('/') if t.startswith('/') else 'xl/' + t
            out.append({'name': s.get('name'), 'path': path, 'state': s.get('state', 'visible'), 'el': s})
        return out

    def sheet(self, name):
        if name not in self._sheets:
            info = next((s for s in self.sheet_list() if s['name'] == name), None)
            if info is None:
                raise XlsxError(f"'{name}' 시트가 없습니다.")
            self._sheets[name] = Sheet(self, name, info['path'])
        return self._sheets[name]

    def all_sheets(self):
        return [self.sheet(s['name']) for s in self.sheet_list() if s['path'] in self.files]

    def map_defined_names(self, sheet_name, rm):
        dn = self.wb.find(q('definedNames'))
        if dn is None:
            return
        for d in dn:
            if d.text and sheet_name in d.text.replace("''", "'"):
                nt = _map_refs(d.text, rm, sheet_name, False)
                if nt != d.text:
                    d.text = nt
                    self.touched.add('xl/workbook.xml')

    def part_rels(self, path):
        p = rels_path(path)
        if p not in self.files:
            return []
        out = []
        for r in etree.fromstring(self.files[p]):
            if r.get('TargetMode') == 'External':
                continue
            out.append((r.get('Type', '').rsplit('/', 1)[-1], _norm(posixpath.dirname(path) + '/' + r.get('Target'))))
        return out

    def _ct_override(self, part):
        return next((o for o in self.ct if o.get('PartName') == '/' + part), None)

    def _copy_part(self, part):
        """부품 하나를 새 이름으로 복사 (콘텐츠 형식·딸린 관계 파일 포함). 새 경로를 돌려준다."""
        m = re.match(r'(.*?)(\d*)(\.[^./]+)$', part)
        base, ext = m.group(1), m.group(3)
        k = 1
        while f'{base}{k}{ext}' in self.files:
            k += 1
        new = f'{base}{k}{ext}'
        self.files[new] = self.files[part]
        self.order.append(new)
        ov = self._ct_override(part)
        if ov is not None:
            nov = copy.deepcopy(ov)
            nov.set('PartName', '/' + new)
            self.ct.append(nov)
        rp = rels_path(part)
        if rp in self.files:         # 그림이 쓰는 이미지 등은 같은 파일을 같이 가리킨다
            nrp = rels_path(new)
            self.files[nrp] = self.files[rp]
            self.order.append(nrp)
        return new

    # ---------------------------------------------------------------- 공유문자열
    def add_string(self, text):
        si = etree.SubElement(self.sst_root, q('si'))
        t = etree.SubElement(si, q('t'))
        t.text = text
        if text != text.strip() or '\n' in text:
            t.set(XML_SPACE, 'preserve')
        self.sst.append(si)
        return len(self.sst) - 1

    def derive_string(self, idx, old, new):
        key = (idx, old, new)
        if key not in self._derived:
            si = copy.deepcopy(self.sst[idx])
            _replace_runs(si, old, new)
            self.sst_root.append(si)
            self.sst.append(si)
            self._derived[key] = len(self.sst) - 1
        return self._derived[key]

    # ---------------------------------------------------------------- 셀에 맞춤
    def shrink_style(self, s):
        if not hasattr(self, '_shrink'):
            self._shrink = {}
        if s in self._shrink:
            return self._shrink[s]
        if self._styles is None:
            self._styles = etree.fromstring(self.files['xl/styles.xml'])
        xfs = self._styles.find(q('cellXfs'))
        al = xfs[s].find(q('alignment'))
        if al is not None and al.get('shrinkToFit') == '1' and al.get('wrapText') != '1':
            self._shrink[s] = s
            return s
        nx = copy.deepcopy(xfs[s])
        al = nx.find(q('alignment'))
        if al is None:
            al = etree.Element(q('alignment'))
            nx.insert(0, al)
        al.set('shrinkToFit', '1')
        al.attrib.pop('wrapText', None)
        nx.set('applyAlignment', '1')
        xfs.append(nx)
        xfs.set('count', str(len(xfs)))
        self._shrink[s] = len(xfs) - 1
        return self._shrink[s]

    # ---------------------------------------------------------------- 노란색 표시
    def yellow_style(self, s):
        if s in self._yellow:
            return self._yellow[s]
        if self._styles is None:
            self._styles = etree.fromstring(self.files['xl/styles.xml'])
        st = self._styles
        fills = st.find(q('fills'))
        if not hasattr(self, '_fill_id'):
            f = etree.SubElement(fills, q('fill'))
            pf = etree.SubElement(f, q('patternFill'))
            pf.set('patternType', 'solid')
            fg = etree.SubElement(pf, q('fgColor'))
            fg.set('rgb', 'FFFFFF00')
            bg = etree.SubElement(pf, q('bgColor'))
            bg.set('indexed', '64')
            fills.set('count', str(len(fills)))
            self._fill_id = len(fills) - 1
        xfs = st.find(q('cellXfs'))
        nx = copy.deepcopy(xfs[s])
        nx.set('fillId', str(self._fill_id))
        nx.set('applyFill', '1')
        xfs.append(nx)
        xfs.set('count', str(len(xfs)))
        self._yellow[s] = len(xfs) - 1
        return self._yellow[s]

    # ---------------------------------------------------------------- 시트 복사
    def copy_sheet(self, src_name, new_name, after=None, before=None):
        """src_name 시트를 복사해 new_name 으로. after 시트 바로 뒤(또는 before 시트 바로 앞)에 둔다 (없으면 원본 뒤)."""
        sheets = self.sheet_list()
        src = next(s for s in sheets if s['name'] == src_name)
        names = [s['name'] for s in sheets]
        if before in names:
            pos = names.index(before)
            anchor = sheets[pos - 1] if pos else None
        else:
            anchor = next((s for s in sheets if s['name'] == after), src)
            pos = names.index(anchor['name']) + 1
        nums = [int(m.group(1)) for n in self.files for m in [re.match(r'xl/worksheets/sheet(\d+)\.xml$', n)] if m]
        new_path = f'xl/worksheets/sheet{max(nums) + 1}.xml'
        # 시트 XML (편집 중이면 현재 상태로)
        data = self._sheets[src_name].tostring() if src_name in self._sheets else self.files[src['path']]
        root = etree.fromstring(data)
        for sv in root.iter(q('sheetView')):
            sv.attrib.pop('tabSelected', None)
        self.files[new_path] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        self.order.append(new_path)
        # 시트 rels (인쇄설정·그림·메모 등)
        src_rels = rels_path(src['path'])
        if src_rels in self.files:
            rr = etree.fromstring(self.files[src_rels])
            copied = []
            for rel in rr:
                if rel.get('TargetMode') == 'External':
                    continue
                full = _norm(posixpath.dirname(src['path']) + '/' + rel.get('Target'))
                if full in self.files:
                    nf = self._copy_part(full)
                    copied.append((rel.get('Type', '').rsplit('/', 1)[-1], nf))
                    rel.set('Target', posixpath.relpath(nf, posixpath.dirname(new_path)))
            # 스레드 메모의 고유번호(GUID)는 통합문서 안에서 겹치면 안 된다 → 복사본은 새 번호로
            tc = [p for k, p in copied if k == 'threadedComment']
            if tc:
                ids = set()
                for p in tc:
                    ids |= set(re.findall(r'\bid="\{([0-9A-Fa-f\-]{36})\}"', self.files[p].decode('utf-8')))
                new_ids = {i: str(uuid.uuid4()).upper() for i in ids}
                for k, p in copied:
                    if k in ('threadedComment', 'comments'):
                        txt = self.files[p].decode('utf-8')
                        for o, n in new_ids.items():
                            txt = re.sub(re.escape(o), n, txt, flags=re.I)
                        self.files[p] = txt.encode('utf-8')
            nr = rels_path(new_path)
            self.files[nr] = etree.tostring(rr, xml_declaration=True, encoding='UTF-8', standalone=True)
            self.order.append(nr)
        # 통합문서 관계
        ids = [int(re.sub(r'\D', '', r.get('Id')) or 0) for r in self.rels]
        rid = f'rId{max(ids) + 1}'
        rel = etree.SubElement(self.rels, q('Relationship', PNS))
        rel.set('Id', rid)
        rel.set('Type', WS_TYPE)
        rel.set('Target', new_path[len('xl/'):])
        # 콘텐츠 형식
        ov = etree.SubElement(self.ct, q('Override', CNS))
        ov.set('PartName', '/' + new_path)
        ov.set('ContentType', WS_CT)
        # 시트 목록
        sid = max(int(s['el'].get('sheetId')) for s in sheets) + 1
        el = etree.Element(q('sheet'))
        el.set('name', new_name)
        el.set('sheetId', str(sid))
        el.set(q('id', RNS), rid)
        if anchor is not None:
            anchor['el'].addnext(el)
        else:
            sheets[0]['el'].addprevious(el)
        # 열린 시트 번호(activeTab)가 새 시트 뒤를 가리키면 한 칸 민다 (안 그러면 시트 두 개가 같이 선택된다)
        for wv in self.wb.iter(q('workbookView')):
            for attr in ('activeTab', 'firstSheet'):
                v = wv.get(attr)
                if v is not None and int(v) >= pos:
                    wv.set(attr, str(int(v) + 1))
        # 이름 정의 (인쇄영역 등): 뒤 시트 번호 밀기 + 새 시트용 복사
        dn = self.wb.find(q('definedNames'))
        if dn is not None:
            src_idx = [s['name'] for s in sheets].index(src_name)
            add = []
            for d in dn:
                lid = d.get('localSheetId')
                if lid is None:
                    continue
                lid = int(lid)
                if lid == src_idx:
                    nd = copy.deepcopy(d)
                    nd.set('localSheetId', str(pos))
                    nd.text = re.sub(r"^('?)" + re.escape(src_name) + r"\1!", lambda m: _qname(new_name) + '!', d.text or '')
                    add.append(nd)
                if lid >= pos:
                    d.set('localSheetId', str(lid + 1))
            for nd in add:
                dn.append(nd)
            # Excel 은 이름 정의를 이름순으로 정렬해 두지만 순서는 필수가 아니다
        self.touched.add('xl/workbook.xml')
        return self.sheet(new_name)

    # ---------------------------------------------------------------- 시트 이름 바꾸기
    def rename_sheet(self, old, new):
        info = next(s for s in self.sheet_list() if s['name'] == old)
        info['el'].set('name', new)
        refs = ["'" + old.replace("'", "''") + "'!"] + ([old + '!'] if _qname(old) == old else [])

        def fix(t):
            for r in refs:
                if t and r in t:
                    t = re.sub(r'(?<![\w.\'])' + re.escape(r), lambda m: _qname(new) + '!', t)
            return t
        dn = self.wb.find(q('definedNames'))
        if dn is not None:
            for d in dn:
                d.text = fix(d.text)
        for sh in self.all_sheets():
            for c in sh.data.iter(q('c')):
                f = c.find(q('f'))
                if f is not None and f.text and old in f.text:
                    f.text = fix(f.text)
                    self.touched.add(sh.path)
        sh = self._sheets.pop(old, None)
        if sh is not None:
            sh.name = new
            self._sheets[new] = sh
        self.touched.add('xl/workbook.xml')

    # ---------------------------------------------------------------- 시트 지우기
    def delete_sheet(self, name):
        sheets = self.sheet_list()
        idx = [s['name'] for s in sheets].index(name)
        info = sheets[idx]
        rid = info['el'].get(q('id', RNS))
        info['el'].getparent().remove(info['el'])
        for r in list(self.rels):
            if r.get('Id') == rid:
                self.rels.remove(r)
        parts = self.part_rels(info['path'])
        self._sheets.pop(name, None)
        # 다른 시트가 같이 쓰지 않는 딸린 부품(그림·메모·인쇄설정)도 지운다
        used = {p for s in self.sheet_list() for _, p in self.part_rels(s['path'])}
        for _, p in parts:
            if p not in used:
                self._drop_part(p)
        self._drop_part(info['path'])
        dn = self.wb.find(q('definedNames'))
        if dn is not None:
            for d in list(dn):
                lid = d.get('localSheetId')
                if lid is None:
                    continue
                lid = int(lid)
                if lid == idx:
                    dn.remove(d)
                elif lid > idx:
                    d.set('localSheetId', str(lid - 1))
            if not len(dn):
                self.wb.remove(dn)
        for wv in self.wb.iter(q('workbookView')):
            for attr in ('activeTab', 'firstSheet'):
                v = wv.get(attr)
                if v is not None and int(v) > idx:
                    wv.set(attr, str(int(v) - 1))
        self._structure_changed = True
        self.touched.add('xl/workbook.xml')

    def _drop_part(self, part):
        for p in (part, rels_path(part)):
            if p in self.files:
                del self.files[p]
        for o in list(self.ct):
            if o.get('PartName') == '/' + part:
                self.ct.remove(o)

    def _fix_selection(self):
        """열린 시트(activeTab)가 보이는 시트를 가리키게 하고, 그 시트 하나만 선택 표시한다."""
        sheets = self.sheet_list()
        wv = next(self.wb.iter(q('workbookView')), None)
        if wv is None or not sheets:
            return
        act = min(int(wv.get('activeTab') or 0), len(sheets) - 1)
        if sheets[act]['state'] != 'visible':
            act = next((i for i, s in enumerate(sheets) if s['state'] == 'visible'), 0)
        wv.set('activeTab', str(act))
        if int(wv.get('firstSheet') or 0) > act:
            wv.set('firstSheet', str(act))
        for i, s in enumerate(sheets):
            sh = self.sheet(s['name'])
            for sv in sh.root.iter(q('sheetView')):
                if i == act:
                    sv.set('tabSelected', '1')
                else:
                    sv.attrib.pop('tabSelected', None)

    # ---------------------------------------------------------------- 저장
    def save(self):
        if self._structure_changed:
            self._fix_selection()
        for name, sh in self._sheets.items():
            self.files[sh.path] = sh.tostring()
        if self.sst_path:
            self.sst_root.set('uniqueCount', str(len(self.sst)))
            self.sst_root.set('count', str(len(self.sst)))
            self.files[self.sst_path] = etree.tostring(self.sst_root, xml_declaration=True, encoding='UTF-8', standalone=True)
        if self._styles is not None:
            self.files['xl/styles.xml'] = etree.tostring(self._styles, xml_declaration=True, encoding='UTF-8', standalone=True)
        # 수식 계산순서 정보는 지우고(Excel 이 다시 만든다), 열 때 전체 재계산하게 한다
        calc = self._rel_target('calcChain')
        if calc:
            for r in list(self.rels):
                if r.get('Type', '').endswith('/calcChain'):
                    self.rels.remove(r)
            for o in list(self.ct):
                if o.get('PartName') == '/' + calc:
                    self.ct.remove(o)
            self.files.pop(calc, None)
        cp = self.wb.find(q('calcPr'))
        if cp is None:
            # calcPr 은 sheets / externalReferences / definedNames 다음 자리
            cp = etree.Element(q('calcPr'))
            anchor = None
            for tag in ('sheets', 'functionGroups', 'externalReferences', 'definedNames'):
                e = self.wb.find(q(tag))
                if e is not None:
                    anchor = e
            anchor.addnext(cp)
        cp.set('fullCalcOnLoad', '1')
        self.files['xl/workbook.xml'] = etree.tostring(self.wb, xml_declaration=True, encoding='UTF-8', standalone=True)
        self.files['xl/_rels/workbook.xml.rels'] = etree.tostring(self.rels, xml_declaration=True, encoding='UTF-8', standalone=True)
        self.files['[Content_Types].xml'] = etree.tostring(self.ct, xml_declaration=True, encoding='UTF-8', standalone=True)
        out = io.BytesIO()
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
            for n in self.order:
                if n in self.files:
                    z.writestr(n, self.files[n])
        return out.getvalue()


def _norm(path):
    parts = []
    for p in path.split('/'):
        if p == '..':
            parts.pop()
        elif p and p != '.':
            parts.append(p)
    return '/'.join(parts)


def _qname(name):
    return "'" + name.replace("'", "''") + "'" if re.search(r"[^A-Za-z0-9_.]", name) or name[:1].isdigit() else name
