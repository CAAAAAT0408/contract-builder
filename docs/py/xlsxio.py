"""엑셀(.xlsx)을 내부 XML 그대로 고친다. 서식·인쇄설정·셀 안 글자서식(리치텍스트)을 건드리지 않는다.

지원: 셀 글자/숫자 읽기·쓰기, 문자열 바꾸기(리치텍스트 유지), 시트 복사(인쇄영역 포함),
      간단한 수식(사칙연산·셀참조, 공유수식 포함)의 표시값 재계산, 바뀐 칸 노란색 표시.
"""
import copy
import io
import re
import zipfile

from lxml import etree

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PNS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CNS = 'http://schemas.openxmlformats.org/package/2006/content-types'
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
    def copy_sheet(self, src_name, new_name):
        sheets = self.sheet_list()
        src = next(s for s in sheets if s['name'] == src_name)
        pos = [s['name'] for s in sheets].index(src_name) + 1
        nums = [int(m.group(1)) for n in self.files for m in [re.match(r'xl/worksheets/sheet(\d+)\.xml$', n)] if m]
        new_path = f'xl/worksheets/sheet{max(nums) + 1}.xml'
        # 시트 XML (편집 중이면 현재 상태로)
        data = self._sheets[src_name].tostring() if src_name in self._sheets else self.files[src['path']]
        root = etree.fromstring(data)
        for sv in root.iter(q('sheetView')):
            sv.attrib.pop('tabSelected', None)
        self.files[new_path] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        self.order.append(new_path)
        # 시트 rels (인쇄설정 등)
        src_rels = src['path'].replace('worksheets/', 'worksheets/_rels/') + '.rels'
        if src_rels in self.files:
            rr = etree.fromstring(self.files[src_rels])
            for rel in rr:
                tgt = rel.get('Target')
                full = _norm('xl/worksheets/' + tgt)
                if full in self.files:
                    base, ext = re.match(r'(.*?)(\d*)(\.[^.]+)$', full).group(1), re.match(r'.*(\.[^.]+)$', full).group(1)
                    k = 1
                    while f'{base}{k}{ext}' in self.files:
                        k += 1
                    nf = f'{base}{k}{ext}'
                    self.files[nf] = self.files[full]
                    self.order.append(nf)
                    rel.set('Target', '../' + nf[len('xl/'):])
            nr = new_path.replace('worksheets/', 'worksheets/_rels/') + '.rels'
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
        src['el'].addnext(el)
        # 열린 시트 번호(activeTab)가 새 시트 뒤를 가리키면 한 칸 민다 (안 그러면 시트 두 개가 같이 선택된다)
        for wv in self.wb.iter(q('workbookView')):
            for attr in ('activeTab', 'firstSheet'):
                v = wv.get(attr)
                if v is not None and int(v) >= pos:
                    wv.set(attr, str(int(v) + 1))
        # 이름 정의 (인쇄영역 등): 뒤 시트 번호 밀기 + 새 시트용 복사
        dn = self.wb.find(q('definedNames'))
        if dn is not None:
            src_idx = pos - 1
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

    # ---------------------------------------------------------------- 저장
    def save(self):
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
