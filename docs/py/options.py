"""추가선택품목(발코니 확장·시스템에어컨·유상옵션) 계약서 엑셀을 새 현장 모집공고 기준으로 고친다.

샘플 계약서 = 주택형(타입)마다 시트 하나. 각 시트는
  1) 발코니확장  2) 시스템에어컨  3) 추가선택품목  4) 패키지 옵션  5) 무상선택  제2조 납부처 …
표 제목으로 구역을 찾고, 공고의 같은 표에서 타입·등급·품목·제조사가 맞는 줄의 금액(과 에어컨 설치부위 문구)을 가져온다.
공고에 있는데 시트가 없는 타입은 가장 비슷한 시트를 복사해 만든다.
"""
import difflib
import re
from collections import OrderedDict

from hwp_reader import HwpError, read_hwp, read_hwp_tables, table_grid
from notice import extract_fields
from xlsxio import Book, XlsxError, col2num, num2col

CIRCLED = '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳'


class OptionError(Exception):
    pass


def nospace(s):
    return re.sub(r'\s+|⏎', '', s or '')


def money(s):
    d = re.sub(r'[^\d]', '', s or '')
    return int(d) if d else None


def types_in(s):
    s = s or ''
    if re.search(r'전\s*(주택형|타입|세대)', s):
        return {'*'}
    return set(re.findall(r'\d{2,3}[A-Z]{0,3}', s.replace('⏎', ' ')))


def okey(s):
    s = re.sub('[' + CIRCLED + ']', '', s or '')
    s = re.sub(r'\([^)]*\)', '', s.replace('⏎', ' '))
    return re.sub(r'[\s+/·,.\-]', '', s)


def clean_desc(s):
    s = (s or '').replace('⏎', ' ').strip()
    s = re.sub(r'^-\s*', '', s)
    return re.sub(r'\s{2,}', ' ', s).rstrip(',').strip()


GRADES = ('일반형', '고급형', '프리미엄형', '기본형')
TYPE_RE = re.compile(r'\d{2,3}[A-Z]{1,3}')
DATE_RE = re.compile(r'(\d{4})\s*[.\-]\s*(\d{1,2})\s*[.\-]\s*(\d{1,2})')
HOLDER_RE = re.compile(r'신탁|\(주\)|㈜|주식회사|공사|금융센터')
# 가전 제조사 (옵션 표 마지막 구분 칸이 제조사인지 가린다)
BRAND_RE = re.compile(r'전자|매직|LG|삼성|SK|코웨이|쿠쿠|쿠첸|위니아|캐리어|린나이|경동|귀뚜라미|밀레|보쉬|일렉트로룩스')


def grade_of(s):
    s = nospace(s)
    for g in GRADES:
        if g in s:
            return g
    return s


def ymd(s):
    """'2024-06-11', '2025.3.5' → '2024.06.11', '2025.03.05' (날짜가 없으면 '')"""
    m = DATE_RE.search(s or '')
    return f'{m.group(1)}.{int(m.group(2)):02d}.{int(m.group(3)):02d}' if m else ''


def sheet_types(name):
    """시트 이름이 뜻하는 주택형 목록. '84A' → ['84A'], '84A,B' · '84A,84B' → ['84A', '84B'], 그 밖 → []"""
    out, num = [], ''
    for p in re.split(r'[,·/]', nospace(name)):
        m = re.fullmatch(r'(\d{2,3})?([A-Z]{0,3})', p)
        if not p or not m or not (m.group(1) or num):
            return []
        num = m.group(1) or num
        out.append(num + m.group(2))
    return out


def read_matrix(g, t, hr, tcols, out):
    """주택형이 열 제목인 옵션·패키지 표 (예: 구분 | 구분 | 구분 | 84A | 84B | 비고) 를 타입별 줄로 푼다."""
    first = min(tcols)
    rows = []
    for r in range(hr + 1, t['rows']):
        labels = []
        for c in range(first):
            txt, k = g[r][c]
            if nospace(txt) and (not labels or labels[-1][1] != k):
                labels.append((txt, k))
        prices = {T: money(g[r][c][0]) for c, T in tcols.items()}
        if labels and any(v and v >= 10000 for v in prices.values()):
            rows.append(([x[0] for x in labels], prices))
    if sum(nospace(labels[0]).endswith('패키지') for labels, _ in rows) * 2 > len(rows):
        for labels, prices in rows:
            for T, v in prices.items():
                if v:
                    out['packages'].append({'name': nospace(labels[0]), 'types': {T}, 'price': v})
        return
    for labels, prices in rows:
        brand = labels.pop() if len(labels) >= 2 and BRAND_RE.search(nospace(labels[-1])) else ''
        cat = labels[-2] if len(labels) >= 2 else ''
        for T, v in prices.items():
            if v:
                out['options'].append({'cat': cat, 'item': labels[-1], 'brand': nospace(brand), 'types': {T},
                                       'price': v, 'types_txt': T})


# ------------------------------------------------------------------ 공고 읽기
def read_notice_options(nbytes):
    try:
        paras = read_hwp(nbytes)
        tables = read_hwp_tables(nbytes)
    except HwpError as e:
        raise OptionError(str(e))
    fields = extract_fields(paras)
    out = {'balcony': [], 'aircon': [], 'options': [], 'packages': [], 'date': '', 'bank': '', 'account': '',
           'holder': '', 'fields': fields}
    for t in tables:
        if not t['rows'] or not t['cols']:
            continue
        g = table_grid(t)
        head = [nospace(x[0]) for x in g[0]] + ([nospace(x[0]) for x in g[1]] if t['rows'] > 1 else [])
        alltxt = ' '.join(nospace(c['t']) for c in t['cells'])

        def col(name, rows=(0, 1)):
            # 머리글 칸만 본다 (본문이 긴 칸에 우연히 든 낱말은 무시: 유의사항 표 등)
            for r in rows:
                if r >= t['rows']:
                    continue
                for c in range(t['cols']):
                    h = nospace(g[r][c][0])
                    if name in h and len(h) <= len(name) + 14:
                        return c
            return None
        pcol = next((c for c in (col('공급금액'), col('공급가격'), col('금액')) if c is not None), None)
        # 납부계좌·일정 표
        if '납부계좌' in alltxt and ('추가선택품목' in alltxt or '발코니확장' in alltxt) \
                and not (out['account'] and out['date']):
            for r in range(t['rows']):
                for c in range(t['cols']):
                    txt = nospace(g[r][c][0])
                    if not out['account'] and re.fullmatch(r'\d[\d\-]{7,}\d', txt):
                        out['account'] = txt
                        bank = next((nospace(g[r][k][0]) for k in range(t['cols'])
                                     if re.search(r'(은행|농협|수협|금고|신협)$', nospace(g[r][k][0]))), '')
                        out['bank'] = bank
                        hol = next((g[r][k][0] for k in range(c + 1, t['cols']) if HOLDER_RE.search(g[r][k][0])), '')
                        out['holder'] = re.sub(r'(\s|⏎)+', ' ', hol).strip()
            mc = col('중도금', (0, 1, 2))
            if mc is not None and not out['date']:
                for r in range(t['rows']):
                    if ymd(g[r][mc][0]):
                        out['date'] = ymd(g[r][mc][0])
                        break
            continue
        if pcol is None:
            hr = next((r for r in range(min(2, t['rows']))
                       if any(TYPE_RE.fullmatch(nospace(x[0])) for x in g[r])), None)
            ctx = t['before'] + alltxt
            if hr is not None and '마이너스' not in ctx and re.search(r'옵션|선택품목|패키지', ctx):
                tcols = {c: nospace(g[hr][c][0]) for c in range(t['cols']) if TYPE_RE.fullmatch(nospace(g[hr][c][0]))}
                read_matrix(g, t, hr, tcols, out)
            continue
        if '설치내역' in head and any('미선택' in h for h in head):
            tcol = col('약식표기') if col('약식표기') is not None else col('주택형')
            start = 2 if t['rows'] > 1 and not money(g[1][pcol][0]) else 1
            for r in range(start, t['rows']):
                item_txt, item_k = g[r][1]
                brand_txt, brand_k = g[r][2]
                brand = brand_txt if brand_k != item_k else ''
                out['options'].append({'cat': g[r][0][0], 'item': item_txt, 'brand': nospace(brand),
                                       'types': types_in(g[r][tcol][0]), 'price': money(g[r][pcol][0]),
                                       'types_txt': g[r][tcol][0]})
        elif col('설치부위') is not None:
            # 칸 순서가 공고마다 달라 (등급 칸이 없기도 함) 내용으로 대수·등급 칸을 찾는다
            start = 2 if t['rows'] > 1 and not money(g[1][pcol][0]) else 1
            dcol = col('설치부위')
            tcol = next((c for c in (col('약식표기'), col('주택형'), col('타입')) if c is not None), 0)
            for r in range(start, t['rows']):
                cells = [x[0] for x in g[r]]
                count = next((x for c, x in enumerate(cells) if c != dcol and re.fullmatch(r'\d+대', nospace(x))), '')
                grade = next((grade_of(x) for c, x in enumerate(cells)
                              if c not in (dcol, tcol) and any(k in nospace(x) for k in GRADES)), '')
                out['aircon'].append({'types': types_in(g[r][tcol][0]), 'grade': grade, 'count': count,
                                      'desc': clean_desc(g[r][dcol][0]), 'price': money(g[r][pcol][0])})
        elif '패키지' in alltxt and col('품목') is not None:
            tcol = col('약식표기') if col('약식표기') is not None else 1
            start = 2 if t['rows'] > 1 and not money(g[1][pcol][0]) else 1
            for r in range(start, t['rows']):
                out['packages'].append({'name': nospace(g[r][0][0]), 'types': types_in(g[r][tcol][0]),
                                        'price': money(g[r][pcol][0])})
        elif ('발코니' in t['before'] or '발코니' in alltxt) and col('약식표기') is not None and '설치' not in alltxt:
            tcol = col('약식표기')
            ccol = col('주택형')
            start = 2 if t['rows'] > 1 and not re.search(r'\d', g[1][tcol][0]) else 1
            for r in range(start, t['rows']):
                ty = types_in(g[r][tcol][0])
                if len(ty) == 1 and '*' not in ty:
                    code = nospace(g[r][ccol][0]) if ccol is not None else ''
                    out['balcony'].append({'type': next(iter(ty)), 'code': re.sub(r'^0+(?=\d)', '', code),
                                           'price': money(g[r][pcol][0])})
            if not out['date']:
                mc = col('중도금', (0, 1))
                if mc is not None and t['rows'] > 1 and ymd(g[1][mc][0]):
                    out['date'] = ymd(g[1][mc][0])
    return out


# ------------------------------------------------------------------ 계약서 시트 읽기
SECTIONS = (('무상', 'free'), ('선택가능', 'free'), ('패키지', 'package'), ('에어컨', 'aircon'), ('추가선택', 'options'),
            ('발코니', 'balcony'))


def parse_sheet(sh):
    merged = sh.merged()
    owner = {}
    for (r1, c1), (r2, c2) in merged:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                owner[(r, c)] = (r1, c1)

    def val(r, c):
        r0, c0 = owner.get((r, c), (r, c))
        return sh.text(f'{num2col(c0)}{r0}')

    def ref(r, c):
        return f'{num2col(c)}{r}'
    maxr = sh.max_row()
    info = {'sections': {}, 'header_row': None}
    r = 1
    while r <= maxr:
        b = nospace(sh.text(f'B{r}'))
        kind = None
        if re.match(r'^\d\)', b):
            for key, k in SECTIONS:
                if key in b:
                    kind = k
                    break
        if kind and kind not in info['sections']:
            hr = r + 1
            heads = {nospace(val(hr, c)): c for c in range(2, 12)}
            pcol = next((c for h, c in heads.items() if '공급금액' in h), None)
            if pcol is None:        # 예전 양식: 금액 칸 머리글이 '계'
                pcol = next((c for h, c in heads.items() if h in ('계', '금액', '공급가격')), None)
            rows = []
            rr = hr + 1
            while rr <= maxr:
                bb = nospace(sh.text(f'B{rr}'))
                if re.match(r'^\d\)', bb) or bb.startswith('제') and '조' in bb[:4]:
                    break
                cells = [val(rr, c) for c in range(2, 11)]
                if not any(x.strip() for x in cells):
                    break
                rows.append(rr)
                rr += 1
            info['sections'][kind] = {'title_row': r, 'head_row': hr, 'pcol': pcol, 'rows': rows}
            if kind == 'balcony':
                info['header_row'] = hr
            r = rr
            continue
        r += 1
    info['val'] = val
    info['ref'] = ref
    info['owner'] = owner
    return info


# ------------------------------------------------------------------ 짝짓기
def match_option(name, brand, cands):
    k = okey(name)
    best, bs = None, 0
    for n in cands:
        nb = n['brand']
        if brand:
            if not nb or not (nb.startswith(brand) or brand.startswith(nb[:2])):
                continue
        elif nb:
            continue
        nk = okey(n['item'])
        if not nk:
            continue
        if nk == k:
            s = 3.0
        elif nk in k or k in nk:
            s = 2 + min(len(nk), len(k)) / max(len(nk), len(k))
        else:
            s = difflib.SequenceMatcher(None, k, nk).ratio()
            if s < 0.6:
                continue
        if s > bs:
            best, bs = n, s
    return best


def sheet_changes(sh, T, N, info, Ts=None):
    """시트 sh(이름 T) 를 공고 N 에 맞출 때 바뀌는 칸 목록.
    Ts: 시트가 맡는 주택형들 ('84A,B' 시트면 ['84A', '84B']). 없으면 [T]."""
    out, warns = [], []
    val = info['val']
    sec = info['sections']
    Ts = Ts or [T]

    def add(ref_, kind, label, old, new, default=True):
        out.append({'cell': ref_, 'kind': kind, 'label': label, 'old': old, 'new': new, 'default': default})

    def mine(x):
        return '*' in x['types'] or any(t in x['types'] for t in Ts)

    def split_price(same, label):
        """여러 타입을 한 시트에 둔 경우: 타입마다 공고 금액이 다르면 경고하고 True."""
        if len(Ts) < 2:
            return False
        per = {t: o['price'] for o in same for t in Ts if t in o['types'] and o['price']}
        missing = [t for t in Ts if t not in per and not any('*' in o['types'] for o in same)]
        if len(set(per.values())) > 1:
            warns.append(f'{T}: {label} 공고 금액이 타입마다 다릅니다 ('
                         + ', '.join(f'{t} {v:,}' for t, v in per.items()) + '). 시트를 타입별로 나눠야 할 수 있습니다.')
            return True
        if missing:
            warns.append(f'{T}: {label} 은(는) 공고상 {", ".join(missing)} 에는 없는 품목입니다.')
        return False

    # 시스템에어컨
    if 'aircon' in sec and N['aircon']:
        s = sec['aircon']
        cands = [a for a in N['aircon'] if mine(a)]
        if not cands:
            warns.append(f'{T}: 공고 시스템에어컨 표에 이 타입이 없습니다.')
        else:
            grade = ''
            groups = OrderedDict()
            gcol = col2num('J')
            for r in s['rows']:
                g = grade_of(val(r, gcol))
                if g:
                    grade = g
                groups.setdefault(grade, []).append(r)
            ng = OrderedDict()
            for a in cands:
                ng.setdefault(a['grade'], []).append(a)
            graded = any(ng)          # 공고에 일반형·고급형 같은 등급 칸이 있는가
            if len(groups) == 1 and '' in groups:
                pairs = [(r, a, a['grade']) for r, a in zip(groups[''], cands)]
                if len(groups['']) != len(cands):
                    warns.append(f'{T}: 시스템에어컨 줄 수({len(groups[""])})와 공고({len(cands)})가 다릅니다.')
            elif not graded and not any(g in GRADES for g in groups):
                # 샘플은 제조사별로 묶였고 공고는 등급 구분이 없음: 묶음마다 공고 금액표를 그대로 맞춘다
                pairs = []
                for g, rows in groups.items():
                    if len(rows) != len(cands):
                        warns.append(f'{T}: 시스템에어컨 {g} 줄 수({len(rows)})와 공고({len(cands)})가 다릅니다. 확인하세요.')
                    pairs += [(r, a, g) for r, a in zip(rows, cands)]
            elif not graded:
                pairs = []
                warns.append(f'{T}: 공고 시스템에어컨 표에는 등급(일반형·고급형) 구분이 없어 금액을 바꾸지 않았습니다. 직접 확인하세요.')
            else:
                pairs = []
                for g, rows in groups.items():
                    lst = ng.get(g, [])
                    if len(rows) != len(lst):
                        warns.append(f'{T}: 시스템에어컨 {g} 줄 수({len(rows)})와 공고({len(lst)})가 다릅니다. 확인하세요.')
                    pairs += [(r, a, a['grade']) for r, a in zip(rows, lst)]
            for r, a, g in pairs:
                name = ' '.join(x for x in ('시스템에어컨', g, nospace(a['count'])) if x)
                pref = f'{num2col(s["pcol"])}{r}'
                old = sh.value(pref)
                if a['price'] and old != a['price']:
                    add(pref, 'price', name, old, a['price'])
                dref = f'C{r}'
                if a['desc'] and nospace(val(r, 3)) != nospace(a['desc']):
                    add(dref, 'desc', f'{name} 설치부위', val(r, 3), a['desc'])
    # 유상옵션
    if 'options' in sec and N['options']:
        s = sec['options']
        cands = [o for o in N['options'] if mine(o)]
        cur_item = ''
        for r in s['rows']:
            item = val(r, 3)
            if item.strip():
                cur_item = item
            brand = nospace(val(r, 4))
            if brand == nospace(cur_item):
                brand = ''
            m = match_option(cur_item, brand, cands)
            if m is None and brand and not BRAND_RE.search(brand):
                m = match_option(cur_item, '', cands)     # '통합형/수납형' 같은 세부 구분 칸
            label = clean_desc(cur_item) + (f' ({brand})' if brand else '')
            if m is None:
                warns.append(f'{T}: 유상옵션 "{label}" 을(를) 공고에서 찾지 못했습니다 (이 타입에 없거나 이름이 다름).')
                continue
            same = [o for o in cands if okey(o['item']) == okey(m['item']) and o['brand'] == m['brand']]
            if split_price(same, f'유상옵션 "{label}"'):
                continue
            # 이름이 같거나 한쪽에 통째로 들면 그대로 적용, 비슷하기만 하면 기본 해제 (현장마다 품목 구성이 다름)
            a, b = okey(cur_item), okey(m['item'])
            sure = a == b or a in b or b in a
            if a != b:
                label += f' → 공고 "{clean_desc(m["item"])}"'
            pref = f'{num2col(s["pcol"])}{r}'
            old = sh.value(pref)
            if m['price'] and old != m['price']:
                add(pref, 'price', f'유상옵션 {label}', old, m['price'], sure)
    # 패키지
    if 'package' in sec and N['packages']:
        s = sec['package']
        for r in s['rows']:
            name = nospace(val(r, 2))
            same = [p for p in N['packages'] if (p['name'] == name or name in p['name'] or p['name'] in name) and mine(p)]
            if not same:
                warns.append(f'{T}: 패키지 "{name}" 을(를) 공고에서 찾지 못했습니다.')
                continue
            m = same[0]
            if split_price(same, f'패키지 "{name}"'):
                continue
            pref = f'{num2col(s["pcol"])}{r}'
            old = sh.value(pref)
            if m['price'] and old != m['price']:
                add(pref, 'price', f'패키지 {name}', old, m['price'])
    return out, warns


def balcony_cells(info, i=0):
    """발코니확장 표 i 번째 줄의 (주택형 칸, 금액 칸)."""
    s = info['sections'].get('balcony')
    if not s or len(s['rows']) <= i:
        return None, None
    r = s['rows'][i]
    return f'B{r}', f'{num2col(s["pcol"])}{r}'


def pay_table(sh, info):
    """예전 양식의 제2조 납부처 표 (납부기관 | 납부계좌 | 예금주) 아래 줄의 값 칸 {'bank'|'account'|'holder': 칸}."""
    val, owner = info['val'], info['owner']
    for r in range(1, sh.max_row()):
        heads = {}
        for c in range(2, 12):
            heads.setdefault(nospace(val(r, c)), c)
        if '납부기관' in heads and '납부계좌' in heads:
            out = {}
            for key, h in (('bank', '납부기관'), ('account', '납부계좌'), ('holder', '예금주')):
                if h in heads:
                    r0, c0 = owner.get((r + 1, heads[h]), (r + 1, heads[h]))
                    out[key] = f'{num2col(c0)}{r0}'
            return out
    return {}


def name_tail(name):
    m = re.search(r'린\s*(.+)$', name or '')
    return (m.group(1) if m else (name or '')).strip()


# ------------------------------------------------------------------ 분석
def analyze_options(xbytes, nbytes):
    try:
        book = Book(xbytes)
    except XlsxError as e:
        raise OptionError(str(e))
    N = read_notice_options(nbytes)
    if not (N['options'] or N['aircon'] or N['balcony']):
        raise OptionError('공고문에서 추가선택품목(발코니 확장·시스템에어컨·유상옵션) 표를 찾지 못했습니다.')
    sheets = [s for s in book.sheet_list() if s['state'] == 'visible']
    ntypes = []
    for b in N['balcony']:
        if b['type'] not in ntypes:
            ntypes.append(b['type'])
    if not ntypes:
        for a in N['aircon'] + N['options']:
            for t in sorted(a['types'] - {'*'}):
                if t not in ntypes:
                    ntypes.append(t)
    type_sheets = [s['name'] for s in sheets if s['name'] in ntypes or sheet_types(s['name'])]
    if not type_sheets:
        raise OptionError('주택형 이름(예: 84A, 84A,B)으로 된 시트를 찾지 못했습니다.')
    covers = {n: sheet_types(n) or [n] for n in type_sheets}     # 시트 → 맡는 주택형들
    covered = {t for ts in covers.values() for t in ts}
    warnings, changes = [], []
    infos = {}
    for name in type_sheets:
        infos[name] = parse_sheet(book.sheet(name))
        if 'balcony' not in infos[name]['sections']:
            warnings.append(f"'{name}' 시트에서 '1) 발코니확장' 표를 찾지 못했습니다.")

    # 새 타입 → 복사할 시트
    def pick_template(T):
        num = re.match(r'\d+', T).group(0)
        same = [s for s in type_sheets if re.match(r'\d+', s).group(0) == num]
        if same:
            return difflib.get_close_matches(T, same, 1, 0)[0] if len(same) > 1 else same[0]
        return difflib.get_close_matches(T, type_sheets, 1, 0)[0]
    new_sheets = [{'type': T, 'template': pick_template(T), 'apply': True} for T in ntypes if T not in covered]
    for name in type_sheets:
        if ntypes and not any(t in ntypes for t in covers[name]):
            warnings.append(f"'{name}' 시트의 타입이 공고에 없습니다. 시트를 지울지 확인하세요.")

    # 시트별 변경
    cid = [0]

    def push(sheet, ch, new_sheet=False):
        cid[0] += 1
        ch.update(id=f'o{cid[0]}', sheet=sheet, default=ch.get('default', True), new_sheet=new_sheet)
        changes.append(ch)
    targets = [(n, n, False) for n in type_sheets] + [(ns['type'], ns['template'], True) for ns in new_sheets]
    for T, src, is_new in targets:
        if ntypes and not any(t in ntypes for t in covers.get(T, [T])):
            continue            # 공고에 없는 타입 시트 (위에서 경고함)
        sh = book.sheet(src)
        chs, warns = sheet_changes(sh, T, N, infos[src], covers.get(T, [T]))
        warnings += warns
        for ch in chs:
            push(T, ch, is_new)

    # 발코니 확장 (공고 금액이 비어 있을 수 있어 입력칸으로 보여 준다). 여러 타입을 둔 시트는 타입마다 한 줄
    balcony = []
    nb = {b['type']: b for b in N['balcony']}
    blank = []
    for T, src, is_new in targets:
        sh = book.sheet(src)
        Ts = covers.get(T, [T])
        rows = infos[src]['sections'].get('balcony', {}).get('rows', [])
        for i, t in enumerate(Ts):
            if len(Ts) > 1:
                i = next((k for k, r in enumerate(rows) if nospace(sh.text(f'B{r}')) == t), i)
            cref, pref = balcony_cells(infos[src], i)
            old_code = sh.text(cref) if cref else ''
            old_price = sh.value(pref) if pref else None
            b = nb.get(t)
            nprice = b['price'] if b else None
            if b and not nprice:
                blank.append(t)
            # 샘플 칸이 '84A' 처럼 약식표기면 그대로 두고, 전용면적 주택형(84.xxxxA) 이면 공고 값으로
            code_new = b['code'] if b and b['code'] and not TYPE_RE.fullmatch(nospace(old_code)) else old_code
            balcony.append({'type': t, 'sheet': T, 'new_sheet': is_new, 'template': src if is_new else '',
                            'code_old': old_code if not is_new else '', 'code_new': code_new,
                            'old': old_price if not is_new else None, 'notice': nprice,
                            'value': nprice if nprice else (old_price if not is_new else None),
                            'cell_code': cref, 'cell_price': pref})
    if blank:
        warnings.insert(0, f"공고에 발코니 확장 금액이 비어 있습니다 ({', '.join(blank)}). 아래 '발코니 확장' 표에 금액을 직접 넣어 주세요."
                           " 비워 두면 샘플 금액을 그대로 둡니다.")

    # 1행 재산표시·납부일·계좌
    first = book.sheet(type_sheets[0])
    header_txt = ''
    for r in range(1, 6):
        t = first.text(f'B{r}')
        if re.search(r'[A-Za-z]{1,4}-?\d{1,3}\s*(블록|블럭|BL)', t):
            header_txt = t
            break
    rules = []
    nf = N['fields']
    m = re.search(r'([A-Za-z]{1,4}-?\d{1,3})\s*(블록|블럭|BL)', header_txt)
    if m and nf.get('block') and m.group(1) != nf['block']:
        rules.append({'old': m.group(1), 'new': nf['block'], 'apply': True})
    if header_txt and nf.get('complex_name'):
        rest = header_txt[m.end():] if m else header_txt
        old_name = re.split(r'\s{2,}', rest.strip())[0]
        ot, nt = name_tail(old_name), name_tail(nf['complex_name'])
        if old_name.endswith('린') and nt and nt != old_name:
            rules.append({'old': old_name, 'new': f'{old_name} {nt}', 'apply': True})   # '우미린' 뒤에 단지명이 없던 샘플
        elif ot and nt and ot != nt:
            rules.append({'old': ot, 'new': nt, 'apply': True})
    hr = infos[type_sheets[0]]['header_row'] or 7
    old_date = ''
    for c in range(2, 11):
        mm = re.search(r'중도금.*?\((\d{4}\.\d{1,2}\.\d{1,2})\)', first.text(f'{num2col(c)}{hr}').replace('\n', ''))
        if mm:
            old_date = mm.group(1)
    old_bank = old_acct = old_holder = ''
    for r in range(1, first.max_row() + 1):
        t = first.text(f'B{r}')
        mm = re.search(r'([가-힣A-Z]*(?:은행|농협|수협|금고|신협))\s*,\s*([\d]{2,}[\d\-]+\d)', t)
        if mm:
            old_bank, old_acct = mm.group(1), mm.group(2)
            h = re.search(r'예금주\s*[:：]\s*\n?\s*([^\]]+)\]', t)
            old_holder = h.group(1).strip() if h else ''
            break
    pay = {}            # 예전 양식: 제2조 납부처가 표 → 시트마다 값 칸에 직접 쓴다
    if not old_acct:
        for T, src, _ in targets:
            cells = pay_table(book.sheet(src), infos[src])
            for k, ref in cells.items():
                pay.setdefault(k, {})[T] = ref
        if pay:
            sh0 = book.sheet(type_sheets[0])
            got = {k: sh0.text(v[type_sheets[0]]).strip() for k, v in pay.items() if type_sheets[0] in v}
            old_bank, old_acct, old_holder = got.get('bank', ''), got.get('account', ''), got.get('holder', '')
    same = lambda a, b: nospace(a).replace('㈜', '(주)') == nospace(b).replace('㈜', '(주)')
    fields = [
        {'key': 'date', 'label': '중도금 납부일 (표 머리글)', 'old': old_date, 'new': N['date'], 'apply': bool(old_date and N['date'] and old_date != N['date'])},
        {'key': 'bank', 'label': '납부계좌 · 은행', 'old': old_bank, 'new': N['bank'], 'apply': bool(old_bank and N['bank'] and old_bank != N['bank'])},
        {'key': 'account', 'label': '납부계좌 · 계좌번호', 'old': old_acct, 'new': N['account'], 'apply': bool(old_acct and N['account'] and old_acct != N['account'])},
        {'key': 'holder', 'label': '납부계좌 · 예금주', 'old': old_holder, 'new': N['holder'], 'apply': bool(old_holder and N['holder'] and not same(old_holder, N['holder']))},
    ]
    for f in fields:
        if f['key'] in pay:
            f['cells'] = pay[f['key']]
    if not N['account']:
        warnings.append('공고에서 추가선택품목 납부계좌를 찾지 못했습니다. 직접 입력하세요.')
    if not old_acct:
        warnings.append('샘플 계약서에서 납부계좌 문장(○○은행, 계좌번호)을 찾지 못했습니다.')
    for r in rules:
        r['count'] = sum(1 for n in type_sheets for x in [book.sheet(n)] for ref in x.cells if r['old'] in x.text(ref))
    summary = {'sheets': len(type_sheets), 'new_sheets': len(new_sheets), 'changes': len(changes),
               'warnings': len(warnings)}
    return {'kind': 'options', 'type_sheets': type_sheets, 'new_sheets': new_sheets, 'balcony': balcony,
            'fields': fields, 'rules': rules, 'changes': changes, 'warnings': warnings, 'summary': summary,
            'notice': {k: v for k, v in nf.items() if k in ('complex_name', 'block', 'location')}}


# ------------------------------------------------------------------ 생성
def build_options(xbytes, plan, decisions, fields, rules, balcony, new_sheets, highlight=False):
    """fields: {key: 새 값 or None}, balcony: {type: {'value': 금액 or None, 'code': 주택형}},
    new_sheets: {type: {'apply': bool, 'template': 시트}}"""
    book = Book(xbytes)
    log = []
    created = []
    # 1) 새 시트 만들기
    for ns in plan['new_sheets']:
        d = new_sheets.get(ns['type'], {})
        if not d.get('apply', ns['apply']):
            continue
        tpl = d.get('template') or ns['template']
        if ns['type'] in [s['name'] for s in book.sheet_list()]:
            continue
        book.copy_sheet(tpl, ns['type'])
        created.append(ns['type'])
        log.append(('시트 추가', ns['type'], f'{tpl} 복사', ''))
    targets = [n for n in plan['type_sheets']] + created
    touched = {}

    def mark(sh, ref):
        if highlight:
            sh.highlight(ref)

    # 2) 칸별 변경
    for ch in plan['changes']:
        if not decisions.get(ch['id'], ch['default']):
            continue
        if ch['sheet'] not in targets:
            continue
        sh = book.sheet(ch['sheet'])
        if ch['kind'] == 'price':
            sh.set_number(ch['cell'], ch['new'])
        else:
            sh.set_text(ch['cell'], ch['new'])
        mark(sh, ch['cell'])
        log.append(('변경', ch['sheet'], f"{ch['label']}: {ch['old']}", str(ch['new'])))
    # 3) 발코니 확장
    for b in plan['balcony']:
        sheet = b.get('sheet') or b['type']
        if sheet not in targets:
            continue
        sh = book.sheet(sheet)
        d = balcony.get(b['type'], {})
        v = d.get('value', b['value'])
        code = d.get('code') or b['code_new']
        if v in (None, '') and b['new_sheet']:
            log.append(('확인 필요', b['type'], '발코니 확장 금액을 넣지 않아 복사한 시트 금액이 그대로 남았습니다',
                        str(sh.value(b['cell_price']) if b['cell_price'] else '')))
        if b['cell_price'] and v not in (None, '') and sh.value(b['cell_price']) != int(v):
            log.append(('발코니 확장', b['type'], str(sh.value(b['cell_price'])), str(int(v))))
            sh.set_number(b['cell_price'], int(v))
            mark(sh, b['cell_price'])
        if b['cell_code'] and code and sh.text(b['cell_code']) != code:
            log.append(('주택형', b['type'], sh.text(b['cell_code']), code))
            sh.set_text(b['cell_code'], code)
            mark(sh, b['cell_code'])
    # 4) 글자 바꾸기: 블록·단지명, 납부일, 계좌
    reps = []
    for r in rules:
        if r.get('apply') and r.get('old') and r.get('new'):
            reps.append((r['old'], r['new']))
    fmap = {f['key']: f for f in plan['fields']}
    for key in ('date', 'account', 'bank', 'holder'):
        v = fields.get(key)
        if key not in fmap or not v or not fmap[key]['old'] or v == fmap[key]['old']:
            continue
        if fmap[key].get('cells'):          # 납부처 표: 칸에 바로 쓴다 ('-' 같은 값은 글자 바꾸기로 못 찾음)
            for name, ref in fmap[key]['cells'].items():
                if name in targets:
                    sh = book.sheet(name)
                    sh.set_text(ref, v)
                    mark(sh, ref)
                    log.append(('변경', name, f"{fmap[key]['label']}: {fmap[key]['old']}", v))
        else:
            reps.append((fmap[key]['old'], v))
            alt = fmap[key]['old'].replace('(주)', '㈜') if '(주)' in fmap[key]['old'] else fmap[key]['old'].replace('㈜', '(주)')
            if key == 'holder' and alt != fmap[key]['old']:
                reps.append((alt, v))       # 같은 예금주를 '(주)' / '㈜' 로 섞어 쓴 샘플
    # 이름이 길어지면 '동  호' 칸 맞춤을 위해 뒤 공백을 줄인다
    for name in targets:
        sh = book.sheet(name)
        for ref in list(sh.cells):
            txt = sh.text(ref)
            if not txt or not isinstance(sh.value(ref), str):
                continue
            for old, new in reps:
                if old in txt:
                    if sh.replace_text(ref, old, new):
                        txt = sh.text(ref)
                        touched.setdefault(name, 0)
                        touched[name] += 1
                        mark(sh, ref)
                        if len(new) > len(old):
                            pad = ' ' * (len(new) - len(old))
                            idx = txt.find(new) + len(new)
                            if txt[idx:idx + len(pad) + 2].startswith(pad + ' '):
                                sh.replace_text(ref, new + pad, new)
                                txt = sh.text(ref)
    for old, new in reps:
        log.append(('글자 바꾸기', '전체 타입 시트', old, new))
    # 5) 수식 표시값 다시 계산
    for name in targets:
        book.sheet(name).recalc()
    return book.save(), log
