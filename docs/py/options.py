"""추가선택품목(발코니 확장·시스템에어컨·유상옵션) 계약서 엑셀을 새 현장 모집공고 기준으로 고친다.

샘플 계약서 = 주택형(타입)마다 시트 하나. 각 시트는
  1) 발코니확장  2) 시스템에어컨  3) 추가선택품목  4) 패키지 옵션  5) 무상선택  제2조 납부처 …
표 제목으로 구역을 찾는다. 시스템에어컨·추가선택품목·패키지 표는 공고에 그 타입으로 나온 줄만 공고 순서대로 다시 짠다
(공고에 없는 품목 줄은 지우고, 새 품목은 같은 모양의 줄을 본떠 넣는다. 서식·줄 높이·수식은 샘플 그대로).
품목 번호(①…)와 패키지의 항목 번호는 새 순서로 다시 매긴다.
공고에 있는데 시트가 없는 타입은 가장 비슷한 시트를 복사해 만들고, 공고에 없는 타입 시트는 지운다.
여러 타입을 한 시트에 둔 샘플('84A,B')은 공고 내용이 타입마다 다르면 타입별 시트로 나눈다.
"""
import difflib
import itertools
import re
from collections import OrderedDict

from hwp_reader import HwpError, read_hwp, read_hwp_tables, table_grid
from notice import extract_fields
from xlsxio import Book, XlsxError, col2num, num2col, split_ref

CIRCLED = '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳㉑㉒㉓㉔㉕㉖㉗㉘㉙㉚㉛㉜㉝㉞㉟㊱㊲㊳㊴㊵㊶㊷㊸㊹㊺㊻㊼㊽㊾㊿'
NUM_RE = re.compile(r'^\s*([' + CIRCLED + r']|\(\d+\))\s*')


class OptionError(Exception):
    pass


def nospace(s):
    return re.sub(r'\s+|⏎', '', s or '')


def money(s):
    d = re.sub(r'[^\d]', '', s or '')
    return int(d) if d else None


TYPE_RE = re.compile(r'\d{2,3}[A-Z]{1,3}')
TYPE_LOOSE = re.compile(r'\d{2,3}[A-Z]{0,3}')     # '72', '101' 처럼 알파벳 없는 타입도


def types_in(s):
    s = s or ''
    if re.search(r'전\s*(주택형|타입|세대)', s):
        return {'*'}
    return set(re.findall(r'(?<![\d.])\d{2,3}[A-Z]{0,3}(?![\d.])', s.replace('⏎', ' ')))


def strip_num(s):
    return NUM_RE.sub('', s or '')


def okey(s):
    """품목 이름 비교용: 번호·괄호·띄어쓰기·기호를 뺀다."""
    s = re.sub('[' + CIRCLED + ']', '', strip_num(s))
    s = re.sub(r'\([^)]*\)', '', s.replace('⏎', ' ').replace('\n', ' '))
    return re.sub(r'[\s+/·,.\-]', '', s)


def skey(s):
    """같은 글자인지 비교용: 번호·띄어쓰기·줄바꿈·'+'·쉼표만 무시한다."""
    return re.sub(r'[\s+·,⏎]', '', strip_num(s))


def clean_desc(s):
    s = (s or '').replace('⏎', ' ').strip()
    s = re.sub(r'^-\s*', '', s)
    return re.sub(r'\s{2,}', ' ', s).rstrip(',').strip()


def join_lines(s):
    """공고 칸 안에서 줄이 바뀐 이름을 한 줄로. '식기⏎세척기' → '식기세척기', '부부욕실⏎고급특화' → '부부욕실 고급특화'
    (줄바꿈 앞뒤에 띄어쓰기가 있으면 띄운다: '안방⏎ 붙박이장' → '안방 붙박이장')"""
    out, spaced = '', False
    for raw in (s or '').split('⏎'):
        p = raw.strip()
        if not p:
            spaced = True
            continue
        if not out:
            out, spaced = p, raw != raw.rstrip()
            continue
        a = out.split()[-1]
        b = p.split()[0]
        glue = not spaced and raw == raw.lstrip() and re.match(r'[가-힣]', a[-1]) and re.match(r'[가-힣]', b[0])             and (len(a) <= 2 or len(b) <= 2)
        out += p if glue else ' ' + p
        spaced = raw != raw.rstrip()
    return re.sub(r'\s{2,}', ' ', out).strip()


def note_text(s):
    return '\n'.join(x.strip() for x in (s or '').split('⏎') if x.strip())


GRADES = ('일반형', '고급형', '프리미엄형', '기본형')
DATE_RE = re.compile(r'(\d{4})\s*[.\-]\s*(\d{1,2})\s*[.\-]\s*(\d{1,2})')
HOLDER_RE = re.compile(r'신탁|\(주\)|㈜|주식회사|공사|금융센터')
# 가전 제조사 (옵션 표 마지막 구분 칸이 제조사인지 가린다)
BRAND_RE = re.compile(r'전자|매직|LG|삼성|SK|코웨이|쿠쿠|쿠첸|위니아|캐리어|린나이|경동|귀뚜라미|밀레|보쉬|일렉트로룩스')
HEADING_RE = re.compile(r'^([0-9]{1,2}|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫIVX]+)(-\d+)?\.?$')


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
    """시트 이름이 뜻하는 주택형 목록. '84A' → ['84A'], '84A,B' · '84A,84B' → ['84A', '84B'], '59B(수정)' → ['59B'], 그 밖 → []"""
    out, num = [], ''
    for p in re.split(r'[,·/]', re.sub(r'\([^)]*\)$', '', nospace(name))):
        m = re.fullmatch(r'(\d{2,3})?([A-Z]{0,3})', p)
        if not p or not m or not (m.group(1) or num):
            return []
        num = m.group(1) or num
        out.append(num + m.group(2))
    return out


def brand_core(s):
    return re.sub(r'\([^)]*\)|전자$', '', nospace(s))


def brand_like(a, b):
    """같은 회사 표기인가 ('삼성' ~ '삼성전자' ~ '삼성(12인용)')"""
    a, b = brand_core(a), brand_core(b)
    return bool(a and b and (a.startswith(b) or b.startswith(a)))


def brands_of(s):
    s = re.sub(r'제조사\s*[:：]?', '', (s or '').replace('⏎', ','))
    out = [x.strip() for x in re.split(r'[,/·]|및', s) if x.strip()]
    return [x for x in out if BRAND_RE.search(x)]


def type_cols(row):
    """머리글 줄에서 주택형 열 {열: 타입}. 타입이 둘 이상이거나 '84A' 꼴이 하나라도 있어야 한다."""
    cols = {c: nospace(x[0]) for c, x in enumerate(row) if TYPE_LOOSE.fullmatch(nospace(x[0]))}
    return cols if len(cols) >= 2 or any(TYPE_RE.fullmatch(v) for v in cols.values()) else {}


def fits_type(T, hint):
    """공고 괄호 속 타입 표시('84', '59A')가 타입 T 에 해당하는가"""
    return T == hint or (T.startswith(hint) and T[len(hint):].isalpha())


def read_matrix(g, t, hr, tcols, out):
    """주택형이 열 제목인 옵션·패키지 표 (예: 구분 | 구분 | 구분 | 84A | 84B | 비고) 를 타입별 줄로 푼다."""
    first, last = min(tcols), max(tcols)
    rows = []
    for r in range(hr + 1, t['rows']):
        labels = []
        for c in range(first):
            txt, k = g[r][c]
            if nospace(txt) and (not labels or labels[-1][1] != k):
                labels.append((txt, k))
        prices = {T: money(g[r][c][0]) for c, T in tcols.items()}
        note = next((g[r][c][0] for c in range(last + 1, t['cols']) if nospace(g[r][c][0])), '')
        if labels and any(v and v >= 10000 for v in prices.values()):
            rows.append(([x[0] for x in labels], prices, note))
    if not rows:
        return
    if sum(any(nospace(x).endswith('패키지') for x in labels) for labels, _, _ in rows) * 2 > len(rows):
        for labels, prices, note in rows:
            names = [nospace(x) for x in labels]
            name = next((n for n in names if n.endswith('패키지') and n != '패키지'), names[0])
            rest = [x for x, n in zip(labels, names) if n not in (name, '패키지')]
            comp = ' '.join(rest) if rest else note
            for T, v in prices.items():
                if v:
                    out['packages'].append({'name': name, 'types': {T}, 'price': v, 'items_txt': comp,
                                            'note': note if rest else ''})
        return
    for labels, prices, note in rows:
        L = list(labels)
        if len(L) >= 3 and nospace(L[0]) == '가전':
            brand = L.pop()           # 가전: 구분 | 품목 | 제조사(또는 설치 위치)
        elif len(L) >= 2 and BRAND_RE.search(nospace(L[-1])):
            brand = L.pop()
        else:
            brand = ''
        cat = L[-2] if len(L) >= 2 else ''
        for T, v in prices.items():
            if v:
                out['options'].append({'cat': cat, 'item': L[-1], 'brand': nospace(brand), 'types': {T},
                                       'price': v, 'note': note, 'types_txt': T})


# ------------------------------------------------------------------ 공고 읽기
def read_notice_options(nbytes):
    try:
        paras = read_hwp(nbytes)
        tables = read_hwp_tables(nbytes)
    except HwpError as e:
        raise OptionError(str(e))
    fields = extract_fields(paras)
    out = {'balcony': [], 'aircon': [], 'options': [], 'packages': [], 'date': '', 'bank': '', 'account': '',
           'holder': '', 'fields': fields, 'free': []}
    for p in paras:
        t = p['text'].strip()
        if '무상' in t and '선택' in t and ('발코니' in t or '확장' in t) and len(t) < 200 and t not in out['free']:
            out['free'].append(t)
    heading, prev = '', ''
    for t in tables:
        if not t['rows'] or not t['cols']:
            continue
        g = table_grid(t)
        alltxt = ' '.join(nospace(c['t']) for c in t['cells'])
        # 장 제목 표 ('Ⅷ-2 | | 추가 선택품목 계약') 를 따라가며 문맥으로 쓴다
        if t['cols'] <= 3 and HEADING_RE.match(nospace(g[0][0][0])):
            heading = nospace(g[0][-1][0])
        head = [nospace(x[0]) for x in g[0]] + ([nospace(x[0]) for x in g[1]] if t['rows'] > 1 else [])
        ctx = heading + nospace(t['before']) + alltxt

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

        def col_any(*names, rows=(0, 1)):
            return next((c for c in (col(n, rows) for n in names) if c is not None), None)
        pcol = col_any('공급금액', '공급가격', '공급가액', '금액')
        before_prev, prev = prev, alltxt
        # 납부계좌·일정 표
        if '납부계좌' in alltxt and ('추가선택품목' in alltxt or '발코니확장' in alltxt) \
                and not (out['account'] and out['date']):
            for r in range(t['rows']):
                for c in range(t['cols']):
                    txt = nospace(g[r][c][0])
                    am = re.fullmatch(r'(\d[\d\-]{7,}\d)(\([^)]*\))?', txt)      # '1010-2514-6879 (모계좌)'
                    if not out['account'] and am:
                        out['account'] = am.group(1)
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
        acol = col_any('설치부위', '설치위치')
        # 타입이 열인 옵션·패키지 표
        hr = next((r for r in range(min(2, t['rows'])) if type_cols(g[r])), None)
        if hr is not None and acol is None and re.search(r'옵션|선택품목|패키지', ctx) \
                and '마이너스' not in nospace(t['before']) + alltxt + before_prev:
            read_matrix(g, t, hr, type_cols(g[hr]), out)
            continue
        if pcol is None:
            continue
        start = 2 if t['rows'] > 1 and not money(g[1][pcol][0]) else 1
        if '설치내역' in head and any('미선택' in h for h in head):
            tcol = col('약식표기') if col('약식표기') is not None else col('주택형')
            ncol = col('설치내역')
            for r in range(start, t['rows']):
                item_txt, item_k = g[r][1]
                brand_txt, brand_k = g[r][2]
                brand = brand_txt if brand_k != item_k else ''
                out['options'].append({'cat': g[r][0][0], 'item': item_txt, 'brand': nospace(brand),
                                       'types': types_in(g[r][tcol][0]), 'price': money(g[r][pcol][0]),
                                       'note': clean_desc(g[r][ncol][0]) if ncol is not None else '',
                                       'types_txt': g[r][tcol][0]})
        elif acol is not None:
            # 칸 순서가 공고마다 달라 (등급 칸이 없기도 함) 내용으로 대수·등급·선택안·제조사 칸을 찾는다
            tcol = next((c for c in (col('약식표기'), col('주택형'), col('타입')) if c is not None), 0)
            bcol = col('제조사')
            for r in range(start, t['rows']):
                price = money(g[r][pcol][0])
                if not price:
                    continue
                cells = [x[0] for x in g[r]]
                other = [(c, x) for c, x in enumerate(cells) if c not in (acol, tcol, pcol)]
                count = next((nospace(x) for c, x in other if re.fullmatch(r'\d+대', nospace(x))), '')
                grade = next((grade_of(x) for c, x in other if any(k in nospace(x) for k in GRADES)), '')
                label = next((nospace(x) for c, x in other if re.fullmatch(r'선택\d+', nospace(x))), '')
                btxt = cells[bcol] if bcol is not None else next((x for c, x in other if BRAND_RE.search(x)), '')
                out['aircon'].append({'types': types_in(g[r][tcol][0]), 'grade': grade, 'count': count, 'label': label,
                                      'brands': brands_of(btxt) or ([nospace(btxt)] if nospace(btxt) else []),
                                      'desc': clean_desc(g[r][acol][0]), 'price': price})
        elif '패키지' in alltxt and col('품목') is not None:
            tcol = col('약식표기') if col('약식표기') is not None else 1
            icol = col('품목')
            for r in range(start, t['rows']):
                out['packages'].append({'name': nospace(g[r][0][0]), 'types': types_in(g[r][tcol][0]),
                                        'price': money(g[r][pcol][0]), 'items_txt': g[r][icol][0], 'note': ''})
        elif '발코니' in t['before'] + alltxt + heading and '설치' not in alltxt:
            tcol, ccol = col('약식표기'), col('주택형')
            if tcol is None and ccol is not None and t['rows'] > start and types_in(g[start][ccol][0]):
                tcol, ccol = ccol, None        # 주택형 칸에 약식표기('59A')를 바로 쓴 공고
            if tcol is None:
                continue
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


def for_type(items, t):
    return [x for x in items if x.get('price') and ('*' in x['types'] or t in x['types'])]


def notice_rows(N, kind, t):
    """공고에서 타입 t 에 해당하는 줄 (공고 순서). 같은 품목이 타입 묶음마다 반복되면 처음 것만."""
    if kind == 'options':
        out, seen = [], set()
        for o in for_type(N['options'], t):
            k = (nospace(o['item']), nospace(o['brand']))
            if k not in seen:
                seen.add(k)
                out.append(o)
        return out
    if kind == 'package':
        out, seen = [], set()
        for p in for_type(N['packages'], t):
            if p['name'] not in seen:
                seen.add(p['name'])
                out.append(p)
        return out
    rows = for_type(N['aircon'], t)
    # 제조사 여러 곳을 한 칸에 적은 공고('삼성, LG')는 제조사마다 같은 줄을 둔다
    brands = []
    for a in rows:
        for b in a['brands'] or ['']:
            if b not in brands:
                brands.append(b)
    out = []
    for b in brands:
        n = {}
        for a in rows:
            if b in (a['brands'] or ['']):
                n[a['grade']] = n.get(a['grade'], 0) + 1
                out.append(dict(a, brand=b, label=a['label'] or f'선택{n[a["grade"]]}'))
    return out


def rows_sig(kind, rows):
    if kind == 'options':
        return [(nospace(o['cat']), nospace(o['item']), o['brand'], o['price']) for o in rows]
    if kind == 'package':
        return [(p['name'], p['price']) for p in rows]
    return [(a['brand'], a['grade'], nospace(a['desc']), a['price']) for a in rows]


# ------------------------------------------------------------------ 계약서 시트 읽기
SECTIONS = (('무상', 'free'), ('선택가능', 'free'), ('패키지', 'package'), ('에어컨', 'aircon'), ('추가선택', 'options'),
            ('발코니', 'balcony'))
SECTION_LABEL = {'balcony': '발코니 확장', 'aircon': '시스템에어컨', 'options': '추가선택품목', 'package': '패키지'}


def is_title(b):
    return bool(re.match(r'^\d\)', b)) or (b.startswith('제') and '조' in b[:4])


def parse_sheet(sh):
    merged = sh.merged()
    owner, mext = {}, {}
    for (r1, c1), (r2, c2) in merged:
        mext[(r1, c1)] = (r2, c2)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                owner[(r, c)] = (r1, c1)

    def val(r, c):
        r0, c0 = owner.get((r, c), (r, c))
        return sh.text(f'{num2col(c0)}{r0}')

    def ref(r, c):
        return f'{num2col(c)}{r}'

    def blank(r):
        return not any(val(r, c).strip() for c in range(2, 11))

    def head_sig(r):
        return tuple(nospace(val(r, c)) for c in range(2, 11))
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
            heads = {}
            for c in range(2, 12):
                heads.setdefault(nospace(val(hr, c)), c)
            pcol = next((c for h, c in heads.items() if '공급금액' in h or '약정금액' in h), None)
            if pcol is None:        # 예전 양식: 금액 칸 머리글이 '계'
                pcol = next((c for h, c in heads.items() if h in ('계', '금액', '공급가격')), None)
            blocks, head, gap = [], hr, []
            while True:
                rows, rr = [], head + 1
                while rr <= maxr and not is_title(nospace(sh.text(f'B{rr}'))) and not blank(rr):
                    rows.append(rr)
                    rr += 1
                blocks.append({'gap': gap, 'head': head, 'rows': rows})
                # 빈 줄 뒤에 같은 머리글이 다시 나오면 쪽 나눔 때문에 둘로 나눈 같은 표
                nxt = rr
                while nxt <= maxr and nxt - rr < 2 and blank(nxt):
                    nxt += 1
                if rows and nxt > rr and nxt <= maxr and nospace(val(nxt, 2)) and head_sig(nxt) == head_sig(hr):
                    gap, head = list(range(rr, nxt)), nxt
                    continue
                break
            info['sections'][kind] = {'title_row': r, 'head_row': hr, 'pcol': pcol, 'ncol': heads.get('비고'),
                                      'rows': [x for bl in blocks for x in bl['rows']], 'blocks': blocks,
                                      'end': rr - 1}
            if kind == 'balcony':
                info['header_row'] = hr
            r = rr
            continue
        r += 1
    info['val'] = val
    info['ref'] = ref
    info['owner'] = owner
    info['mext'] = mext
    return info


def role_of(info, r, c):
    """세로 병합 안에서 칸의 자리: single / top / mid / bot"""
    o = info['owner'].get((r, c))
    if not o:
        return 'single'
    r2 = info['mext'][o][0]
    if o[0] == r2:
        return 'single'
    return 'top' if r == o[0] else ('bot' if r == r2 else 'mid')


# ------------------------------------------------------------------ 짝짓기
def name_score(a, b):
    if skey(a) == skey(b):
        return 4.0
    ka, kb = okey(a), okey(b)
    if not ka or not kb:
        return 0.0
    if ka == kb:
        s = 3.0
    elif ka in kb or kb in ka:
        s = 2 + min(len(ka), len(kb)) / max(len(ka), len(kb))
    else:
        s = difflib.SequenceMatcher(None, ka, kb).ratio()
        if s < 0.75:
            return 0.0
    pa = set(re.findall(r'[가-힣A-Za-z0-9]+', ' '.join(re.findall(r'\(([^)]*)\)', a or ''))))
    pb = set(re.findall(r'[가-힣A-Za-z0-9]+', ' '.join(re.findall(r'\(([^)]*)\)', b or ''))))
    return s + (0.3 if pa & pb else 0)


def pair_up(samples, wants, score):
    """샘플 줄과 공고 줄을 점수 높은 순으로 하나씩 짝짓는다. {공고 줄 번호: 샘플 줄 번호}"""
    pairs = sorted(((-score(s, w), i, j) for i, s in enumerate(samples) for j, w in enumerate(wants)))
    pairs = [(-sc, i, j) for sc, i, j in pairs]
    used_s, out = set(), {}
    for sc, i, j in pairs:
        if sc <= 0 or i in used_s or j in out:
            continue
        used_s.add(i)
        out[j] = i
    return out


def num_style(texts):
    """샘플 품목 번호 꼴: ('circ'|'paren'|'', 번호 뒤 띄어쓰기)"""
    kinds, seps = [], []
    for t in texts:
        m = NUM_RE.match(t or '')
        if m:
            kinds.append('circ' if m.group(1)[0] in CIRCLED else 'paren')
            seps.append(t[m.end(1):m.end()].replace('\n', ' ')[:1])
    if not kinds:
        return '', ''
    return max(set(kinds), key=kinds.count), max(set(seps), key=seps.count)


def num_text(style, i):
    if style == 'circ':
        return CIRCLED[i - 1] if i <= len(CIRCLED) else f'({i})'
    return f'({i})' if style == 'paren' else ''


def pkg_style(texts):
    """패키지 '항목' 칸 꼴: ('circ'|'paren'|'digit'|'names', 구분자)"""
    for t in texts:
        t = (t or '').strip()
        if not t:
            continue
        sep = ', ' if ', ' in t else ','
        if any(ch in CIRCLED for ch in t):
            return 'circ', sep
        if re.search(r'\(\d+\)', t):
            return 'paren', sep
        if re.fullmatch(r'\d+(\s*,\s*\d+)*', t):
            return 'digit', sep
        return 'names', ' + '
    return 'names', ' + '


def sample_numbers(texts):
    """'③, ④' · '(2),(3)' · '2,3,4' → [3, 4] …"""
    out = []
    for t in texts:
        for ch in t or '':
            if ch in CIRCLED:
                out.append(CIRCLED.index(ch) + 1)
        if not out:
            out = [int(x) for x in re.findall(r'\d+', t or '')]
    return out


# ------------------------------------------------------------------ 패키지 구성
RATES = (0.95, 0.9, 1.0, 0.97, 0.93, 0.92, 0.85, 0.8)


def fits_price(total, price):
    return any(abs(total * k - price) <= 10000 for k in RATES)


def comp_for_type(txt, T):
    """공고 패키지 구성 글 중 타입 T 에 해당하는 부분 ('[59] …, [74] …' 처럼 타입별로 적은 공고)"""
    txt = (txt or '').replace('⏎', ' ')
    segs = re.split(r'(\[[^\]]*\])', txt)
    if len(segs) == 1:
        return txt
    out, cur = [segs[0]], None
    for s in segs[1:]:
        if s.startswith('['):
            hints = types_in(s)
            cur = (not hints) or any(fits_type(T, h) for h in hints) or '*' in hints
            continue
        if cur:
            out.append(s)
    return ' '.join(out)


def comp_parts(txt, T):
    """패키지 구성 글을 낱 품목으로 나눈다. 괄호 속 타입 표시('주방장식장(84)')가 T 에 안 맞으면 뺀다."""
    parts, depth, cur = [], 0, ''
    for ch in comp_for_type(txt, T):
        if ch in '([':
            depth += 1
        elif ch in ')]':
            depth = max(0, depth - 1)
        if depth == 0 and ch in '+,':
            parts.append(cur)
            cur = ''
        else:
            cur += ch
    parts.append(cur)
    out = []
    for p in parts:
        p = re.sub(r'^[\s\-/]+', '', p).strip()
        if not p:
            continue
        hint = re.search(r'\(([^)]*)\)\s*$', p)
        if hint:
            hs = types_in(hint.group(1))
            if hs and '*' not in hs:
                if not any(fits_type(T, h) for h in hs):
                    continue
                p = p[:hint.start()].strip()
        out.append(p)
    return out


def resolve_package(pkg, T, items, sample_comp):
    """패키지에 드는 품목 번호들과 근거. items: [(번호, 이름, 금액, 구분)] (제조사 줄 없는 품목)."""
    price = pkg['price']

    def total(nums):
        return sum(next(p for n, _, p, _ in items if n == x) for x in nums)

    def by_text():
        nums = []
        for part in comp_parts(pkg.get('items_txt'), T):
            sc = sorted(((name_score(nm, part), n) for n, nm, _, _ in items), reverse=True)
            if not sc or sc[0][0] < 2 or (len(sc) > 1 and sc[1][0] == sc[0][0]):
                return None
            if sc[0][1] not in nums:
                nums.append(sc[0][1])
        return sorted(nums) if len(nums) >= 1 else None

    def by_sample():
        if not sample_comp:
            return None
        nums = []
        for nm in sample_comp:
            sc = sorted(((name_score(nm, x), n) for n, x, _, _ in items), reverse=True)
            if not sc or sc[0][0] < 2:
                return None
            nums.append(sc[0][1])
        return sorted(set(nums))

    def by_price():
        core = pkg['name'].replace('패키지', '')
        pool = [it for it in items if core and (core in nospace(it[3]) or core in nospace(it[1]))]
        for cand in (pool, items) if len(pool) >= 2 else (items,):
            if len(cand) > 14:
                continue
            for k in RATES:
                hits = [c for n in range(2, len(cand) + 1) for c in itertools.combinations(cand, n)
                        if abs(sum(x[2] for x in c) * k - price) <= 10000]
                if len(hits) == 1:
                    return sorted(x[0] for x in hits[0])
                if hits:
                    return None
        return None
    found = []
    for how, f in (('공고 구성', by_text), ('샘플 구성', by_sample), ('금액 검산', by_price)):
        nums = f()
        if nums:
            if fits_price(total(nums), price):
                return nums, how, True
            found.append((nums, how))
    if found:
        return found[0][0], found[0][1], False
    return None, '', False


# ------------------------------------------------------------------ 표 다시 짜기 계획
def plan_section(sh, info, kind, want, T, ctx):
    """샘플 시트 sh 의 kind 표를 공고 줄 want 로 다시 짤 계획. ctx: 같은 시트의 품목 번호 등 공유 정보."""
    sec = info['sections'][kind]
    val = info['val']
    pcol, ncol = sec['pcol'], sec['ncol']
    R = sec['rows']
    warns = []
    # 샘플 줄 정보
    S = []
    cur_cat = cur_item = ''
    for r in R:
        b, c = val(r, 2), val(r, 3)
        if b.strip():
            cur_cat = b
        if c.strip():
            cur_item = c
        own_c = info['owner'].get((r, 3), (r, 3))
        own_d = info['owner'].get((r, 4), (r, 4))
        d = val(r, 4) if own_d != own_c else ''
        S.append({'r': r, 'B': b, 'C': c, 'D': d, 'J': val(r, ncol) if ncol else '', 'cat': cur_cat, 'item': cur_item,
                  'kind': 'brand' if kind == 'options' and nospace(d) else 'plain',
                  'price': sh.value(f'{num2col(pcol)}{r}') if pcol else None})
    # 세로로 묶는 열: 샘플에서 세로 병합이 있던 열만
    vcols = sorted({c for (r1, c), (r2, c2) in info['mext'].items()
                    if r2 > r1 and R and R[0] <= r1 <= R[-1] and c in (2, 3, ncol or 0)})
    roles = {}
    for s in S:
        for c in vcols:
            st = sh.style(f'{num2col(c)}{s["r"]}')
            if st is not None:
                roles.setdefault(str(c), {}).setdefault(f'{role_of(info, s["r"], c)}|{s["kind"]}', st)
    hm = {}
    for (r1, c1), (r2, c2) in info['mext'].items():
        if r1 == r2 and r1 in R and c2 > c1:
            hm.setdefault(str(r1), []).append([c1, c2])
    notes = ncol and sum(1 for s in S if nospace(s['J'])) * 2 >= len(S) > 0
    sample_brands = [s['D'] for s in S if nospace(s['D'])] + [s['B'] for s in S] + [s['J'] for s in S]

    def brand_text(b):
        return next((nospace(x) for x in sample_brands if nospace(x) and brand_like(x, b)), b)

    rows = []         # [{'tpl', 'kind', 'vals': {열: 값}, 'group': {열: 묶음 열쇠}, …}]
    if kind == 'options':
        match = pair_up(S, want, lambda s, w: (name_score(s['item'], w['item']) if (
            (not nospace(s['D']) and not w['brand']) or (nospace(s['D']) and w['brand'] and brand_like(s['D'], w['brand']))
            or (nospace(s['D']) and w['brand'] and not BRAND_RE.search(w['brand']))) else 0))
        style, sep = num_style([s['C'] for s in S])
        has_brand_tpl = any(s['kind'] == 'brand' for s in S)
        no, last_key, numbers = 0, None, []
        for j, w in enumerate(want):
            name = join_lines(w['item'])
            i = match.get(j)
            sm = S[i] if i is not None else None
            if sm and skey(sm['item']) == skey(w['item']):
                name = strip_num(sm['item']).strip()        # 같은 글자면 샘플의 줄바꿈 모양을 살린다
            cat = join_lines(w['cat'])
            cat_s = next((s['cat'] for s in S if skey(s['cat']) == skey(cat)), None)
            if cat_s:
                cat = cat_s
            key = (nospace(w['cat']), okey(w['item']) + nospace(w['item']))
            if key != last_key:
                no += 1
                last_key = key
                if not w['brand']:      # 패키지 구성 찾기용 (가전처럼 제조사별 줄은 빼고)
                    numbers.append((no, name, w['price'], cat))
            rk = 'brand' if w['brand'] else 'plain'
            brand = brand_text(w['brand']) if w['brand'] else ''
            if rk == 'brand' and not has_brand_tpl:
                rk, name, brand = 'plain', f'{name} ({brand})', ''
            tpl = sm['r'] if sm and sm['kind'] == rk else next(
                (s['r'] for s in S if s['kind'] == rk and skey(s['cat']) == skey(cat)),
                next((s['r'] for s in S if s['kind'] == rk), S[0]['r']))
            nt = num_text(style, no)
            vals = {2: cat, 3: (nt + sep + name) if nt else name, 4: brand, pcol: w['price']}
            if notes:       # 같은 품목은 샘플에서 다듬어 쓴 비고를, 새 품목만 공고 비고를
                vals[ncol] = sm['J'] if sm else note_text(w.get('note'))
            elif ncol:
                vals[ncol] = ''
            status = 'new'
            if sm:
                status = 'same' if skey(sm['item']) == skey(name) and sm['price'] == w['price'] and \
                    nospace(sm['D']) == nospace(brand) else ('price' if skey(sm['item']) == skey(name) else 'text')
            alt = {}
            if status == 'text':    # 화면에서 '샘플 표기 유지'를 고르면 쓸 값
                old_nm = strip_num(sm['item']).strip()
                alt = {3: (nt + sep + old_nm) if nt else old_nm}
            rows.append({'tpl': tpl, 'kind': rk, 'vals': vals, 'status': status, 'alt': alt,
                         'group': {'2': nospace(cat), '3': str(no), str(ncol or 0): str(no)},
                         'show': {'no': nt, 'cat': cat, 'name': name, 'brand': brand, 'price': w['price'],
                                  'old': sm['price'] if sm else None, 'alt': bool(alt),
                                  'old_name': strip_num(sm['item']).strip() if sm and status == 'text' else ''}})
        ctx['numbers'] = numbers
        ctx['sample_numbers'] = {n: strip_num(s['item']).strip() for s in S
                                 for m in [NUM_RE.match(s['item'])] if m
                                 for n in [CIRCLED.index(m.group(1)) + 1 if m.group(1)[0] in CIRCLED
                                           else int(m.group(1)[1:-1])]}
        removed = [{'name': strip_num(s['item']).strip() + (f' ({s["D"]})' if nospace(s['D']) else ''), 'price': s['price']}
                   for i, s in enumerate(S) if i not in match.values()]
    elif kind == 'aircon':
        brole = 'keep'
        for s in S:
            t = nospace(s['B'])
            if re.fullmatch(r'선택\d+', t):
                brole = 'label'
            elif BRAND_RE.search(t):
                brole = 'brand'
            elif re.fullmatch(r'\d+대', t):
                brole = 'count'
            elif t in GRADES:
                brole = 'grade'
            if brole != 'keep':
                break
        jtexts = [nospace(s['J']) for s in S if nospace(s['J'])]
        jrole = 'brand' if any(BRAND_RE.search(x) for x in jtexts) else (
            'grade' if any(x in GRADES for x in jtexts) else 'none')
        match = pair_up(S, want, lambda s, w: (2 if nospace(s['C']) == nospace(w['desc']) else 0)
                        + (1 if grade_of(s['J']) == w['grade'] or not w['grade'] else 0)
                        + (1 if brand_like(s['B'] + s['J'], w['brand']) else 0)
                        + (0.5 if s['price'] == w['price'] else 0))
        for j, w in enumerate(want):
            i = match.get(j)
            desc = next((s['C'] for s in S if nospace(s['C']) == nospace(w['desc'])), w['desc'])
            brand = brand_text(w['brand']) if w['brand'] else ''
            bval = {'label': w['label'], 'brand': brand, 'count': w['count'], 'grade': w['grade']}.get(brole)
            vals = {3: desc, pcol: w['price']}
            group = {}
            if bval is not None:
                vals[2] = bval
                group['2'] = brand if brole == 'brand' else f'{j}'
            if ncol and jrole != 'none':
                vals[ncol] = brand if jrole == 'brand' else w['grade']
                group[str(ncol)] = brand if jrole == 'brand' else f'{brand}|{w["grade"]}'
            tpl = S[i]['r'] if i is not None else S[0]['r']
            status = 'new' if i is None else ('same' if nospace(S[i]['C']) == nospace(desc) and S[i]['price'] == w['price']
                                               else ('price' if nospace(S[i]['C']) == nospace(desc) else 'text'))
            alt = {3: S[i]['C']} if status == 'text' else {}
            rows.append({'tpl': tpl, 'kind': 'plain', 'vals': vals, 'group': group, 'status': status, 'alt': alt,
                         'show': {'no': w['label'], 'cat': ' '.join(x for x in (brand, w['grade']) if x),
                                  'name': desc, 'brand': w['count'], 'price': w['price'], 'alt': bool(alt),
                                  'old': S[i]['price'] if i is not None else None,
                                  'old_name': S[i]['C'] if i is not None and status == 'text' else ''}})
        removed = [{'name': s['C'], 'price': s['price']} for i, s in enumerate(S) if i not in match.values()]
    else:   # package
        style, sep = pkg_style([s['C'] for s in S])
        match = pair_up(S, want, lambda s, w: 4 if nospace(s['B']) == nospace(w['name']) else
                        (2 if okey(s['B']) == okey(w['name']) else 0))
        items = ctx.get('numbers') or []
        for j, w in enumerate(want):
            i = match.get(j)
            name = S[i]['B'] if i is not None and nospace(S[i]['B']) == nospace(w['name']) else join_lines(w['name'])
            sample_comp = None
            if i is not None and ctx.get('sample_numbers'):
                sample_comp = [ctx['sample_numbers'][n] for n in sample_numbers([S[i]['C']])
                               if n in ctx['sample_numbers']] or None
            nums, how, ok = resolve_package(w, T, items, sample_comp) if items else (None, '', False)
            if style == 'names':
                parts = comp_parts(w.get('items_txt'), T)
                comp = sep.join(parts) if parts else sep.join(nm for n, nm, _, _ in items if nums and n in nums)
            elif nums:
                comp = sep.join(num_text('circ' if style == 'circ' else 'paren', n) if style != 'digit' else str(n)
                                for n in nums)
            else:
                comp = S[i]['C'] if i is not None else ''
            if style != 'names':
                if not nums:
                    warns.append(f'{T}: 패키지 "{name}" 에 드는 품목을 공고에서 가리지 못했습니다. 항목 칸을 확인하세요.')
                elif not ok:
                    warns.append(f'{T}: 패키지 "{name}" 항목({comp})의 금액 합계가 공고 패키지 금액과 맞지 않습니다. 확인하세요.')
            vals = {2: name, 3: comp, pcol: w['price']}
            if ncol:
                vals[ncol] = S[i]['J'] if i is not None else (note_text(w.get('note')) if notes else '')
            tpl = S[i]['r'] if i is not None else S[0]['r']
            status = 'new' if i is None else ('same' if nospace(S[i]['C']) == nospace(comp) and S[i]['price'] == w['price']
                                               else ('price' if nospace(S[i]['C']) == nospace(comp) else 'text'))
            rows.append({'tpl': tpl, 'kind': 'plain', 'vals': vals, 'group': {}, 'status': status,
                         'show': {'no': '', 'cat': '', 'name': name, 'brand': comp, 'price': w['price'],
                                  'old': S[i]['price'] if i is not None else None, 'old_name': '',
                                  'old_brand': S[i]['C'] if i is not None and status == 'text' else ''}})
        removed = [{'name': s['B'], 'price': s['price']} for i, s in enumerate(S) if i not in match.values()]
    # 여러 덩이(쪽 나눔)로 된 표: 뒤 덩이 첫 구분('가전' 등)부터 뒤 덩이로
    assign = [[] for _ in sec['blocks']]
    k = 0
    for i, row in enumerate(rows):
        while k + 1 < len(sec['blocks']):
            nb = sec['blocks'][k + 1]
            first = val(nb['rows'][0], 2) if nb['rows'] else ''
            if nospace(first) and skey(row['vals'].get(2, '')) == skey(first):
                k += 1
                continue
            break
        assign[k].append(i)
    spec = {'kind': kind, 'title_row': sec['title_row'], 'end': sec['end'], 'blocks': sec['blocks'], 'assign': assign,
            'rows': rows, 'vcols': vcols, 'roles': roles, 'hm': hm, 'pcol': pcol}
    return spec, removed, warns


def balcony_trim_spec(info, keep_rows):
    """여러 타입 시트를 나눌 때: 발코니 확장 표에서 keep_rows 줄만 남긴다."""
    sec = info['sections']['balcony']
    hm = {}
    for (r1, c1), (r2, c2) in info['mext'].items():
        if r1 == r2 and r1 in sec['rows'] and c2 > c1:
            hm.setdefault(str(r1), []).append([c1, c2])
    rows = [{'tpl': r, 'kind': 'plain', 'vals': {}, 'group': {}, 'status': 'same'} for r in keep_rows]
    return {'kind': 'balcony', 'title_row': sec['title_row'], 'end': sec['end'], 'blocks': sec['blocks'],
            'assign': [list(range(len(rows)))] + [[] for _ in sec['blocks'][1:]], 'rows': rows, 'vcols': [],
            'roles': {}, 'hm': hm, 'pcol': sec['pcol']}


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
        for a in N['aircon'] + N['options'] + N['packages']:
            for t in sorted(a['types'] - {'*'}):
                if t not in ntypes:
                    ntypes.append(t)
    type_sheets = [s['name'] for s in sheets if s['name'] in ntypes or sheet_types(s['name'])]
    if not type_sheets:
        raise OptionError('주택형 이름(예: 84A, 84A,B)으로 된 시트를 찾지 못했습니다.')
    covers = {n: sheet_types(n) or [n] for n in type_sheets}     # 시트 → 맡는 주택형들
    warnings, changes = [], []
    infos = {name: parse_sheet(book.sheet(name)) for name in type_sheets}
    for name in type_sheets:
        if 'balcony' not in infos[name]['sections']:
            warnings.append(f"'{name}' 시트에서 '1) 발코니확장' 표를 찾지 못했습니다.")
    have = {'aircon': bool(N['aircon']), 'options': bool(N['options']), 'package': bool(N['packages'])}
    for k, lab in (('aircon', '시스템에어컨'), ('options', '추가선택품목(유상옵션)'), ('package', '패키지 옵션')):
        if not have[k] and any(k in i['sections'] for i in infos.values()):
            warnings.append(f'공고에서 {lab} 표를 찾지 못해 샘플의 {lab} 표를 그대로 두었습니다. 확인하세요.')

    def want(kind, t):
        return notice_rows(N, kind, t)

    # 여러 타입 시트: 공고 내용이 타입마다 다르면 타입별 시트로 나눈다
    split = {}
    for name in type_sheets:
        ts = [t for t in covers[name] if t in ntypes]
        if len(ts) > 1 and any(rows_sig(k, want(k, ts[0])) != rows_sig(k, want(k, t))
                               for k in ('aircon', 'options', 'package') for t in ts[1:]):
            split[name] = ts
    covered = {t for n, ts in covers.items() if n not in split for t in ts}

    # 새 타입 → 복사할 시트
    def pick_template(T):
        for n, ts in split.items():
            if T in ts:
                return n
        num = re.match(r'\d+', T).group(0)
        same = [s for s in type_sheets if re.match(r'\d+', s).group(0) == num]
        if same:
            return difflib.get_close_matches(T, same, 1, 0)[0] if len(same) > 1 else same[0]
        return difflib.get_close_matches(T, type_sheets, 1, 0)[0]
    new_sheets = [{'type': T, 'template': pick_template(T), 'apply': True,
                   'split': any(T in ts for ts in split.values())} for T in ntypes if T not in covered]
    delete_sheets = []
    seen = set()
    for name in type_sheets:
        if name not in split and covers[name] and set(covers[name]) <= seen:
            delete_sheets.append({'sheet': name, 'reason': '같은 타입 시트가 앞에 있음', 'apply': True})
            continue
        seen |= set(covers[name])
        if name in split:
            delete_sheets.append({'sheet': name, 'reason': f'타입별 시트({", ".join(split[name])})로 나눔', 'apply': True})
            warnings.append(f"'{name}' 시트는 공고상 타입마다 품목·금액이 달라 타입별 시트로 나눕니다.")
        elif ntypes and not any(t in ntypes for t in covers[name]):
            delete_sheets.append({'sheet': name, 'reason': '공고에 없는 타입', 'apply': True})

    # 시트별 표 다시 짜기
    cid = [0]
    sec_specs = {}

    def push(ch):
        cid[0] += 1
        ch['id'] = f'o{cid[0]}'
        changes.append(ch)
    targets = [(n, n, False, [t for t in covers[n] if t in ntypes] or covers[n]) for n in type_sheets
               if n not in split and not any(d['sheet'] == n for d in delete_sheets)]
    targets += [(ns['type'], ns['template'], True, [ns['type']]) for ns in new_sheets]
    for T, src, is_new, Ts in targets:
        info = infos[src]
        sh = book.sheet(src)
        ctx = {}
        for kind in ('aircon', 'options', 'package'):
            if kind not in info['sections'] or not have[kind]:
                continue
            if not info['sections'][kind]['rows']:
                continue
            rows = want(kind, Ts[0])
            if not rows:
                sec = info['sections'][kind]
                push({'sheet': T, 'new_sheet': is_new, 'section': kind, 'label': f'{SECTION_LABEL[kind]} 표 삭제',
                      'default': True, 'summary': {'before': len(sec['rows']), 'after': 0, 'add': 0,
                                                   'del': len(sec['rows']), 'price': 0, 'text': 0},
                      'rows': [], 'removed': [], 'changed': True,
                      'spec': {'drop': True, 'title_row': sec['title_row'], 'end': sec['end']}})
                warnings.append(f'{T}: 공고에 이 타입의 {SECTION_LABEL[kind]} 이(가) 없어 샘플의 {SECTION_LABEL[kind]} 표를 지웁니다.')
                continue
            spec, removed, warns = plan_section(sh, info, kind, rows, Ts[0], ctx)
            warnings += warns
            summ = {'before': len(info['sections'][kind]['rows']), 'after': len(spec['rows']),
                    'add': sum(r['status'] == 'new' for r in spec['rows']), 'del': len(removed),
                    'price': sum(r['status'] == 'price' for r in spec['rows']),
                    'text': sum(r['status'] == 'text' for r in spec['rows'])}
            changed = summ['add'] or summ['del'] or summ['price'] or summ['text'] or is_new
            push({'sheet': T, 'new_sheet': is_new, 'section': kind, 'label': SECTION_LABEL[kind], 'default': True,
                  'summary': summ, 'rows': [dict(r['show'], status=r['status']) for r in spec['rows']],
                  'removed': removed, 'spec': spec, 'changed': bool(changed)})
        for k in ('aircon', 'options', 'package'):
            if have[k] and k not in info['sections'] and want(k, Ts[0]):
                warnings.append(f'{T}: 샘플 시트에 {SECTION_LABEL[k]} 표가 없어 공고의 {SECTION_LABEL[k]} 을(를) 넣지 못했습니다.')
        sec_specs[T] = src
    if N['free'] and any('free' in i['sections'] for i in infos.values()):
        warnings.append('무상 선택품목은 공고에 문장으로만 있어 자동으로 고치지 않았습니다. 샘플의 무상선택 표를 확인하세요: '
                        + ' / '.join(N['free'][:3]))

    # 발코니 확장 (공고 금액이 비어 있을 수 있어 입력칸으로 보여 준다). 여러 타입을 둔 시트는 타입마다 한 줄
    balcony = []
    nb = {b['type']: b for b in N['balcony']}
    blank = []
    for T, src, is_new, Ts in targets:
        sh = book.sheet(src)
        rows = infos[src]['sections'].get('balcony', {}).get('rows', [])
        for i, t in enumerate(Ts):
            if len(rows) > 1:
                i = next((k for k, r in enumerate(rows) if nospace(sh.text(f'B{r}')) == t),
                         next((k for k, r in enumerate(rows) if t in types_in(sh.text(f'B{r}'))), i))
            cref, pref = balcony_cells(infos[src], i)
            old_code = sh.text(cref) if cref else ''
            old_price = sh.value(pref) if pref else None
            b = nb.get(t)
            nprice = b['price'] if b else None
            if b and not nprice:
                blank.append(t)
            # 샘플 칸이 '84A' 처럼 약식표기면 타입 이름, 전용면적 주택형(84.xxxxA) 이면 공고 값으로
            if TYPE_LOOSE.fullmatch(nospace(old_code) or '-'):
                code_new = t
            elif b and b['code']:
                code_new = b['code']
            else:
                code_new = t if is_new or not old_code else old_code
            balcony.append({'type': t, 'sheet': T, 'new_sheet': is_new, 'template': src if is_new else '',
                            'code_old': old_code if not is_new else '', 'code_new': code_new,
                            'old': old_price if not is_new else None, 'notice': nprice,
                            'value': nprice if nprice else (old_price if not is_new else None),
                            'cell_code': cref, 'cell_price': pref})
            # 나눠 만든 시트: 발코니 표에서 다른 타입 줄을 지운다
            if is_new and len(rows) > 1 and cref:
                spec = balcony_trim_spec(infos[src], [int(cref[1:])])
                push({'sheet': T, 'new_sheet': True, 'section': 'balcony', 'label': '발코니 확장 (다른 타입 줄 지움)',
                      'default': True, 'summary': {'before': len(rows), 'after': 1, 'add': 0, 'del': len(rows) - 1,
                                                   'price': 0, 'text': 0},
                      'rows': [], 'removed': [], 'spec': spec, 'changed': True, 'hidden': True})
    if blank:
        warnings.insert(0, f"공고에 발코니 확장 금액이 비어 있습니다 ({', '.join(blank)}). 아래 '시트 · 발코니 확장' 표에 금액을 직접"
                           " 넣어 주세요. 비워 두면 샘플 금액을 그대로 둡니다.")

    # 1행 재산표시·납부일·계좌
    first = book.sheet(type_sheets[0])
    header_txt = ''
    for r in range(1, 6):
        t = first.text(f'B{r}')
        if t.strip().startswith('■') or re.search(r'[A-Za-z]{1,4}-?\d{1,3}\s*(블록|블럭|BL)', t):
            header_txt = t
            break
    rules = []
    nf = N['fields']
    # '■ (주소) 단지명      동    호' 의 앞부분을 공고의 위치·단지명으로 통째로
    mh = re.match(r'\s*■\s*(.*?)(?:\s{2,}|$)', header_txt)
    old_head = mh.group(1).strip() if mh else ''
    cname = re.sub(r'\s*(본청약|사전청약)$', '', nf.get('complex_name') or '').strip()
    if old_head and cname:
        has_addr = bool(re.search(r'[가-힣]+(도|시|군)\s|번지|일원|지구\s*내', old_head))
        new_head = f"{nf['location']} {cname}" if has_addr and nf.get('location') else cname
        if nospace(new_head) != nospace(old_head):
            rules.append({'old': old_head, 'new': new_head, 'apply': True})
        header_txt = ''         # 아래 블록·단지명 낱 규칙은 쓰지 않는다
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
        for T, src, _, _ in targets:
            cells = pay_table(book.sheet(src), infos[src])
            for k, ref in cells.items():
                pay.setdefault(k, {})[T] = ref
        if pay:
            sh0 = book.sheet(type_sheets[0])
            info0 = parse_sheet(sh0)
            got = {k: v.strip() for k, v in
                   ((k, sh0.text(ref)) for k, ref in pay_table(sh0, info0).items())}
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
            # 납부처 표는 샘플 칸이 비어 있어도 공고 값을 넣는다
            f['apply'] = bool(f['new'] and not same(f['old'], f['new']))
    if not N['account']:
        warnings.append('공고에서 추가선택품목 납부계좌를 찾지 못했습니다. 직접 입력하세요.')
    if not old_acct and not pay:
        warnings.append('샘플 계약서에서 납부계좌 문장(○○은행, 계좌번호)을 찾지 못했습니다.')
    for r in rules:
        r['count'] = sum(1 for n in type_sheets for x in [book.sheet(n)] for ref in x.cells if r['old'] in x.text(ref))
    # '59B(수정)' 처럼 타입 뒤에 꼬리가 붙은 시트는 타입 이름으로
    gone = {d['sheet'] for d in delete_sheets}
    names = {s['name'] for s in sheets}
    hidden = [s['name'] for s in book.sheet_list() if s['state'] != 'visible']
    for h in hidden:        # 숨긴 시트(예전 판·시청 제출용 등)는 기본으로 두되, 지울 수 있게 보여 준다
        delete_sheets.append({'sheet': h, 'reason': '숨긴 시트', 'apply': False})
    renames = [{'sheet': n, 'new': covers[n][0], 'apply': True, 'hidden_clash': covers[n][0] in hidden} for n in type_sheets
               if n not in gone and len(covers[n]) == 1 and n != covers[n][0] and covers[n][0] not in names
               and covers[n][0] not in [ns['type'] for ns in new_sheets]]
    visible = [c for c in changes if not c.get('hidden')]
    summary = {'sheets': len(type_sheets), 'new_sheets': len(new_sheets), 'changes': sum(c['changed'] for c in visible),
               'delete_sheets': len(delete_sheets), 'warnings': len(warnings)}
    return {'kind': 'options', 'type_sheets': type_sheets, 'new_sheets': new_sheets, 'delete_sheets': delete_sheets,
            'renames': renames,
            'balcony': balcony, 'fields': fields, 'rules': rules, 'changes': changes, 'warnings': warnings,
            'summary': summary, 'ntypes': ntypes,
            'notice': {k: v for k, v in nf.items() if k in ('complex_name', 'block', 'location')}}


# ------------------------------------------------------------------ 생성
def rebuild_section(sh, spec, mark, log):
    """spec 대로 표 줄을 다시 짠다. 아래 덩이부터 고쳐서 위 덩이의 줄 번호가 흔들리지 않게 한다."""
    rows, blocks = spec['rows'], spec['blocks']
    sh.unshare()            # 공유수식을 먼저 풀어야 줄을 본떠도 수식 범위가 어긋나지 않는다
    snaps = {r['tpl']: sh.row_copy(r['tpl']) for r in rows}
    roles = spec['roles']
    vcols = spec['vcols']
    if not any(spec['assign']):
        return
    for k in range(len(blocks) - 1, -1, -1):
        bl = blocks[k]
        R = bl['rows']
        D = [rows[i] for i in spec['assign'][k]]
        if not R:
            continue
        if not D:
            top = (bl['gap'] or [bl['head']])[0]
            sh.delete_rows(top, R[-1] - top + 1)       # 쓸 줄이 없는 뒤 덩이는 빈 줄·머리글째 지운다
            continue
        for (r1, c1), (r2, c2) in sh.merged():
            if R[0] <= r1 <= R[-1]:
                sh.unmerge(f'{num2col(c1)}{r1}:{num2col(c2)}{r2}')
        n, m = len(R), len(D)
        if m > n:
            sh.insert_rows(R[-1] + 1, m - n)
        elif m < n:
            sh.delete_rows(R[0] + m, n - m)
        r0 = R[0]
        for i, d in enumerate(D):
            sh.put_row(snaps[d['tpl']], r0 + i)
        marks = []
        for i, d in enumerate(D):
            rr = r0 + i
            # 인쇄영역 밖(L열~) 메모 글자는 옛 품목 설명이라 지운다 (검산 수식·숫자는 둔다)
            for ref in [x for x in sh.cells if split_ref(x)[1] == rr and col2num(split_ref(x)[0]) > spec.get('last_col', 11)]:
                if isinstance(sh.value(ref), str) and sh.formula(ref) is None:
                    sh.clear(ref)
            for c, v in d['vals'].items():
                if not c:
                    continue
                ref = f'{num2col(int(c))}{rr}'
                old = sh.value(ref)
                if v is None or v == '':
                    sh.clear(ref)
                    continue
                if isinstance(v, (int, float)):
                    sh.set_number(ref, v)
                else:
                    sh.set_text(ref, v)
                if d['status'] == 'new' or (old != v and d['status'] != 'same'):
                    marks.append(ref)
            if spec.get('pcol'):
                for c in range(spec['pcol'], spec['pcol'] + 4):      # 금액·계약금·중도금·잔금
                    sh.shrink(f'{num2col(c)}{rr}')
            for c1, c2 in spec['hm'].get(str(d['tpl']), []):
                sh.merge(f'{num2col(c1)}{rr}:{num2col(c2)}{rr}')
        # 세로 묶음 (구분·품목·제조사 등): 같은 열쇠가 이어지면 첫 줄에만 글자, 병합
        for c in vcols:
            key = str(c)
            i = 0
            while i < len(D):
                g = D[i]['group'].get(key)
                j = i
                while j + 1 < len(D) and g is not None and D[j + 1]['group'].get(key) == g:
                    j += 1
                for x in range(i, j + 1):
                    rr = r0 + x
                    ref = f'{num2col(c)}{rr}'
                    role = 'single' if i == j else ('top' if x == i else ('bot' if x == j else 'mid'))
                    rs = roles.get(key, {})
                    st = rs.get(f'{role}|{D[x]["kind"]}', next((v for kk, v in rs.items() if kk.startswith(role + '|')), None))
                    if st is not None and not sh.style(ref) == st:
                        sh.set_style(ref, st)
                    if x > i:
                        sh.clear(ref)
                if j > i:
                    # 이 열에 걸린 가로 병합(C:D 등)이 있으면 풀고 세로로 묶는다
                    for x in range(i, j + 1):
                        for (r1, c1), (r2, c2) in sh.merged():
                            if r1 == r0 + x and c1 <= c <= c2 and r1 == r2:
                                sh.unmerge(f'{num2col(c1)}{r1}:{num2col(c2)}{r2}')
                    sh.merge(f'{num2col(c)}{r0 + i}:{num2col(c)}{r0 + j}')
                i = j + 1
        for ref in marks:           # 묶음 서식을 다 입힌 뒤에 노란색 표시
            if sh.value(ref) not in (None, ''):
                mark(sh, ref)


def drop_section(sh, spec):
    """표 하나를 통째로 지우고 (제목 줄부터 다음 빈 줄까지) 뒤 표의 번호를 하나씩 당긴다."""
    top, end = spec['title_row'], spec['end']
    if not any(sh.text(f'{num2col(c)}{end + 1}').strip() for c in range(2, 11)):
        end += 1
    sh.delete_rows(top, end - top + 1)
    for r in range(top, sh.max_row() + 1):
        t = sh.text(f'B{r}')
        m = re.match(r'^\s*(\d)\)', t)
        if m:
            sh.replace_text(f'B{r}', m.group(1) + ')', f'{int(m.group(1)) - 1})')
        elif t.strip().startswith('제') and '조' in t[:6]:
            break


def build_options(xbytes, plan, decisions, fields, rules, balcony, new_sheets, highlight=False, delete_sheets=None,
                  renames=None, keep_names=None):
    """fields: {key: 새 값 or None}, balcony: {type: {'value': 금액 or None, 'code': 주택형}},
    new_sheets: {type: {'apply': bool, 'template': 시트}}, delete_sheets: {시트: bool}, renames: {시트: bool},
    keep_names: {변경 id: [샘플 표기를 살릴 줄 번호]}"""
    book = Book(xbytes)
    log = []
    created = []
    delete_sheets = delete_sheets or {}
    renames = renames or {}
    keep_names = keep_names or {}
    # 1) 새 시트 만들기 (공고 타입 순서대로, 앞 타입 시트 바로 뒤에)
    order = plan.get('ntypes') or [ns['type'] for ns in plan['new_sheets']]
    for ns in sorted(plan['new_sheets'], key=lambda x: order.index(x['type']) if x['type'] in order else 999):
        d = new_sheets.get(ns['type'], {})
        if not d.get('apply', ns['apply']):
            continue
        tpl = d.get('template') or ns['template']
        names = [s['name'] for s in book.sheet_list()]
        if ns['type'] in names:
            continue
        i = order.index(ns['type']) if ns['type'] in order else -1

        def sheet_of(t):
            if t in names:
                return t
            return next((n for n in plan['type_sheets'] if n in names and t in (sheet_types(n) or [n])), None)
        prev = next((sheet_of(t) for t in reversed(order[:i]) if sheet_of(t)), None) if i > 0 else None
        nxt = next((sheet_of(t) for t in order[i + 1:] if sheet_of(t)), None) if i >= 0 else None
        if prev is None and ns.get('split'):
            prev = tpl                                  # 나눈 시트는 원래 시트 자리에
        if prev is None and nxt is not None:
            book.copy_sheet(tpl, ns['type'], before=nxt)
        else:
            book.copy_sheet(tpl, ns['type'], after=prev or tpl)
        created.append(ns['type'])
        log.append(('시트 추가', ns['type'], f'{tpl} 복사', ''))
    targets = [n for n in plan['type_sheets']] + created
    touched = {}

    def mark(sh, ref):
        if highlight:
            sh.highlight(ref)

    # 2) 발코니 확장
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
            c0, r0 = split_ref(b['cell_price'])
            for c in range(col2num(c0), col2num(c0) + 4):           # 금액·계약금·중도금·잔금
                sh.shrink(f'{num2col(c)}{r0}')
            mark(sh, b['cell_price'])
        if b['cell_code'] and code and sh.text(b['cell_code']) != code:
            log.append(('주택형', b['type'], sh.text(b['cell_code']), code))
            sh.set_text(b['cell_code'], code)
            mark(sh, b['cell_code'])
    # 3) 글자 바꾸기: 블록·단지명, 납부일, 계좌
    reps = []
    for r in rules:
        if r.get('apply') and r.get('old') and r.get('new'):
            reps.append((r['old'], r['new']))
    fmap = {f['key']: f for f in plan['fields']}
    for key in ('date', 'account', 'bank', 'holder'):
        v = fields.get(key)
        if key not in fmap or not v or v == fmap[key]['old'] or not (fmap[key]['old'] or fmap[key].get('cells')):
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
    # 4) 표 다시 짜기: 시트마다 아래 표부터
    by_sheet = OrderedDict()
    for ch in plan['changes']:
        if ch['sheet'] in targets and decisions.get(ch['id'], ch['default']):
            by_sheet.setdefault(ch['sheet'], []).append(ch)
    for name, chs in by_sheet.items():
        sh = book.sheet(name)
        for ch in sorted(chs, key=lambda c: -c['spec']['title_row']):
            if ch['spec'].get('drop'):
                drop_section(sh, ch['spec'])
                log.append(('표 삭제', name, ch['label'], ''))
                continue
            spec = ch['spec']
            keep = set(keep_names.get(ch['id']) or [])
            if keep:            # 고른 줄은 샘플 표기 그대로 (같은 품목의 제조사 줄도 함께)
                groups = {spec['rows'][i]['group'].get('3') for i in keep if i < len(spec['rows'])} - {None}
                rows2 = []
                for i, r in enumerate(spec['rows']):
                    if r.get('alt') and (i in keep or (r['group'].get('3') in groups and spec['kind'] == 'options')):
                        old = ch['rows'][i].get('old')
                        r = dict(r, vals={**r['vals'], **r['alt']},
                                 status='price' if old is not None and old != ch['rows'][i]['price'] else 'same')
                    rows2.append(r)
                spec = dict(spec, rows=rows2)
            rebuild_section(sh, spec, mark, log)
            if ch.get('hidden'):
                continue
            s = ch['summary']
            log.append(('표 다시 짜기', name, f"{ch['label']}: {s['before']}줄", f"{s['after']}줄 (추가 {s['add']}, 삭제 {s['del']}, "
                                                                             f"금액 {s['price']}, 이름 {s['text']})"))
            for r in ch['rows']:
                label = ' '.join(x for x in (r.get('no'), r.get('name'), r.get('brand')) if x).replace('\n', ' ')
                if r['status'] == 'new':
                    log.append(('추가', name, label, f"{r['price']:,}" if r.get('price') else ''))
                elif r['status'] == 'price':
                    log.append(('금액', name, f"{label}: {r['old']:,}" if isinstance(r.get('old'), (int, float)) else label,
                                f"{r['price']:,}"))
                elif r['status'] == 'text' and ch['rows'].index(r) not in set(keep_names.get(ch['id']) or []):
                    log.append(('이름', name, (r.get('old_name') or r.get('old_brand') or '').replace('\n', ' '), label))
            for r in ch['removed']:
                log.append(('삭제', name, (r.get('name') or '').replace('\n', ' '),
                            f"{r['price']:,}" if isinstance(r.get('price'), (int, float)) else ''))
    # 5) 공고에 없는 타입 시트 지우기
    for d in plan.get('delete_sheets', []):
        if delete_sheets.get(d['sheet'], d['apply']) and d['sheet'] in [s['name'] for s in book.sheet_list()]:
            if sum(1 for s in book.sheet_list() if s['state'] == 'visible') <= 1:
                continue
            book.delete_sheet(d['sheet'])
            targets = [t for t in targets if t != d['sheet']]
            log.append(('시트 삭제', d['sheet'], d['reason'], ''))
    # 6) '59B(수정)' → '59B'
    for rn in plan.get('renames', []):
        names = [s['name'] for s in book.sheet_list()]
        if not renames.get(rn['sheet'], rn.get('apply', True)) or rn['sheet'] not in names:
            continue
        if rn['new'] in names:
            log.append(('확인 필요', rn['sheet'], f"같은 이름의 숨긴 시트 '{rn['new']}' 가 있어 시트 이름을 바꾸지 않았습니다", ''))
            continue
        book.rename_sheet(rn['sheet'], rn['new'])
        targets = [rn['new'] if t == rn['sheet'] else t for t in targets]
        log.append(('시트 이름', rn['sheet'], rn['sheet'], rn['new']))
    # 7) 수식 표시값 다시 계산
    names = [s['name'] for s in book.sheet_list()]
    for name in targets:
        if name in names:
            book.sheet(name).recalc()
    return book.save(), log
