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


def grade_of(s):
    s = nospace(s)
    for g in ('일반형', '고급형', '프리미엄형', '기본형'):
        if g in s:
            return g
    return s


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
            for r in rows:
                if r >= t['rows']:
                    continue
                for c in range(t['cols']):
                    if name in nospace(g[r][c][0]):
                        return c
            return None
        pcol = col('공급금액')
        # 납부계좌·일정 표
        if '납부계좌' in alltxt and ('추가선택품목' in alltxt or '발코니확장' in alltxt) and not out['account']:
            for r in range(t['rows']):
                for c in range(t['cols']):
                    txt = nospace(g[r][c][0])
                    if re.fullmatch(r'\d[\d\-]{7,}\d', txt):
                        out['account'] = txt
                        bank = next((nospace(g[r][k][0]) for k in range(t['cols'])
                                     if re.search(r'(은행|농협|수협|금고|신협)$', nospace(g[r][k][0]))), '')
                        out['bank'] = bank
                        hol = next((g[r][k][0] for k in range(c + 1, t['cols']) if '신탁' in g[r][k][0] or '(주)' in g[r][k][0]), '')
                        out['holder'] = nospace(hol)
            mc = col('중도금', (0, 1, 2))
            if mc is not None:
                for r in range(t['rows']):
                    m = re.search(r'\d{4}\.\d{1,2}\.\d{1,2}', g[r][mc][0])
                    if m:
                        out['date'] = m.group(0)
                        break
            continue
        if pcol is None:
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
        elif '설치부위' in alltxt and '제조사' in alltxt:
            start = 2 if t['rows'] > 1 and not money(g[1][pcol][0]) else 1
            dcol = col('설치부위')
            for r in range(start, t['rows']):
                out['aircon'].append({'types': types_in(g[r][0][0]), 'grade': grade_of(g[r][1][0]), 'count': g[r][2][0],
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
                if mc is not None and t['rows'] > 1:
                    m = re.search(r'\d{4}\.\d{1,2}\.\d{1,2}', g[1][mc][0])
                    if m:
                        out['date'] = m.group(0)
    return out


# ------------------------------------------------------------------ 계약서 시트 읽기
SECTIONS = (('무상', 'free'), ('패키지', 'package'), ('에어컨', 'aircon'), ('추가선택', 'options'), ('발코니', 'balcony'))


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
        if kind:
            hr = r + 1
            heads = {nospace(val(hr, c)): c for c in range(2, 12)}
            pcol = next((c for h, c in heads.items() if '공급금액' in h), None)
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


def sheet_changes(sh, T, N, info):
    """시트 sh 를 타입 T 기준으로 공고 N 에 맞출 때 바뀌는 칸 목록."""
    out, warns = [], []
    val = info['val']
    sec = info['sections']

    def add(ref_, kind, label, old, new):
        out.append({'cell': ref_, 'kind': kind, 'label': label, 'old': old, 'new': new})

    # 시스템에어컨
    if 'aircon' in sec and N['aircon']:
        s = sec['aircon']
        cands = [a for a in N['aircon'] if T in a['types'] or '*' in a['types']]
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
            if len(groups) == 1 and '' in groups:
                pairs = list(zip(groups[''], cands))
                if len(groups['']) != len(cands):
                    warns.append(f'{T}: 시스템에어컨 줄 수({len(groups[""])})와 공고({len(cands)})가 다릅니다.')
            else:
                pairs = []
                for g, rows in groups.items():
                    lst = ng.get(g, [])
                    if len(rows) != len(lst):
                        warns.append(f'{T}: 시스템에어컨 {g} 줄 수({len(rows)})와 공고({len(lst)})가 다릅니다. 확인하세요.')
                    pairs += list(zip(rows, lst))
            for r, a in pairs:
                pref = f'{num2col(s["pcol"])}{r}'
                old = sh.value(pref)
                if a['price'] and old != a['price']:
                    add(pref, 'price', f'시스템에어컨 {a["grade"]} {nospace(a["count"])}', old, a['price'])
                dref = f'C{r}'
                if a['desc'] and nospace(val(r, 3)) != nospace(a['desc']):
                    add(dref, 'desc', f'시스템에어컨 {a["grade"]} {nospace(a["count"])} 설치부위', val(r, 3), a['desc'])
    # 유상옵션
    if 'options' in sec and N['options']:
        s = sec['options']
        cands = [o for o in N['options'] if T in o['types'] or '*' in o['types']]
        cur_item = ''
        for r in s['rows']:
            item = val(r, 3)
            if item.strip():
                cur_item = item
            brand = nospace(val(r, 4))
            if brand == nospace(cur_item):
                brand = ''
            m = match_option(cur_item, brand, cands)
            label = clean_desc(cur_item) + (f' ({brand})' if brand else '')
            if m is None:
                warns.append(f'{T}: 유상옵션 "{label}" 을(를) 공고에서 찾지 못했습니다 (이 타입에 없거나 이름이 다름).')
                continue
            pref = f'{num2col(s["pcol"])}{r}'
            old = sh.value(pref)
            if m['price'] and old != m['price']:
                add(pref, 'price', f'유상옵션 {label}', old, m['price'])
    # 패키지
    if 'package' in sec and N['packages']:
        s = sec['package']
        for r in s['rows']:
            name = nospace(val(r, 2))
            m = next((p for p in N['packages'] if (p['name'] == name or name in p['name'] or p['name'] in name)
                      and (T in p['types'] or '*' in p['types'])), None)
            if m is None:
                warns.append(f'{T}: 패키지 "{name}" 을(를) 공고에서 찾지 못했습니다.')
                continue
            pref = f'{num2col(s["pcol"])}{r}'
            old = sh.value(pref)
            if m['price'] and old != m['price']:
                add(pref, 'price', f'패키지 {name}', old, m['price'])
    return out, warns


def balcony_cells(info):
    s = info['sections'].get('balcony')
    if not s or not s['rows']:
        return None, None
    r = s['rows'][0]
    return f'B{r}', f'{num2col(s["pcol"])}{r}'


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
    type_sheets = [s['name'] for s in sheets if s['name'] in ntypes or re.fullmatch(r'\d{2,3}[A-Z]{0,3}', s['name'])]
    if not type_sheets:
        raise OptionError('주택형 이름(예: 84A)으로 된 시트를 찾지 못했습니다.')
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
    new_sheets = [{'type': T, 'template': pick_template(T), 'apply': True} for T in ntypes if T not in type_sheets]
    for name in type_sheets:
        if ntypes and name not in ntypes:
            warnings.append(f"'{name}' 시트의 타입이 공고에 없습니다. 시트를 지울지 확인하세요.")

    # 시트별 변경
    cid = [0]

    def push(sheet, ch, new_sheet=False):
        cid[0] += 1
        ch.update(id=f'o{cid[0]}', sheet=sheet, default=True, new_sheet=new_sheet)
        changes.append(ch)
    targets = [(n, n, False) for n in type_sheets] + [(ns['type'], ns['template'], True) for ns in new_sheets]
    for T, src, is_new in targets:
        sh = book.sheet(src)
        chs, warns = sheet_changes(sh, T, N, infos[src])
        warnings += warns
        for ch in chs:
            push(T, ch, is_new)

    # 발코니 확장 (공고 금액이 비어 있을 수 있어 입력칸으로 보여 준다)
    balcony = []
    nb = {b['type']: b for b in N['balcony']}
    blank = []
    for T, src, is_new in targets:
        cref, pref = balcony_cells(infos[src])
        sh = book.sheet(src)
        old_code = sh.text(cref) if cref else ''
        old_price = sh.value(pref) if pref else None
        b = nb.get(T)
        nprice = b['price'] if b else None
        if b and not nprice:
            blank.append(T)
        balcony.append({'type': T, 'new_sheet': is_new, 'template': src if is_new else '',
                        'code_old': old_code if not is_new else '', 'code_new': (b['code'] if b and b['code'] else old_code),
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
        if ot and nt and ot != nt:
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
            old_holder = nospace(h.group(1)) if h else ''
            break
    fields = [
        {'key': 'date', 'label': '중도금 납부일 (표 머리글)', 'old': old_date, 'new': N['date'], 'apply': bool(old_date and N['date'] and old_date != N['date'])},
        {'key': 'bank', 'label': '납부계좌 · 은행', 'old': old_bank, 'new': N['bank'], 'apply': bool(old_bank and N['bank'] and old_bank != N['bank'])},
        {'key': 'account', 'label': '납부계좌 · 계좌번호', 'old': old_acct, 'new': N['account'], 'apply': bool(old_acct and N['account'] and old_acct != N['account'])},
    ]
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
        if b['type'] not in targets:
            continue
        sh = book.sheet(b['type'])
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
    for key in ('date', 'account', 'bank'):
        v = fields.get(key)
        if v and fmap[key]['old'] and v != fmap[key]['old']:
            reps.append((fmap[key]['old'], v))
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
