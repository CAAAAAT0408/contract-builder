"""샘플 계약서 + 새 현장 모집공고 → 변경안(plan) 만들기 / 변경안 적용해 계약서 만들기.

방식
  1) 1페이지: 재산의 표시(블록·단지명), 입주예정일, 약정일자, 납부계좌를 공고에서 뽑아 바꾼다.
  2) 제○조(기타사항): 계약서 항목과 공고 '유의사항' 항목을 짝지어
       - 내용이 바뀐 곳만 고친다 (계약서에서 다듬어 둔 끝맺음 문구는 그대로 둔다)
       - 공고에만 있는 항목은 끼워 넣는다
       - 계약서에만 있는 항목은 '확인 필요'로 남긴다 (계약서 전용 문구일 수도, 공고에서 빠진 항목일 수도 있다)
       - 동별 유의사항은 동 번호가 다르면 통째로 바꾼다
  3) 계약서 전체에서 이전 현장 블록·단지명을 새 값으로 바꾼다.
  바뀐 글자는 노란 형광펜으로 표시한다.
"""
import difflib
import re
from collections import Counter, defaultdict

from contract import (Contract, ContractError, edit_range, make_para, nospace, ptext, replace_all, w)
from hwp_reader import HwpError, read_hwp
from notice import parse_notice

TOK = re.compile(r'\s+|\d+(?:\.\d+)?|[A-Za-z]+|[가-힣]+|.')
WORD = re.compile(r'[가-힣A-Za-z0-9]+')
# 공고 문구를 계약서에 옮기며 담당자가 흔히 다듬는 끝맺음 낱말. 이 낱말끼리만 다르면 계약서 쪽을 그대로 둔다.
WORDING = set('''청약 및 계약 체결 전 전에 후 확인 인지 양지 진행 것으로 신청 청약신청 을 를 이를 반드시 사전에 충분히
이점을 방문 함 바람 바랍니다 하여야 하시기 하고 하며 하시어 있 숙지'''.split())
ENDINGS = sorted('''하여야 하시기 하시길 하였으며 하였음 하고 하신 하신후 한 한후 하며 하시어 하여 하시고 하는 바람 하시기바람
하여야함 함 으니 음 니 을 를 에 바람'''.split(), key=len, reverse=True)
BRACKETS = [('(', ')'), ('「', '」'), ('『', '』'), ('[', ']'), ('“', '”')]


def stem(word):
    for pre in ('청약', '계약'):
        if word.startswith(pre) and word != pre:
            word = word[len(pre):]
    changed = True
    while changed and word:
        changed = False
        for e in ENDINGS:
            if word.endswith(e) and len(word) > len(e) - 1:
                word = word[:-len(e)]
                changed = True
                break
    return word


def content_words(s):
    out = Counter()
    for wd in WORD.findall(s):
        st = stem(wd)
        if st and st not in WORDING and wd not in WORDING:
            out[st] += 1
    return out


DOTS = str.maketrans({'․': '·', 'ㆍ': '·', '・': '·', '‧': '·', '⏎': ''})


def canon(s):
    return s.translate(DOTS)


def balanced(s):
    for a, b in BRACKETS:
        depth = 0
        for ch in s:
            if ch == a: depth += 1
            elif ch == b:
                depth -= 1
                if depth < 0: return False
        if depth: return False
    return True


class PlanError(Exception):
    pass


# ------------------------------------------------------------------ 문자열 도구
def nkey(s):
    s = re.sub(r'청약\s*및\s*|청약\(계약\)|청약신청\s*및\s*|청약\s*전에?\s*|청약', '', s)
    return re.sub(r'[\s■\-.,·ㆍ・‧"“”‘’\'()\[\]「」『』<>:;~/※]', '', s)


def squash(s):
    return re.sub(r'[\s.,·]', '', s)


def cheongyak(s):
    """계약서는 청약 뒤에 쓰므로 공고의 '청약' 표현을 뺀다."""
    s = re.sub(r'\(\s*청약\s*전[^)]*\)', '', s)
    s = re.sub(r'청약신청\s*및\s*계약체결', '계약체결', s)
    s = re.sub(r'청약\s*및\s*계약', '계약', s)
    s = s.replace('청약(계약)', '계약')
    s = re.sub(r'청약\s*전에\s*', '', s)
    s = re.sub(r'확인\s*후\s*청약하여야', '확인 후 계약하여야', s)
    s = re.sub(r'(?<![가-힣])청약하여야', '계약하여야', s)
    return s


def tidy(s):
    s = re.sub(r'(?<=\S) {2,}(?=\S)', ' ', s)
    return s.rstrip()


def apply_strike_rules(s, rules):
    for r in rules:
        s = s.replace(r['left'] + r['text'] + ' ', r['left'])
        s = s.replace(r['left'] + r['text'], r['left'])
    return s


def apply_edits(s, edits):
    for i, j, new in sorted(edits, key=lambda e: -e[0]):
        s = s[:i] + new + s[j:]
    return s


def segs_before(s, edits):
    out, pos = [], 0
    for i, j, _ in sorted(edits):
        if i > pos: out.append([s[pos:i], 0])
        if j > i: out.append([s[i:j], 1])
        pos = j
    if pos < len(s): out.append([s[pos:], 0])
    return out


def segs_after(s, edits):
    out, pos = [], 0
    for i, j, new in sorted(edits):
        if i > pos: out.append([s[pos:i], 0])
        if new: out.append([new, 1])
        pos = j
    if pos < len(s): out.append([s[pos:], 0])
    return out


def edits_from(old, new):
    a, b = TOK.findall(old), TOK.findall(new)
    offs = [0]
    for t in a: offs.append(offs[-1] + len(t))
    out = []
    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
        if op != 'equal':
            out.append((offs[i1], offs[i2], ''.join(b[j1:j2])))
    return out


def merge(c, n):
    """계약서 문구 c 에 공고 문구 n 의 '내용 차이'만 반영.
    (고칠곳, 계약서에만 있는 괄호문구, 괄호짝 경고) 를 돌려준다."""
    a, b = TOK.findall(c), TOK.findall(n)
    offs = [0]
    for t in a: offs.append(offs[-1] + len(t))
    edits, extras, held = [], [], []
    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
        if op == 'equal':
            continue
        A, B = ''.join(a[i1:i2]), ''.join(b[j1:j2])
        e = (offs[i1], offs[i2], B)
        if nospace(canon(A)) == nospace(canon(B)):
            continue                                   # 띄어쓰기·가운뎃점 모양·줄바꿈만 다름
        wa, wb = WORD.findall(A), WORD.findall(B)
        pa, pb = re.sub(r'[\w\s]', '', canon(A)), re.sub(r'[\w\s]', '', canon(B))
        if not wa and not wb:
            if pa.replace('.', '') == pb.replace('.', ''):
                continue                               # 마침표만 다름
            if any(ch in pa + pb for pair in BRACKETS for ch in pair):
                held.append(e)                         # 괄호만 다름: 아래에서 짝을 보고 정한다
                continue
            edits.append(e)
            continue
        if not wb and A.strip().startswith('(') and A.strip().endswith(')'):
            extras.append((offs[i1], offs[i2], ''))
            continue
        if content_words(A) == content_words(B) and re.sub('[.,]', '', pa) == re.sub('[.,]', '', pb):
            continue                                   # 끝맺음 표현만 다름 (계약서 쪽 유지)
        edits.append(e)
    warn = ''
    if balanced(c) and not balanced(apply_edits(c, edits)):
        for h in held:
            if balanced(apply_edits(c, edits + [h])):
                edits.append(h)
                break
        else:
            warn = '괄호 짝이 맞지 않습니다. 공고 원문을 확인하세요.'
    # 공백만 사이에 둔 고칠곳은 하나로 합쳐 형광펜이 잘게 쪼개지지 않게 한다
    edits.sort()
    joined = []
    for e in edits:
        if joined and not c[joined[-1][1]:e[0]].strip() and e[0] - joined[-1][1] <= 1:
            p = joined[-1]
            joined[-1] = (p[0], e[1], p[2] + c[p[1]:e[0]] + e[2])
        else:
            joined.append(e)
    return joined, extras, warn


class Sim:
    def __init__(self):
        self.cache = {}

    def __call__(self, a, b):
        k = (a, b)
        if k not in self.cache:
            if not a or not b:
                v = 0.0
            else:
                la, lb = len(a), len(b)
                if min(la, lb) / max(la, lb) < 0.4:
                    v = 0.0
                else:
                    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
                    v = sm.ratio() if sm.quick_ratio() > 0.45 else 0.0
            self.cache[k] = v
        return self.cache[k]


def align(cs, ns, sim, thr=0.5):
    """순서를 지키며 짝짓기 (최대 유사도 합). cs, ns 는 key 문자열 목록. [(ci, ni)]"""
    m, n = len(cs), len(ns)
    if not m or not n:
        return []
    S = [[0.0] * (n + 1) for _ in range(m + 1)]
    for i in range(1, m + 1):
        for j in range(1, n + 1):
            best = max(S[i - 1][j], S[i][j - 1])
            s = sim(cs[i - 1], ns[j - 1])
            if s >= thr:
                best = max(best, S[i - 1][j - 1] + s)
            S[i][j] = best
    out, i, j = [], m, n
    while i and j:
        s = sim(cs[i - 1], ns[j - 1])
        if s >= thr and abs(S[i][j] - (S[i - 1][j - 1] + s)) < 1e-9:
            out.append((i - 1, j - 1)); i -= 1; j -= 1
        elif S[i][j] == S[i - 1][j]:
            i -= 1
        else:
            j -= 1
    return out[::-1]


def name_tail(name):
    """'○○ 우미 린 센트럴' → '센트럴' (브랜드 '린' 뒤의 고유 이름)"""
    if not name:
        return ''
    m = re.search(r'린\s*(.+)$', name)
    return (m.group(1) if m else name).strip()


# ------------------------------------------------------------------ 준비 (분석·생성 공통)
def prepare(cbytes, nbytes, opts):
    try:
        paras = read_hwp(nbytes)
    except HwpError as e:
        raise PlanError(str(e))
    notice = parse_notice(paras)
    if not notice['sections']:
        raise PlanError("공고문에서 '유의사항' 표(구분/내용)를 찾지 못했습니다.")
    try:
        C = Contract(cbytes)
        C.accept_revisions()
        removed, srules = C.apply_strikes() if opts.get('strike', True) else ([], [])
        units, blocks = C.region_units()
    except ContractError as e:
        raise PlanError(str(e))
    return notice, C, units, blocks, removed, srules


# ------------------------------------------------------------------ 분석
def analyze(cbytes, nbytes, opts):
    notice, C, units, blocks, removed, srules = prepare(cbytes, nbytes, opts)
    nf = notice['fields']
    nsecs = notice['sections']
    sim = Sim()
    changes, warnings = [], []
    cid = [0]

    def new_change(**kw):
        cid[0] += 1
        kw['id'] = f'c{cid[0]}'
        changes.append(kw)
        return kw

    # ---------- 공고 항목 펼치기
    N = []
    for si, s in enumerate(nsecs):
        for u in s.units:
            t = tidy(apply_strike_rules(cheongyak(u.text), srules))
            N.append({'si': si, 'sec': s, 'unit': u, 'text': t, 'key': nkey(t), 'dong': s.dong})
    for x in N:
        if '청약' in x['text']:
            warnings.append(f"공고 문구에 '청약'이 남아 있습니다: {x['text'][:60]}…")

    # ---------- 동별: 동 번호가 겹치지 않으면 통째로 교체
    c_dongs = []
    for u in units:
        if u.dong and u.dong not in c_dongs:
            c_dongs.append(u.dong)
    n_dongs = [nospace(s.label) for s in nsecs if s.dong]
    overlap = len(set(c_dongs) & set(n_dongs)) / max(1, len(set(c_dongs)))
    dong_replace = bool(c_dongs) and bool(n_dongs) and overlap < 0.5

    Ckeys = [nkey(u.text) for u in units]
    match = {}          # c index -> N index
    used = set()

    def csec(u):
        return ('dong', u.dong) if u.dong else (u.major, u.section)

    if dong_replace:
        c_pool = [i for i, u in enumerate(units) if not u.dong]
        n_pool = [j for j, x in enumerate(N) if not x['dong']]
    else:
        c_pool = list(range(len(units)))
        n_pool = list(range(len(N)))

    # 1) 계약서 구분(<...>) ↔ 공고 구분 칸 짝짓기: 항목 내용 다수결
    votes = defaultdict(Counter)
    for i in c_pool:
        u = units[i]
        if u.dong:
            continue
        best, bj = 0, None
        for j in n_pool:
            if N[j]['dong']:
                continue
            s = sim(Ckeys[i], N[j]['key'])
            if s > best:
                best, bj = s, j
        if bj is not None and best >= 0.7:
            votes[csec(u)][N[bj]['si']] += 1
    secmap = {k: v.most_common(1)[0][0] for k, v in votes.items()}
    if not dong_replace:
        for i in c_pool:
            u = units[i]
            if u.dong:
                si = next((k for k, s in enumerate(nsecs) if s.dong and nospace(s.label) == u.dong), None)
                if si is not None:
                    secmap[csec(u)] = si

    # 2) 짝지은 구분 안에서 순서대로 맞추기
    order = []
    for i in c_pool:
        k = csec(units[i])
        if k not in order:
            order.append(k)
    for k in order:
        if k not in secmap:
            continue
        ci = [i for i in c_pool if csec(units[i]) == k and i not in match]
        nj = [j for j in n_pool if N[j]['si'] == secmap[k] and j not in used]
        for a, b in align([Ckeys[i] for i in ci], [N[j]['key'] for j in nj], sim):
            match[ci[a]] = nj[b]
            used.add(nj[b])

    # 2-1) 같은 자리(앞뒤로 짝지어진 항목 사이)에 남은 것끼리는 문구가 많이 바뀌었어도 같은 항목으로 본다
    for k in order:
        if k not in secmap:
            continue
        ci = [i for i in c_pool if csec(units[i]) == k]
        nj = [j for j in n_pool if N[j]['si'] == secmap[k]]

        def gap_c(i):
            prev = [q for q in ci if q < i and q in match]
            return match[prev[-1]] if prev else None

        def gap_n(j):
            prev = [q for q in nj if q < j and q in used]
            return prev[-1] if prev else None
        cand = []
        for i in ci:
            if i in match:
                continue
            for j in nj:
                if j in used or gap_c(i) != gap_n(j):
                    continue
                s_ = sim(Ckeys[i], N[j]['key']) if True else 0
                if s_ < 0.35:
                    s_ = difflib.SequenceMatcher(None, Ckeys[i], N[j]['key'], autojunk=False).ratio()
                if s_ >= 0.35:
                    cand.append((s_, i, j))
        for s_, i, j in sorted(cand, reverse=True):
            if i not in match and j not in used:
                match[i] = j
                used.add(j)

    # 3) 남은 것끼리 전체에서 다시 (구분이 옮겨진 항목)
    cand = []
    for i in c_pool:
        if i in match:
            continue
        for j in n_pool:
            if j in used:
                continue
            if N[j]['dong'] != bool(units[i].dong) and not dong_replace:
                continue
            s = sim(Ckeys[i], N[j]['key'])
            if s >= 0.72:
                cand.append((s, i, j))
    for s, i, j in sorted(cand, reverse=True):
        if i not in match and j not in used:
            match[i] = j
            used.add(j)

    old = C.old_identity()
    old_block = old.get('block', '')
    old_tail = name_tail(old.get('name', ''))
    dong_digits = [re.sub(r'\D', '', d) for d in c_dongs]

    def has_old_info(t):
        if old_block and old_block in t:
            return True
        if old_tail and len(old_tail) >= 3 and old_tail in t and old_tail not in name_tail(nf.get('complex_name', '')):
            return True
        return any(d and d in t for d in dong_digits) if dong_replace else False

    # ---------- 짝지어진 항목: 바뀐 내용 반영
    for i, j in sorted(match.items()):
        u, x = units[i], N[j]
        c = u.text
        edits, extras, bwarn = merge(c, x['text'])
        label = u.section or u.major or x['sec'].label
        if u.dong:
            label = f'{u.section} {u.dong}'
        if edits:
            final = apply_edits(c, edits)
            new_change(kind='update', unit=i, edits=edits, group=label, default=True,
                       title='공고 내용 변경 반영', before=segs_before(c, edits), after=segs_after(c, edits),
                       note=bwarn)
        for e in extras:
            new_change(kind='extra', unit=i, edits=[e], group=label, default=False,
                       title='계약서에만 있는 괄호 문구 (유지 중)', before=segs_before(c, [e]),
                       after=segs_after(c, [e]),
                       note='체크하면 이 괄호 문구를 지웁니다. 검토 메모라면 지우세요.')

    # ---------- 계약서에만 있는 항목
    for i in c_pool:
        if i in match:
            continue
        u = units[i]
        t = u.text
        old_info = has_old_info(t)
        new_change(kind='contract_only', unit=i, group=u.section or u.major, default=old_info,
                   title='공고에 없는 계약서 문구' + (' — 이전 현장 정보 포함' if old_info else ''),
                   before=[[t, 0]], after=[], note='체크하면 삭제합니다. 계약서 전용 문구라면 그대로 두세요.'
                   if not old_info else '이전 현장의 동·블록·단지명이 들어 있어 삭제를 권합니다.')

    # ---------- 공고에만 있는 항목 → 끼워 넣기
    rem_keys = [nkey(t) for t in removed]
    c_index_of_n = {j: i for i, j in match.items()}
    for si, s in enumerate(nsecs):
        if s.dong and dong_replace:
            continue
        js = [j for j, x in enumerate(N) if x['si'] == si]
        k = 0
        while k < len(js):
            j = js[k]
            if j in used:
                k += 1
                continue
            run = [j]
            while k + 1 < len(js) and js[k + 1] not in used:
                k += 1
                run.append(js[k])
            k += 1
            prev = next((c_index_of_n[q] for q in reversed(js[:js.index(run[0])]) if q in c_index_of_n), None)
            nxt = next((c_index_of_n[q] for q in js[js.index(run[-1]) + 1:] if q in c_index_of_n), None)
            if prev is not None:
                anchor = ('after', prev)
            elif nxt is not None:
                anchor = ('before', nxt)
            else:
                anchor = ('end', None)
                warnings.append(f"공고의 '{s.label}' 구분과 맞는 계약서 구분을 찾지 못해 20조 끝에 붙였습니다. 위치를 확인하세요.")
            for q in run:
                x = N[q]
                sup = any(sim(x['key'], rk) >= 0.85 for rk in rem_keys)
                where = units[anchor[1]].section if anchor[1] is not None else s.label
                new_change(kind='add', n=q, anchor=anchor, text=x['text'], unit_kind=x['unit'].kind,
                           group=(units[anchor[1]].section or units[anchor[1]].major) if anchor[1] is not None else s.label,
                           default=not sup,
                           title='공고에 새로 생긴 항목 추가' if not sup else '샘플 계약서에서 삭제 표시했던 항목',
                           before=[], after=[[x['text'], 1]],
                           note='' if not sup else '샘플 계약서에서 취소선으로 지운 문구와 같아 기본으로 넣지 않습니다.')

    # ---------- 동별 통째 교체
    if dong_replace:
        blocks_out = []
        for s in nsecs:
            if s.dong:
                blocks_out.append({'dong': nospace(s.label),
                                   'items': [x['text'] for x in N if x['sec'] is s]})
        new_change(kind='dong_replace', group='<동별현황>', default=True,
                   title=f"동별 유의사항 교체: {', '.join(c_dongs[:1])}~{c_dongs[-1]} → {blocks_out[0]['dong']}~{blocks_out[-1]['dong']}",
                   before=[[f"{len(c_dongs)}개 동 {sum(1 for u in units if u.dong)}개 항목", 0]],
                   after=[[f"{b['dong']} ({len(b['items'])}개 항목)", 1] for b in blocks_out],
                   blocks=blocks_out, note='')

    # ---------- 취소선 반영 내역 (정보)
    for t in removed:
        new_change(kind='info', group='취소선 반영', default=True, title='샘플 계약서에서 취소선으로 지운 문단을 뺐습니다',
                   before=[[t, 1]], after=[], note='')
    for r in srules:
        new_change(kind='info', group='취소선 반영', default=True,
                   title=f"취소선 글자 '{r['text']}' 를 지웠고, 새 문구에도 같은 방식으로 적용합니다", before=[[r['para'], 0]],
                   after=[], note='')

    # ---------- 1페이지 값
    fields = []
    mc = C.movein_cell()
    if mc is not None:
        m = re.search(r'\d{4}\s*년\s*\d{1,2}\s*월', ptext(mc))
        fields.append({'key': 'move_in', 'label': '입주예정일', 'old': m.group(0) if m else '',
                       'new': nf.get('move_in', ''), 'apply': bool(m and nf.get('move_in'))})
    slots = C.payment_slots()
    dates = nf.get('dates', [])
    if slots and len(dates) != len(slots):
        warnings.append(f'약정일자 칸은 {len(slots)}개인데 공고에서 찾은 날짜는 {len(dates)}개입니다. 날짜를 확인하세요.')
    for k, (label, _cell, cur) in enumerate(slots):
        fields.append({'key': f'date{k}', 'label': f'약정일자 · {label}', 'old': cur,
                       'new': dates[k] if k < len(dates) else '', 'apply': k < len(dates)})
    acc = C.account_para()
    if acc:
        _p, bank, num, holder = acc
        fields.append({'key': 'bank', 'label': '납부계좌 · 은행', 'old': bank, 'new': nf.get('bank', ''),
                       'apply': bool(nf.get('bank'))})
        fields.append({'key': 'account', 'label': '납부계좌 · 계좌번호', 'old': num, 'new': nf.get('account', ''),
                       'apply': bool(nf.get('account'))})
        fields.append({'key': 'holder', 'label': '납부계좌 · 예금주', 'old': holder,
                       'new': nf.get('holder', '').replace('(주)', '㈜') if '㈜' in holder else nf.get('holder', ''),
                       'apply': bool(nf.get('holder')) and nospace(holder) != nospace(nf.get('holder', '')).replace('(주)', '㈜')})
    else:
        warnings.append('계약서에서 납부계좌 문구(○○은행 계좌번호)를 찾지 못했습니다.')

    # ---------- 이전 현장 이름 바꾸기 규칙
    rules = []
    if old_block and nf.get('block') and old_block != nf['block']:
        rules.append({'old': old_block, 'new': nf['block'], 'apply': True})
    new_tail = name_tail(nf.get('complex_name', ''))
    if old_tail and new_tail and old_tail != new_tail:
        rules.append({'old': old_tail, 'new': new_tail, 'apply': True})
    alltext = '\n'.join(ptext(p) for p in C.body.iter(w('p')))
    for r in rules:
        r['count'] = alltext.count(r['old'])

    if not old.get('block'):
        warnings.append("계약서 1페이지 '재산의 표시' 에서 블록 번호를 찾지 못했습니다.")
    for k, v in (('complex_name', '단지명'), ('block', '블록'), ('move_in', '입주시기'), ('dates', '납부일정'),
                 ('account', '납부계좌')):
        if not nf.get(k):
            warnings.append(f'공고에서 {v}을(를) 찾지 못했습니다. 직접 입력하세요.')

    # ---------- 표 비교 (학교 개교시기 표 등)
    ctables = [b[1] for b in blocks if b[0] == 'table']
    for s in nsecs:
        for t in s.tables:
            ncells = [nospace(c) for c in t['cells']]
            best, bt = 0, None
            for ct in ctables:
                cc = [nospace(ptext(tc)) for tc in ct.iter(w('tc'))]
                inter = len(set(cc) & set(ncells))
                if inter > best:
                    best, bt = inter, cc
            if bt is None:
                continue
            add = [c for c in ncells if c and c not in bt]
            gone = [c for c in bt if c and c not in ncells]
            if add or gone:
                warnings.append(f"'{s.label}' 의 표 내용이 계약서 표와 다릅니다 (공고에만: {', '.join(add[:8]) or '없음'} / "
                                f"계약서에만: {', '.join(gone[:8]) or '없음'}). 표는 자동으로 고치지 않으니 직접 확인하세요.")

    # 검토 화면에서 계약서 순서대로 보이도록 정렬
    first_dong = min((i for i, u in enumerate(units) if u.dong), default=10 ** 6)
    for k, ch in enumerate(changes):
        if 'unit' in ch:
            ch['pos'] = ch['unit'] + (0.1 if ch['kind'] == 'extra' else 0)
        elif ch['kind'] == 'add':
            how, ui = ch['anchor']
            ch['pos'] = (ui + 0.5 if how == 'after' else ui - 0.5) if ui is not None else 10 ** 6
        elif ch['kind'] == 'dong_replace':
            ch['pos'] = first_dong
        else:
            ch['pos'] = 10 ** 7
        ch['pos'] += k * 1e-6
    changes.sort(key=lambda c: c['pos'])
    summary = Counter(c['kind'] for c in changes)
    return {
        'opts': opts,
        'notice': {k: v for k, v in nf.items()},
        'old': old,
        'fields': fields,
        'rules': rules,
        'changes': changes,
        'warnings': warnings,
        'summary': dict(summary),
        'n_sections': len(nsecs),
        'n_units': len(units),
        'matched': len(match),
    }


# ------------------------------------------------------------------ 생성
def build(cbytes, nbytes, plan, decisions, fields, rules):
    """decisions: {change_id: bool}, fields: {key: new or None}, rules: [{'old','new','apply'}]"""
    notice, C, units, blocks, removed, srules = prepare(cbytes, nbytes, plan['opts'])
    rpr = C.plain_rpr(units)
    on = lambda ch: decisions.get(ch['id'], ch['default'])
    log = []

    def tmpl_for(kind, near=None):
        if near is not None and units[near].kind == kind:
            return units[near].paras[0]
        for u in units:
            if u.kind == kind and not u.dong:
                return u.paras[0]
        for u in units:
            if u.kind == 'item':
                return u.paras[0]
        return None

    # 1) 끼워넣기 (지우기 전에 위치를 잡는다)
    last_after = {}
    for ch in plan['changes']:
        if ch['kind'] != 'add' or not on(ch):
            continue
        how, ui = ch['anchor']
        tp = tmpl_for(ch['unit_kind'], ui)
        np_ = make_para(tp, [(ch['text'], True)], rpr)
        if how == 'after':
            ref = last_after.get(ui, units[ui].paras[-1])
            ref.addnext(np_)
            last_after[ui] = np_
        elif how == 'before':
            units[ui].paras[0].addprevious(np_)
        else:
            end_ref = blocks[-1][1]
            end_ref.addnext(np_)
        log.append(('추가', ch['group'], '', ch['text']))

    # 2) 문구 고치기 (같은 항목의 고칠곳을 모아 뒤에서부터)
    per_unit = defaultdict(list)
    for ch in plan['changes']:
        if ch['kind'] in ('update', 'extra') and on(ch):
            per_unit[ch['unit']].extend(ch['edits'])
    for ui, edits in per_unit.items():
        u = units[ui]
        before = u.text
        p = u.merge_paras() if len(u.paras) > 1 else u.paras[0]
        for i, j, new in sorted(edits, key=lambda e: -e[0]):
            edit_range(p, i, j, new, True)
        log.append(('변경', u.section, before, ptext(p)))

    # 3) 지우기
    for ch in plan['changes']:
        if ch['kind'] == 'contract_only' and on(ch):
            u = units[ch['unit']]
            log.append(('삭제', u.section, u.text, ''))
            for p in u.paras:
                if p.getparent() is not None:
                    p.getparent().remove(p)

    # 4) 동별 통째 교체
    for ch in plan['changes']:
        if ch['kind'] != 'dong_replace' or not on(ch):
            continue
        idx = [k for k, b in enumerate(blocks) if b[0] == 'dong']
        ds = idx[0]
        de = ds
        for k in range(ds, len(blocks)):
            if blocks[k][0] in ('header', 'major', 'fixed', 'table', 'other'):
                break
            de = k + 1
        dong_tp = blocks[ds][1]
        item_tp = next((b[1] for b in blocks[ds:de] if b[0] == 'item'), tmpl_for('item'))
        blank_tp = next((b[1] for b in blocks[ds:de] if b[0] == 'blank'), None)
        dong_rpr = next((r.find(w('rPr')) for r in dong_tp.iter(w('r'))), rpr)
        anchor = blocks[ds][1].getprevious()
        parent = blocks[ds][1].getparent()
        pos = parent.index(blocks[ds][1])
        new_nodes = []
        for bi, b in enumerate(ch['blocks']):
            if bi:
                new_nodes.append(make_para(blank_tp if blank_tp is not None else item_tp, [], rpr))
            new_nodes.append(make_para(dong_tp, [(b['dong'], True)], dong_rpr))
            for t in b['items']:
                new_nodes.append(make_para(item_tp, [(t, True)], rpr))
        new_nodes.append(make_para(blank_tp if blank_tp is not None else item_tp, [], rpr))
        for k in range(ds, de):
            el = blocks[k][1]
            if el.getparent() is not None:
                el.getparent().remove(el)
        for n_ in reversed(new_nodes):
            parent.insert(pos, n_)
        log.append(('동별 교체', '<동별현황>', '', ', '.join(b['dong'] for b in ch['blocks'])))

    # 5) 이전 현장 이름 바꾸기 (문서 전체)
    for r in rules:
        if not r.get('apply') or not r.get('old') or not r.get('new'):
            continue
        n = 0
        for p in list(C.body.iter(w('p'))):
            if r['old'] in ptext(p):
                n += replace_all(p, r['old'], r['new'], True)
        log.append(('이름 바꾸기', '전체', r['old'], f"{r['new']} ({n}곳)"))

    # 6) 1페이지 값
    def fval(key):
        return (fields.get(key) or '').strip()

    mc = C.movein_cell()
    if mc is not None and fields.get('move_in') is not None and fval('move_in'):
        for p in mc.iter(w('p')):
            m = re.search(r'\d{4}\s*년\s*\d{1,2}\s*월', ptext(p))
            if m:
                edit_range(p, m.start(), m.end(), fval('move_in'), True)
                log.append(('1페이지', '입주예정일', m.group(0), fval('move_in')))
                break
    for k, (label, cell, cur) in enumerate(C.payment_slots()):
        v = fields.get(f'date{k}')
        if v is None or not v.strip():
            continue
        p = cell.find(w('p'))
        if cur:
            edit_range(p, 0, len(ptext(p)), v.strip(), True)
        else:
            ppr = p.find(w('pPr'))
            prpr = ppr.find(w('rPr')) if ppr is not None else None
            from contract import make_run
            p.append(make_run(prpr, v.strip(), True))
        log.append(('1페이지', f'약정일자 {label}', cur, v.strip()))
    acc = C.account_para()
    if acc:
        p, bank, num, holder = acc
        for key, oldv in (('account', num), ('bank', bank), ('holder', holder)):
            v = fields.get(key)
            if v is None or not v.strip() or not oldv or v.strip() == oldv:
                continue
            t = ptext(p)
            k = t.find(oldv)
            if k >= 0:
                edit_range(p, k, k + len(oldv), v.strip(), True)
                log.append(('1페이지', f'납부계좌 {key}', oldv, v.strip()))

    return C.save(), log
