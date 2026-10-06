"""입주자모집공고에서 계약서에 필요한 값과 '유의사항' 항목을 뽑는다."""
import re

DATE_RE = re.compile(r'^(\d{4})\s*[.\-]\s*(\d{1,2})\s*[.\-]\s*(\d{1,2})\.?$')
ACCT_RE = re.compile(r'^\d[\d\-]{6,}\d$')
DONG_RE = re.compile(r'^\d{3,4}\s*동$')
TERMINAL = ('.', '함', '음', '람', ')', '임', '됨', '요')


class Unit:
    """공고/계약서 공통 '항목' 단위. kind: item(■) / sub(들여쓴 '-' 나 괄호줄)"""
    def __init__(self, kind, text, section):
        self.kind, self.text, self.section = kind, text, section

    def __repr__(self):
        return f'<{self.kind} {self.section!r} {self.text[:30]!r}>'


class Section:
    def __init__(self, label, dong=False):
        self.label, self.dong = label, dong
        self.units, self.tables = [], []


def _clean(s):
    return s.replace('　', ' ')


def find_caution_range(paras):
    """'유의사항' 대표 표의 시작/끝 문단 번호. 시작 다음에 '구분','내용' 머리글이 오는 곳을 찾는다.
    제목은 '유의사항' 또는 '유의사항 : …', 머리글은 '구 분'처럼 띄어 써도 된다."""
    def heads(i):
        return [re.sub(r'\s+', '', q['text']) for q in paras[i + 1:i + 14]]
    best = None
    for i, p in enumerate(paras):
        t = p['text'].strip()
        if not (re.sub(r'\s+', '', t) == '유의사항' or re.match(r'유의사항\s*[:：]', t)):
            continue
        nxt = heads(i)
        if '구분' in nxt and '내용' in nxt:
            best = i
    if best is None:
        return None
    start = best + heads(best).index('내용') + 2
    base = paras[start - 1]['level']
    end = len(paras)
    for j in range(start, len(paras)):
        t = paras[j]['text'].strip()
        if paras[j]['level'] < base:
            end = j; break
        if re.fullmatch(r'\d{1,2}|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ]', t) and paras[j]['level'] <= base:
            end = j; break
    return start, end, base


def parse_cautions(paras):
    rng = find_caution_range(paras)
    if not rng:
        return []
    start, end, base = rng
    sections, cur, labels, last_item = [], None, [], None
    for p in paras[start:end]:
        raw, lvl = _clean(p['text']), p['level']
        t = raw.strip()
        if not t:
            continue
        if lvl > base:                      # 표 안의 표 (학교 개교시기 표 등)
            if cur is not None:
                if not cur.tables or cur.tables[-1]['after'] != len(cur.units):
                    cur.tables.append({'after': len(cur.units), 'cells': []})
                cur.tables[-1]['cells'].append(t)
            continue
        is_item = t.startswith('■')
        is_indented = raw[:1].isspace()
        if not is_item and not is_indented and not t.startswith('-'):
            # 왼쪽 '구분' 칸의 글자
            if last_item is not None or cur is None:
                labels = []
            labels.append(t)
            dong = bool(DONG_RE.match(t))
            cur = Section(t if dong else ' '.join(labels), dong)
            if sections and sections[-1].units == [] and not sections[-1].dong and not dong:
                sections[-1] = cur          # 아직 항목이 없는 구분 칸이 이어지면 이름만 합친다
            else:
                sections.append(cur)
            last_item = None
            continue
        if cur is None:
            cur = Section('(구분 없음)'); sections.append(cur)
        if is_item:
            last_item = Unit('item', raw.lstrip(), cur.label); cur.units.append(last_item)
            continue
        # 들여쓴 줄: '-' 하위항목, 괄호 보충줄, 또는 앞 문장의 이어지는 줄
        if t.startswith('-') or re.match(r'\d\)', t):
            last_item = Unit('sub', '   ' + t, cur.label); cur.units.append(last_item)
        elif t.startswith('(') and last_item is not None and last_item.text.rstrip().endswith(TERMINAL):
            last_item = Unit('sub', '   ' + t, cur.label); cur.units.append(last_item)
        elif last_item is not None:
            last_item.text = join_lines(last_item.text, raw.lstrip())
        else:
            last_item = Unit('item', t, cur.label); cur.units.append(last_item)
    for s in sections:
        for u in s.units:
            u.text = u.text.rstrip()
    return [s for s in sections if s.units or s.tables]


def join_lines(prev, nxt):
    """한글 문서에서 엔터로 끊어 놓은 줄을 잇는다. 앞줄이 공백으로 끝났으면 그대로, 아니면 붙인다."""
    if prev.endswith(' '):
        return prev + nxt
    return prev + nxt


def extract_fields(paras):
    texts = [p['text'].strip() for p in paras]
    f = {}
    for t in texts:
        if '입주자모집공고' in t and len(t) < 80 and '승인' not in t:
            name = re.sub(r'\([^)]*\)', '', t).replace('입주자모집공고', '').strip()
            f['complex_name'] = name
            break
    for t in texts:
        m = re.match(r'■?\s*공급\s*위치\s*:\s*(.+)', t)
        if m:
            f['location'] = m.group(1).strip()
            b = re.search(r'([A-Za-z]{1,4}\s*-?\s*\d{1,3})\s*(?:블록|블럭|BL)', f['location'])
            if b:
                f['block'] = re.sub(r'\s+', '', b.group(1))
            break
    for t in texts:
        if '입주시기' in t or '입주예정' in t:
            m = re.search(r'(\d{4})\s*년\s*(\d{1,2})\s*월', t)
            if m:
                f['move_in'] = f'{m.group(1)}년 {int(m.group(2)):02d}월'
                break
    # 공급금액 및 납부일정 표의 날짜 줄
    for i, t in enumerate(texts):
        key = re.sub(r'\s+', '', t)
        if '납부일정' in key and ('공급금액' in key or '분양금액' in key):
            for j in range(i, min(i + 120, len(texts))):
                if re.fullmatch(r'계약\s*(체결)?\s*시', texts[j]):
                    dates = []
                    for k in range(j + 1, min(j + 15, len(texts))):
                        m = DATE_RE.match(texts[k])
                        if not m:
                            break
                        dates.append(f'{m.group(1)}년 {int(m.group(2)):02d}월 {int(m.group(3)):02d}일')
                    f['dates'] = dates
                    break
            if 'dates' in f:
                break
    # 분양대금 납부계좌
    for i, t in enumerate(texts):
        if '납부계좌' in t and ('분양대금' in t or '공급대금' in t) and len(t) < 40:
            bank = acct = holder = None
            for k in range(i + 1, min(i + 25, len(texts))):
                s = texts[k]
                if bank is None and re.search(r'(은행|농협|수협|새마을금고|신협|뱅크)$', s) and 2 < len(s) < 15:
                    bank = s
                elif acct is None and ACCT_RE.match(s):
                    acct = s
                elif acct and holder is None and re.search(r'신탁|\(주\)|㈜|주식회사|공사|금융센터', s) and len(s) < 30:
                    holder = s
                if bank and acct and holder:
                    break
            if acct:
                f['bank'], f['account'], f['holder'] = bank or '', acct, holder or ''
                break
    return f


def parse_notice(paras):
    return {'fields': extract_fields(paras), 'sections': parse_cautions(paras)}
