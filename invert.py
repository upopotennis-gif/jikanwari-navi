"""先生ごとの表（「科目\\n対象」の表記）を、クラスごとの時間割に組み直す。"""
import re
from collections import Counter

FW = str.maketrans('０１２３４５６７８９ＡＢＣＤＥＦＬＨＲ', '0123456789ABCDEFLHR')
CLASS_IDS = ['1A', '1B', '1C', '1D', '1E', '2A', '2B', '2C', '2D', '2E', '2F',
             '3A服飾', '3A食物', '3B', '3C', '3D', '3E', '3F']
FUTSU = {'1C', '1D', '1E', '2C', '2D', '2E', '2F', '3C', '3D', '3E', '3F'}
BLOCK_CLASSES = {'2文Ⅰ': ['2C', '2D'], '2理': ['2E', '2F'], '2福': ['2B'], '3福': ['3B'],
                 '3生': ['3A服飾', '3A食物']}
for k in 'アイウエオカキ':
    BLOCK_CLASSES[k] = ['3C', '3D']
for k in 'サシスセソタチツテ':
    BLOCK_CLASSES[k] = ['3E', '3F']
ALIAS = {'生支技術': '生活支援技術', 'コミュ技術': 'コミュ技', '政治経済': '政治・経済', '科学と人間': '科学と人間生活',
         'フード': 'フードデザイン', 'ここから理解': 'こころとからだの理解', 'デザイン': 'デザイン文化', '栄養': '栄養文化'}
# 3A服飾・食物: (自分のコース, 相手コースだけの科目)
COURSE_ONLY = {'3A服飾': ('服飾', {'栄養文化', '調理'}), '3A食物': ('食物', {'デザイン文化', 'Ｆ造形'})}
ART = {'音楽', '美術', '書道', '音楽Ⅰ', '美術Ⅰ', '書道Ⅰ'}


def ids_of(grade, letters):
    out = []
    for L in letters:
        if grade == '3' and L == 'A':
            out += ['3A服飾', '3A食物']
        elif grade + L in CLASS_IDS:
            out.append(grade + L)
    return out


def parse_label(label):
    """→ [(クラスid, 種別, 値)] 種別: 'subj' 'block' 'grade'"""
    lines = [l.strip() for l in str(label).split('\n') if l.strip()]
    if len(lines) == 1:  # 「介護福祉基礎１Ｂ」のように1行で書かれた表記
        m = re.fullmatch(r'(.+?)([１２３123][Ａ-ＦA-F]+)', lines[0].replace(' ', ''))
        if not m:
            return []
        lines = [m.group(1), m.group(2)]
    if len(lines) < 2:
        return []
    subj = ''.join(lines[:-1]).replace(' ', '')
    subj = ALIAS.get(subj, subj)
    tgt = lines[-1].replace(' ', '').translate(FW)
    if subj.startswith('(') or subj.startswith('（'):
        return []
    out = []
    m = re.fullmatch(r'([123])([A-F]+)', tgt)
    if m:
        for c in ids_of(m.group(1), m.group(2)):
            if c in COURSE_ONLY and subj in COURSE_ONLY[c][1]:
                continue  # 3A: 服飾と食物で分かれる科目
            if subj in ART:
                out.append((c, 'block', '芸術'))
            else:
                out.append((c, 'subj', subj))
        return out
    m = re.fullmatch(r'([123])年', tgt)
    if m:
        return [(c, 'grade', subj) for c in CLASS_IDS if c[0] == m.group(1) and (subj != '総合' or c in FUTSU)]
    m = re.fullmatch(r'(.+?)選択(\d[A-F])?', tgt)
    if m:
        b, cls = m.group(1), m.group(2)
        if b == 'Ⅱ' and cls:
            return [(cls, 'subj', subj + '|Ⅱ選択')]
        if b == 'Ⅱ':
            return [(c, 'subj', subj + '|Ⅱ選択') for c in ('2C', '2D')]
        key = {'Ⅰ': '2文Ⅰ', '2福': '2福', '3福': '3福', '3生': '3生', '2理系': '2理', '理系': '2理'}.get(b)
        keys = [key] if key else [ch for ch in b if ch in BLOCK_CLASSES]
        for k in keys:
            for c in BLOCK_CLASSES[k]:
                out.append((c, 'block', k))
        return out
    return []


def invert(labels):
    """labels: [(時限, 表記)] → {クラスid: {時限: エントリ}}"""
    hits = {}
    for p, lab in labels:
        for c, kind, val in parse_label(lab):
            hits.setdefault(c, {}).setdefault(p, []).append((kind, val))
    out = {}
    for c, per in hits.items():
        for p, hs in per.items():
            subs = [v for k, v in hs if k == 'subj']
            blks = [v for k, v in hs if k == 'block']
            grd = [v for k, v in hs if k == 'grade']
            if subs:
                uniq = list(dict.fromkeys(subs))
                ent = ('subj', uniq)
            elif blks:
                ent = ('block', Counter(blks).most_common(1)[0][0])
            else:
                ent = ('subj', list(dict.fromkeys(grd)))
            out.setdefault(c, {})[p] = ent
    return out
