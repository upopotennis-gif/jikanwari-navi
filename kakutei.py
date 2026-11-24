"""確定版の特別時間割「教室配当」PDF（教室ごとの表）を読む。
→ {日付: {時限: [(教室, 表記)]}}"""
import re, statistics
import fitz


def parse_rooms(path, year=2026, weekly=False):
    """weekly=True … 見出しが曜日（月〜金）の全体時間割の教室配当。日付の代わりに '月' などを鍵にする"""
    page = fitz.open(path)[0]
    Z = str.maketrans('０１２３４５６７８９', '0123456789')
    words = [w[:4] + (w[4].translate(Z) if re.match(r'[0-9０-９]{1,2}月', w[4]) else w[4],) + w[5:] for w in page.get_text('words')]   # 見出しが「12月１日(火)」の形のとき（2026-11-24）
    heads = sorted([w for w in words if (re.fullmatch(r'[月火水木金]', w[4]) and w[1] < 100 if weekly else re.match(r'\d{1,2}(/|月)\d{1,2}日?\(', w[4]))], key=lambda w: w[0])
    hy = max(w[3] for w in heads)
    nums = sorted([w for w in words if hy < w[1] < hy + 40 and re.fullmatch(r'[1-7]', w[4])], key=lambda w: w[0])
    ny = statistics.median(w[1] for w in nums)
    nums = [w for w in nums if abs(w[1] - ny) < 3]
    groups, cur = [], []
    for w in nums:
        if cur and int(w[4]) <= int(cur[-1][4]):
            groups.append(cur); cur = []
        cur.append(w)
    groups.append(cur)
    cols = []
    for g, h in zip(groups, heads):
        m = re.match(r'(\d{1,2})[/月](\d{1,2})', h[4])
        date = h[4] if weekly else f'{year}-{int(m.group(1)):02d}-{int(m.group(2)):02d}'
        cols += [((w[0] + w[2]) / 2, date, int(w[4])) for w in g]
    left = min(c[0] for c in cols) - 15
    right = max(c[0] for c in cols) + 15
    # 教室名（左端の列）
    names = [w for w in words if w[1] > ny + 5 and w[2] < left and w[0] > 60 and len(w[4]) >= 2]
    rows = sorted(((w[1] + w[3]) / 2, re.sub(r'\s', '', w[4]).replace('Ｈ', 'H').replace('Ｂ', 'B')) for w in names)
    gap = statistics.median([b[0] - a[0] for a, b in zip(rows, rows[1:])])
    cells = {}
    for w in words:
        cx, cy = (w[0] + w[2]) / 2, (w[1] + w[3]) / 2
        if w[1] <= ny + 5 or not (left <= cx <= right):
            continue
        c = min(cols, key=lambda c: abs(c[0] - cx))
        if abs(c[0] - cx) > 16:
            continue
        r = min(rows, key=lambda r: abs(r[0] - cy))
        if abs(r[0] - cy) > gap * 0.6:
            continue
        cells.setdefault((r[1], c[1], c[2]), []).append(w)
    out = {}
    for (room, date, p), ws in cells.items():
        ws.sort(key=lambda w: (round(w[1]), w[0]))
        lab = '\n'.join(w[4] for w in ws)
        if not weekly:
            lab = re.sub(r'^精選', '数学精選', lab)   # 教室配当は「精選ＢＣ」、教員版は「数学精選ＢＣ」（2026-11-24）
        out.setdefault(date, {}).setdefault(p, []).append((room, lab))
    return out


if __name__ == '__main__':
    import sys
    r = parse_rooms(sys.argv[1] if len(sys.argv) > 1 else 'source/kakutei/10-3_教室配当.pdf')
    for d in sorted(r):
        for p in sorted(r[d]):
            print(d, p, ' | '.join(f"{a}:{b.replace(chr(10),'/')}" for a, b in r[d][p]))
