"""特別時間割（教員版PDF：先生ごとの表）を読み、日付ごとの [(時限, 表記)] にする。"""
import re, statistics
import fitz


def parse_pdf(path, year=2026):
    page = fitz.open(path)[0]
    words = page.get_text('words')
    hy = next(w[1] for w in words if w[4] == '分掌')
    nums = sorted([w for w in words if abs(w[1] - hy) < 2 and re.fullmatch(r'[1-7]', w[4])], key=lambda w: w[0])
    groups, cur = [], []
    for w in nums:
        if cur and int(w[4]) <= int(cur[-1][4]):
            groups.append(cur); cur = []
        cur.append(w)
    groups.append(cur)
    heads = [w for w in words if hy - 14 < w[1] < hy - 3 and re.match(r'\d{1,2}/\d{1,2}\(', w[4])]
    days = []
    for g in groups:
        gs, ge = g[0][0], g[-1][2]
        cand = [h for h in heads if gs - 30 <= h[0] <= ge]
        c = (gs + ge) / 2
        h = min(cand, key=lambda h: abs(h[0] - c))
        m = re.match(r'(\d{1,2})/(\d{1,2})\(.\)\s*(.*)', h[4])
        date = f'{year}-{int(m.group(1)):02d}-{int(m.group(2)):02d}'
        days.append({'date': date, 'note': re.sub(r'\s+', ' ', m.group(3).replace('　', ' ')).strip(),
                     'cols': [((w[0] + w[2]) / 2, int(w[4])) for w in g]})
    colx = [(x, di, p) for di, d in enumerate(days) for x, p in d['cols']]
    left = min(x for x, _, _ in colx) - 13
    right = max(x for x, _, _ in colx) + 13
    name_right = left - 1
    rows = sorted({round((w[1] + w[3]) / 2, 1) for w in words
                   if w[1] > hy + 4 and w[2] < name_right and w[0] > 85 and not re.fullmatch(r'[0-9]', w[4])})
    gap = statistics.median([b - a for a, b in zip(rows, rows[1:])])
    cells = {}
    for w in words:
        if w[1] <= hy + 3:
            continue
        cx, cy = (w[0] + w[2]) / 2, (w[1] + w[3]) / 2
        if not (left <= cx <= right):
            continue
        col = min(colx, key=lambda c: abs(c[0] - cx))
        if abs(col[0] - cx) > 13:
            continue
        r = min(rows, key=lambda r: abs(r - cy))
        if abs(r - cy) > gap * 0.6:
            continue
        cells.setdefault((r, col[1], col[2]), []).append(w)
    # 行ごとの先生の名前（氏名列）
    names = {}
    for w in words:
        if w[1] > hy + 4 and w[0] > 85 and w[2] < name_right and not re.fullmatch(r'[0-9]', w[4]):
            r = min(rows, key=lambda r: abs(r - (w[1] + w[3]) / 2))
            names.setdefault(r, []).append(w)
    names = {r: ''.join(x[4] for x in sorted(ws, key=lambda x: x[0])) for r, ws in names.items()}
    out = {d['date']: {'note': d['note'], 'periods': max(p for _, p in d['cols']), 'labels': [], 'teachers': {}} for d in days}
    for (r, di, p), ws in cells.items():
        ws.sort(key=lambda w: (round(w[1]), w[0]))
        lab = '\n'.join(w[4] for w in ws)
        o = out[days[di]['date']]
        o['labels'].append((p, lab))
        o['teachers'].setdefault(names.get(r, '?'), {})[p] = lab
    return out, gap, len(rows)


if __name__ == '__main__':
    import glob
    for f in sorted(glob.glob('source/tokubetsu/*.pdf')):
        res, gap, nr = parse_pdf(f)
        print('##', f, 'rows', nr, 'gap', gap)
        for d, v in res.items():
            print(d, v['periods'], v['note'], len(v['labels']))
