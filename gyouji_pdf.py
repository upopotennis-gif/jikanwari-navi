"""月間行事予定（PDF）を 日付ごと・欄ごとの文字にする（教員用）。"""
import re
import fitz

# 欄の左端（R8 の様式）
COLS = [('学校行事', 60), ('生活科学科', 186), ('福祉科', 258), ('諸会議・研修', 338), ('PTA他', 375),
        ('附属中', 400), ('定時制', 438), ('授業等', 470)]
ABBR = {'学': '学習会', '海': '海外研修事前学習会', 'ジャンプ': 'ジャンプアップセミナー', '部': '部長会', '専': '専門委員会',
        '協': '生徒協議会', '拡協': '拡大協議会', '体実': '体育祭実行委員会', '神実': '神無祭実行委員会', '神責': '神無祭責任者会',
        '卒ア': '卒業アルバム委員会', '校': '校則見直し検討委員会', '修': '修学旅行委員会', '職': '職員会議'}


def col_of(x):
    name = None
    for n, x0 in COLS:
        if x >= x0 - 1:
            name = n
    return name


def parse(path, year, month):
    page = fitz.open(path)[0]
    words = page.get_text('words')
    head = next(w for w in words if w[4].startswith('日曜'))
    foot = min((w[1] for w in words if '略称' in w[4]), default=page.rect.height)
    days = sorted([((w[1] + w[3]) / 2, int(w[4])) for w in words
                   if w[0] < 52 and re.fullmatch(r'\d{1,2}', w[4]) and head[3] < w[1] < foot], key=lambda d: d[0])
    out = {}
    for i, (cy, d) in enumerate(days):
        top = (days[i - 1][0] + cy) / 2 if i else head[3] + 1
        bot = (cy + days[i + 1][0]) / 2 if i + 1 < len(days) else foot
        items = {}
        for w in words:
            wy = (w[1] + w[3]) / 2
            if not (top <= wy < bot) or w[0] < 60:
                continue
            c = col_of(w[0])
            if c in ('授業等', '定時制') or re.fullmatch(r'[〇○]+', w[4]):
                continue
            items.setdefault(c, []).append(w)
        day = []
        for c, _ in COLS:
            if c not in items:
                continue
            ws = sorted(items[c], key=lambda w: (round(w[1] / 3), w[0]))
            for w in ws:
                t = w[4].replace('　', ' ').strip()
                if not t:
                    continue
                # 狭い欄で折り返された切れはし（「会」など）は捨てる
                if len(t) <= 1 and t not in ABBR:
                    continue
                parts = [ABBR.get(x, x) for x in re.split(r'/', t)] if re.fullmatch(r'[学海部専協校修送職拡体神実責卒アジャンプ/]+', t) else [t]
                for x in parts:
                    day.append([c, x])
        texts = [x[1] for x in day]
        day = [x for x in day if not any(t != x[1] and t.startswith(x[1]) for t in texts)]
        out[f'{year}-{month:02d}-{d:02d}'] = day
    return out


if __name__ == '__main__':
    import sys
    for d, v in parse('source/gyouji_10.pdf', 2026, 10).items():
        print(d, v)
