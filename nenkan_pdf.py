"""年間行事予定（1〜3月・教員用。3か月を横に並べた表）を日ごとに読む。
  python3 nenkan_pdf.py source/nenkan_1-3_案_2026-11-13.pdf  → 日ごとの 行事・略称・時数 を表示
月間行事予定（gyouji_pdf.py）とは表の形がちがう。長い矢印（考査の期間・海外研修）は読めないので 予定.csv で足す。"""
import datetime, re, sys
import fitz

AB = re.compile(r'^(職|送|専|学|海|協|拡協|体実|部|校|修|卒ア)(/(職|送|専|学|海|協|拡協|体実|部|校|修|卒ア))*$')
# (月, 行事の欄の左, 右＝時数の欄の左, 時数の欄の右)
COLS = [(1, 48, 193, 203), (2, 226, 371, 381), (3, 403, 548, 558)]


def parse(path, year=2027):
    page = fitz.open(path)[0]
    ws = page.get_text('words')
    days = sorted([((w[1] + w[3]) / 2, int(w[4])) for w in ws if w[0] < 32 and re.fullmatch(r'\d{1,2}', w[4])])
    cs = [d[0] for d in days]
    out = {}
    for w in ws:
        cy = (w[1] + w[3]) / 2
        if cy < cs[0] - 9 or cy > cs[-1] + 9:
            continue
        i = min(range(len(cs)), key=lambda k: abs(cs[k] - cy)) + 1
        t = w[4].replace('　', ' ').strip()
        if not t:
            continue
        for m, x0, x1, j1 in COLS:
            # 曜日とつながって読まれた行事（「火二次募集学力検査」）は曜日を取って行事にする
            if x0 - 14 <= w[0] < x0 - 2 and len(t) > 2 and t[0] in '月火水木金土日':
                out.setdefault((m, i), {'ev': [], 'ab': [], 'js': None})['ev'].append((round(w[1]), round(w[0]), t[1:].strip()))
                continue
            if x0 - 2 <= w[0] < x1:
                e = out.setdefault((m, i), {'ev': [], 'ab': [], 'js': None})
                (e['ab'] if AB.match(t) else e['ev']).append((round(w[1]), round(w[0]), t))
            elif x1 <= w[0] < j1 and re.fullmatch(r'\d', t):
                out.setdefault((m, i), {'ev': [], 'ab': [], 'js': None})['js'] = int(t)
    res = []
    for (m, i), e in sorted(out.items()):
        try:
            d = datetime.date(year, m, i)
        except ValueError:
            continue
        res.append({'d': d.isoformat(), 'w': '月火水木金土日'[d.weekday()], 'js': e['js'],
                    'ev': [t for _, _, t in sorted(e['ev'])],
                    'ab': sum([t.split('/') for _, _, t in sorted(e['ab'])], [])})
    return res


if __name__ == '__main__':
    for r in parse(sys.argv[1]):
        print(r['d'], r['w'], r['js'], ' ／ '.join(r['ev']), '|', ' '.join(r['ab']))
