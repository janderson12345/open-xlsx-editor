# Compares our cell text with Excel's per workbook (the JSON cells.cjs writes for each PDF) and prints a row per
# workbook and the totals: page counts, cell text on the same page, x within 1px, baseline y within 1px, both.
# (x is where left-aligned text starts and right-aligned text ends, so it measures column edges and cell padding;
# y measures row edges and the text baseline in the row.)
# usage: python compare.py <excelJsonDir> <oursJsonDir>
import sys, os, json, difflib
xd, od = sys.argv[1:3]
T = dict(cells=0, matched=0, samePage=0, x1=0, y1=0, both=0, w1=0); pc = n = 0
rows = []
for f in sorted(os.listdir(xd)):
    if not f.endswith('.json') or not os.path.exists(os.path.join(od, f)): continue
    x = json.load(open(os.path.join(xd, f), encoding='utf8')); o = json.load(open(os.path.join(od, f), encoding='utf8'))
    sm = difflib.SequenceMatcher(None, [a[0] for a in x], [a[0] for a in o], autojunk=False)
    pairs = [(x[b.a + i], o[b.b + i]) for b in sm.get_matching_blocks() for i in range(b.size)]
    same = [(a, b) for a, b in pairs if a[1] == b[1]]
    # Right-aligned text (numbers) is compared at its right end, left-aligned at its start: whichever end is closer.
    dx = [min(abs(b[2] - a[2]), abs((b[2] + b[4]) - (a[2] + a[4]))) for a, b in same]
    dy = [abs(b[3] - a[3]) for a, b in same]
    # Text width: how the font measures (the same text, the same font and size, should be as wide as Excel's).
    dw = [abs(b[4] - a[4]) for a, b in same]
    r = dict(cells=len(x), matched=len(pairs), samePage=len(same), x1=sum(v <= 1 for v in dx), y1=sum(v <= 1 for v in dy),
             both=sum(p <= 1 and q <= 1 for p, q in zip(dx, dy)), w1=sum(v <= 1 for v in dw))
    xp, op = max([a[1] for a in x] or [0]), max([a[1] for a in o] or [0])
    n += 1; pc += xp == op
    for k in T: T[k] += r[k]
    mx = sum(dx) / max(1, len(dx)); my = sum(dy) / max(1, len(dy))
    rows.append(f"{f[:-5][:40]:40} pages {xp:>3}/{op:<3} samePage {r['samePage']}/{r['matched']}/{r['cells']} x1px {r['x1']} y1px {r['y1']} both {r['both']} meanAbs x {mx:.1f} y {my:.1f} width1px {r['w1']}")
print('\n'.join(rows))
pct = lambda a, b: f"{100 * a / max(1, b):.1f}%"
print(f"TOTAL workbooks {n} page counts {pc}/{n} cells matched {T['matched']}/{T['cells']} on same page {T['samePage']}/{T['matched']} "
      f"x within 1px {T['x1']}/{T['samePage']} ({pct(T['x1'], T['samePage'])}) y within 1px {T['y1']}/{T['samePage']} ({pct(T['y1'], T['samePage'])}) "
      f"both {T['both']}/{T['samePage']} ({pct(T['both'], T['samePage'])}) width within 1px {T['w1']}/{T['samePage']} ({pct(T['w1'], T['samePage'])})")
