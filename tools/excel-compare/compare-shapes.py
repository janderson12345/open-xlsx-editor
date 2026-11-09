# Compares our filled shapes with Excel's per workbook (the JSON shapes.cjs writes for each PDF): cell fills and the
# borders Excel draws as filled rectangles. An Excel rectangle counts as matched when one of ours on the same page has
# the same colour and every edge within 1px; also counted: those whose colour matches but not their place, and those
# not drawn at all. Excel's hairline strokes (outlines under its rectangles) are left out.
# usage: python compare-shapes.py <excelShapesDir> <oursShapesDir>
import sys, os, json
xd, od = sys.argv[1:3]
T = dict(shapes=0, matched=0, colour=0)
rows = []
for f in sorted(os.listdir(xd)):
    if not f.endswith('.json') or not os.path.exists(os.path.join(od, f)): continue
    x = [s for s in json.load(open(os.path.join(xd, f), encoding='utf8')) if s[0] == 'fill']
    o = [s for s in json.load(open(os.path.join(od, f), encoding='utf8')) if s[0] == 'fill']
    # (Our page background, a white rectangle over the whole page, isn't a cell's.)
    o = [s for s in o if not (s[2] == '#ffffff' and s[5] > 700)]
    used = set()
    m = c = 0
    for s in x:
        near = lambda t: t[1] == s[1] and all(abs(a - b) <= 1 for a, b in ((t[3], s[3]), (t[4], s[4]), (t[3] + t[5], s[3] + s[5]), (t[4] + t[6], s[4] + s[6])))
        hit = next((i for i, t in enumerate(o) if i not in used and t[2] == s[2] and near(t)), None)
        if hit is not None: used.add(hit); m += 1; continue
        if any(t[2] == s[2] and t[1] == s[1] for t in o): c += 1
    T['shapes'] += len(x); T['matched'] += m; T['colour'] += c
    if x: rows.append(f"{f[:-5][:40]:40} shapes {len(x):5} matched {m:5} colour-only {c:5} missing {len(x) - m - c:5} ours {len(o):5}")
print('\n'.join(rows))
print(f"TOTAL SHAPES {T['shapes']} matched within 1px {T['matched']} ({100 * T['matched'] / max(1, T['shapes']):.1f}%) colour only {T['colour']} missing {T['shapes'] - T['matched'] - T['colour']}")
