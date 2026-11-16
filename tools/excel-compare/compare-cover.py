# Compares the colours both PDFs show, point by point, inside Excel's filled shapes (shapes.cjs's JSON): Excel joins
# neighbouring cells of one colour into one rectangle where we draw one per cell, so compare-shapes.py's rectangle
# matching undercounts; here each point 1.5px inside an Excel shape, every 3px, takes the colour of the last shape
# drawn over it in each PDF (our page's background left out). Reports per workbook the points of the same colour,
# within 2 a channel, and the commonest differences (Excel's colour, ours, how many).
# usage: python compare-cover.py <excelShapesDir> <oursShapesDir>
import sys, os, json
from collections import Counter

xd, od = sys.argv[1:3]

def fills(path):
    out = [s for s in json.load(open(path, encoding='utf8')) if s[0] == 'fill']
    return [s for s in out if not (s[2] == '#ffffff' and s[5] > 700)]

def shown(shapes, page, x, y):
    for s in reversed(shapes):
        if s[1] == page and s[3] <= x <= s[3] + s[5] and s[4] <= y <= s[4] + s[6]:
            return s[2]
    return '#ffffff'

def near(a, b):
    return all(abs(int(a[i:i + 2], 16) - int(b[i:i + 2], 16)) <= 2 for i in (1, 3, 5))

for f in sorted(os.listdir(xd)):
    if not f.endswith('.json') or not os.path.exists(os.path.join(od, f)):
        continue
    x = fills(os.path.join(xd, f))
    o = fills(os.path.join(od, f))
    pts = same = close = 0
    diff = Counter()
    for s in x:
        page, x0, y0, w, h = s[1], s[3], s[4], s[5], s[6]
        if w < 3 or h < 3:
            # (A border's rectangle: its middle line.)
            xs = [x0 + w / 2] if w < 3 else [x0 + 1.5 + 3 * i for i in range(int((w - 3) / 3) + 1)]
            ys = [y0 + h / 2] if h < 3 else [y0 + 1.5 + 3 * i for i in range(int((h - 3) / 3) + 1)]
        else:
            xs = [x0 + 1.5 + 3 * i for i in range(int((w - 3) / 3) + 1)]
            ys = [y0 + 1.5 + 3 * i for i in range(int((h - 3) / 3) + 1)]
        for px in xs:
            for py in ys:
                a = shown(x, page, px, py)
                b = shown(o, page, px, py)
                pts += 1
                if a == b:
                    same += 1
                elif near(a, b):
                    close += 1
                else:
                    diff[(a, b)] += 1
    print(f"{f[:-5]:30} points {pts} same {same} ({100 * same / max(1, pts):.1f}%) within 2 {close} differ {pts - same - close}")
    for (a, b), n in diff.most_common(15):
        print(f"   Excel {a} ours {b}: {n}")
