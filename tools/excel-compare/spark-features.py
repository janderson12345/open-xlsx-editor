# What Excel drew in each sparkline picture of its PDF (sparkimages.cjs's output), as numbers the drawing can be
# compared with (test/sparklines.test.js): in CSS pixels from the cell's top left (the picture covers the cell, at 200
# dpi), per picture: its size; bars (runs of picture columns holding ink: left, right, top, bottom, colour); blobs of
# each colour that isn't the most common one (markers, the axis: centre, width, height, colour); and the line's trace
# (for every other picture column, the darkness-weighted middle of the most common colour's ink).
# usage: python tools/excel-compare/spark-features.py <sparkimages outDir> <out.json>
import json
import sys

import numpy as np
from PIL import Image
from scipy import ndimage

S = 200 / 96
src, out = sys.argv[1], sys.argv[2]


def hexof(c):
    return '#%02X%02X%02X' % tuple(int(v) for v in c)


def features(path):
    a = np.asarray(Image.open(path).convert('RGB')).astype(int)
    h, w, _ = a.shape
    flat = a.reshape(-1, 3)
    keys, counts = np.unique(flat, axis=0, return_counts=True)
    pure = [(int(n), k) for k, n in zip(keys, counts) if k.sum() < 720 and n >= 12]
    pure.sort(key=lambda t: -t[0])
    main = pure[0][1] if pure else np.array([0, 0, 0])
    ink = np.clip((765 - a.sum(2)) / max(1, 765 - int(main.sum())), 0, 1)
    # Bars: runs of columns holding a pixel at least 40% inked, rows inked nearly all across (an axis) left out.
    solid = ink * ((ink > 0.4).mean(1) < 0.9)[:, None]
    on = (solid > 0.4).any(0)
    bars = []
    x = 0
    while x < w:
        if not on[x]:
            x += 1
            continue
        s = x
        while x < w and on[x]:
            x += 1
        rough = solid[:, (s + x) // 2]
        if not (rough > 0.05).any():
            continue
        # (Its colour where it's inkiest, and its coverage measured against that colour, so a light bar's edges come
        # out as a dark one's.)
        ry = np.nonzero(rough > 0.05)[0]
        row = ry.min() + int(np.argmax(solid[ry.min():ry.max() + 1, s:x].sum(1)))
        cc = a[row, (s + x) // 2]
        own = np.clip((765 - a.sum(2)) / max(1, 765 - int(cc.sum())), 0, 1) * (solid > 0)
        mid = own[:, (s + x) // 2]
        rows = np.nonzero(mid > 0.05)[0]
        y0, y1 = rows.min(), rows.max() + 1
        # (fractional edges from the partly covered pixels)
        top = y0 + (1 - mid[y0]) if y1 - y0 > 1 else y0
        bot = y1 - (1 - mid[y1 - 1]) if y1 - y0 > 1 else y0 + mid[y0]
        across = own[row, max(0, s - 1):min(w, x + 1)]
        left = max(0, s - 1) + float(np.argmax(across > 0.05)) + (1 - across[across > 0.05][0])
        bars.append([round(left / S, 2), round((left + across.sum()) / S, 2), round(top / S, 2), round(bot / S, 2), hexof(cc)])
    # Blobs of the other pure colours: not a tint of the main colour (its edges' antialiasing against white), nor near
    # one taken already.
    white = np.array([255, 255, 255])
    def tint_of(k, base):
        d = white - base
        t = np.dot(k - base, d) / max(1, np.dot(d, d))
        return np.abs(k - (base + t * d)).sum() < 40
    kept = []
    for n, k in pure[1:]:
        if tint_of(k, main) or any(np.abs(k - q).sum() < 90 or tint_of(k, q) for q in kept):
            continue
        kept.append(k)
    blobs = []
    for k in kept:
        near = np.abs(a - k).sum(2) < 90
        lab, m = ndimage.label(near)
        for i in range(1, m + 1):
            ys, xs = np.nonzero(lab == i)
            if len(ys) < 6:
                continue
            blobs.append([round((xs.mean() + .5) / S, 2), round((ys.mean() + .5) / S, 2), round((xs.max() - xs.min() + 1) / S, 2), round((ys.max() - ys.min() + 1) / S, 2), hexof(k)])
    # The trace of the main colour.
    near = np.abs(a - main).sum(2) < 200
    weight = np.where(near, ink, 0)
    trace = []
    for x in range(0, w, 2):
        cw = weight[:, x]
        trace.append(round(float((cw * (np.arange(h) + .5)).sum() / cw.sum()) / S, 2) if cw.sum() > 0.5 else None)
    return {'main': hexof(main), 'bars': bars, 'blobs': blobs, 'trace': trace, 'step': round(2 / S, 4)}


images = json.load(open(f'{src}/images.json'))
result = [dict(page=i['page'], w=i['w'], h=i['h'], **features(f"{src}/{i['n']}.png")) for i in images]
with open(out, 'w', encoding='utf-8', newline='\n') as f:
    f.write('[\n' + ',\n'.join(json.dumps(r, separators=(',', ':')) for r in result) + '\n]\n')
print(len(result), 'pictures')
