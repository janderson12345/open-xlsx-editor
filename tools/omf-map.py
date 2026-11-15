# Writes src/editor/omf.js: the Microsoft font families the open-metric-fonts packages (OFL, metric-compatible: the
# originals' character widths and line spacing) stand in for, where each face's web fonts are, and each family's
# vertical metrics, so the editor can draw a workbook's Tahoma, Garamond or Century Gothic in them when the computer
# lacks the original (omffonts.js, fonts.js). Numbers and names only; no font is copied.
#
# Ported from open-docx-editor's tools/fonts/build-omf.py (commit 338fa5a, MIT, same owner): the same map, with the
# metrics the cell layout uses (the shape of src/engine/metrics.js) in place of Word's line height.
#
# usage: python tools/omf-map.py <open-metric-fonts checkout, with dist/npm built by its packaging>
#
# Read from the checkout: each package's fontconfig aliases (the Microsoft family a family of ours stands in for), our
# fonts' name tables (which file is a family's Regular, Italic, Bold and Bold Italic, as Office links them), their
# vertical metrics (the Microsoft font's, which ours copies), and the npm package's index.css (each face's WOFF2 files,
# one per alphabet, and the characters each holds).
import glob
import json
import os
import re
import sys
from xml.etree import ElementTree

from fontTools.ttLib import TTFont

root = sys.argv[1]
here = os.path.dirname(os.path.abspath(__file__))
out_path = os.path.join(here, '..', 'src', 'editor', 'omf.js')


def slug(stem):
    return re.sub(r'[^a-z0-9]+', '-', stem.lower()).strip('-')


def metrics(font):
    """The font's vertical metrics per em, as tools/fontmetrics.mjs reads an installed font's: hhea line and ascent,
    hhea [ascent, descent, gap], Windows [ascent, descent], typo [ascent, descent, gap, USE_TYPO_METRICS]."""
    os2, hhea, upm = font['OS/2'], font['hhea'], font['head'].unitsPerEm
    r = lambda v: round(v / upm, 4)
    return {
        'line': r(hhea.ascent - hhea.descent + hhea.lineGap), 'ascent': r(hhea.ascent + hhea.lineGap),
        'hhea': [r(hhea.ascent), r(-hhea.descent), r(hhea.lineGap)],
        'win': [r(os2.usWinAscent), r(os2.usWinDescent)],
        'typo': [r(os2.sTypoAscender), r(-os2.sTypoDescender), r(os2.sTypoLineGap), bool(os2.fsSelection & (1 << 7))],
    }


version = None
ranges = {}  # alphabet -> unicode-range
families = {}  # Microsoft family -> [package, metrics, faces]
for pkg_dir in sorted(glob.glob(os.path.join(root, 'fonts', '*'))):
    pkg = os.path.basename(pkg_dir)
    npm = os.path.join(root, 'dist', 'npm', pkg)
    meta = json.load(open(os.path.join(npm, 'package.json'), encoding='utf8'))
    if version and meta['version'] != version:
        raise SystemExit(f'{pkg} is {meta["version"]}, others {version}')
    version = meta['version']
    # Each WOFF2 file of the package: its face's file stem, its alphabet and the characters it holds.
    css = open(os.path.join(npm, 'index.css'), encoding='utf8').read()
    subsets = {}  # face file stem -> [alphabets]
    for name, body in re.findall(r'/\* ([\w-]+) \*/\s*@font-face \{(.*?)\}', css, re.S):
        src = re.search(r'url\(\./files/([\w-]+)\.woff2\)', body).group(1)
        rng = re.search(r'unicode-range: ([^;]+);', body).group(1).replace(' ', '')
        if not src.endswith('-' + name):
            raise SystemExit(f'{pkg}: {src} is not its alphabet {name}')
        # (An alphabet covering the same characters in every package is listed once; one that doesn't, as each
        # package's "other": its remaining characters, under "<package>:<alphabet>".)
        if ranges.setdefault(name, rng) != rng:
            ranges[f'{pkg}:{name}'] = rng
        subsets.setdefault(src[: -len(name) - 1], []).append(name)
    # Our families' faces as Windows links them (name ID 1 and 2): Regular and Italic at 400, Bold and Bold Italic at 700.
    linked = {}  # our family -> [(weight, italic, stem)]
    lines = {}
    regulars = set()  # families whose metrics are their Regular's
    for ttf in sorted(glob.glob(os.path.join(pkg_dir, '*.ttf'))):
        font = TTFont(ttf, lazy=True)
        fam = font['name'].getDebugName(1)
        stem = slug(os.path.splitext(os.path.basename(ttf))[0])
        if stem not in subsets:
            raise SystemExit(f'{pkg}: no web fonts for {stem}')
        # (Its style bits, as Windows links a family's four faces: Lucida Bright's bold is "Demibold", Lucida Sans
        # Typewriter's italic "Oblique".)
        sel = font['OS/2'].fsSelection
        weight = 700 if sel & 0x20 else 400
        linked.setdefault(fam, []).append((weight, bool(sel & 1), stem))
        # (The metrics from the family's Regular face, as a row of the font takes them: a Bold can have taller metrics,
        # Book Antiqua's and Century Gothic's do. Another face only where there's no Regular.)
        regular = not (sel & 0x21)
        if regular or fam not in lines:
            if regular or fam not in regulars:
                lines[fam] = metrics(font)
            if regular:
                regulars.add(fam)
    # The Microsoft families ours stand in for (the package's fontconfig aliases).
    for conf in glob.glob(os.path.join(pkg_dir, 'fontconfig', '*.conf')):
        for alias in ElementTree.parse(conf).getroot().iter('alias'):
            ms, ours = alias.find('family').text, alias.find('accept/family').text
            if ours not in linked:
                # (A name for the whole typographic family, Sitka or Walbaum, whose fonts are "<name> Text", "<name>
                # Display"...: its text size, the one for body text.)
                sizes = sorted(f for f in linked if f.startswith(ours + ' ') and len(linked[f]) > 0)
                pick = next((f for f in sizes if f == ours + ' Text'), sizes[0] if sizes else None)
                if not pick:
                    raise SystemExit(f'{pkg}: {ms} -> {ours}, which none of its fonts is')
                print(f'  {ms} -> {ours}: its fonts are {sizes}; using {pick}')
                ours = pick
            faces = sorted(linked[ours])
            families[ms] = [pkg, lines[ours], [[w, int(i), s, ','.join(subsets[s])] for w, i, s in faces]]

js = [
    '// Generated by tools/omf-map.py from open-metric-fonts ' + version + ' (https://github.com/open-metric-fonts/open-metric-fonts):',
    "// the Microsoft families its packages stand in for, with the originals' character widths and line metrics. Don't",
    '// edit: run the script again. Loaded by omffonts.js only for fonts a workbook shows that the computer lacks.',
    '',
    f"export const OMF_VERSION = '{version}';",
    '',
    '/** The characters each alphabet\'s WOFF2 file holds (its @font-face unicode-range); "<package>:<alphabet>" where a',
    ' * package\'s differs (its "other": the characters no alphabet above has). */',
    'export const OMF_RANGES = ' + json.dumps(ranges, indent=2) + ';',
    '',
    '/** Microsoft family -> [package (<omf folder>/<package>/files/), its Regular face\'s vertical metrics per em (as',
    ' * engine/metrics.js has an installed font\'s), faces: [weight, italic, file stem, alphabets]], the faces as Office',
    ' * links them (Regular and Italic 400, Bold and Bold Italic 700). */',
    'export const OMF = {',
]
for ms in sorted(families, key=str.lower):
    js.append(f'  {json.dumps(ms)}: {json.dumps(families[ms], separators=(",", ":"))},')
js.append('};')
open(out_path, 'w', encoding='utf8', newline='\n').write('\n'.join(js) + '\n')
print(f'{len(families)} Microsoft families from {len(glob.glob(os.path.join(root, "fonts", "*")))} packages, version {version}, alphabets {list(ranges)}')
