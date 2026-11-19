// Writes src/ooxml/excelwidths.js from gdi-charwidths.ps1's measurement: for each font, regular, bold, italic and bold
// italic, the whole-pixel width Windows gives each printable ASCII character (32 to 126) at each pixel size from 11
// to 48 (8 to 36pt; half points that round to one pixel size have one set of widths, checked here), and Small Fonts'
// at 5 to 10px, which Excel lays every font out in under 8pt. Each character's widths are written as a base width
// and the sizes it doesn't give (encode, below: a fifth of the space of every width written out); a style with the
// widths of the one it falls back to (an italic Windows slants of its upright font, a bold the same as regular) isn't
// written twice. Small Fonts, six sizes, are written out: a width w as the character 48 + w.
// With it, what an italic line takes beyond its characters in Excel (test/italiclines-excel.json: the narrowest
// column that keeps five italic zeros on a line, excel-wrapwidths.ps1 -Mode narrowest on wrapwidths-cases.mjs's
// italic cases), a number of pixels for each font, weight and pixel size.
// usage: node tools/excel-compare/charwidths-table.mjs <gdi-charwidths.json> [<excel-wrapwidths italic output.json>]
//        (the second file, when given, is first written compactly as test/italiclines-excel.json)
import fs from 'node:fs';

const FIRST = 32;
const LAST = 126;
const PX_FROM = 11;
const PX_TO = 48;
const SMALL_FROM = 5;
const SMALL_TO = 10;
const STYLES = ['regular', 'bold', 'italic', 'boldItalic'];
const data = JSON.parse(fs.readFileSync(process.argv[2], 'utf8').trimStart());
if (process.argv[3]) {
  // { font: { italic: [px of the narrowest column, by half point from 8pt], boldItalic: [...] } }
  const raw = JSON.parse(fs.readFileSync(process.argv[3], 'utf8').trimStart());
  const fonts = {};
  for (const c of raw.cases) {
    if (c.text !== '00000' || !c.italic) throw new Error('not the italic cases');
    const f = (fonts[c.font] = fonts[c.font] || { italic: [], boldItalic: [] });
    f[c.bold ? 'boldItalic' : 'italic'][Math.round((c.size - 8) * 2)] = c.px;
  }
  fs.writeFileSync('test/italiclines-excel.json', `${JSON.stringify({ excel: raw.excel, text: '00000', from: 8, fonts })}\n`);
}
const italicLines = JSON.parse(fs.readFileSync('test/italiclines-excel.json', 'utf8').trimStart());
const pack = (byPx, from, to, what) => {
  let s = '';
  for (let ch = FIRST; ch <= LAST; ch++) {
    for (let px = from; px <= to; px++) {
      const w = byPx.get(px)?.[ch - 32];
      if (w == null || w < 0 || 48 + w > 0xffff) throw new Error(`${what} ${px}px: no width for ${ch}`);
      s += String.fromCharCode(48 + w);
    }
  }
  return s;
};
/** What a line of five zeros `zero` px wide each takes beyond them when its narrowest column is `column` px: the
 *  line fits when width + more + 2 x ceil(zero / 4) + max(3, ceil((width + more) / 28)) <= column (rowheights.js). */
const moreOf = (zero, column, what) => {
  const need = (more) => 5 * zero + more + 2 * Math.ceil(zero / 4) + Math.max(3, Math.ceil((5 * zero + more) / 28));
  for (let more = 0; more < 40; more++) if (need(more) === column) return more;
  throw new Error(`${what}: five zeros of ${zero}px in ${column}px fit no allowance`);
};
/** A font and style's widths at PX_FROM to PX_TO, as [base, exceptions]. Most of a character's widths are its base
 *  width b (in 1/2048ths of the pixel size) times the size, rounded: (b x px + 1024) >> 11. `base` is b for each
 *  character as two characters, 48 + (b >> 6) and 48 + (b & 63); `exceptions`, for each character, what its other
 *  widths are off by: the size's index as 48 + i and the difference as 80 + d, a space after each character's. Each
 *  character's b is the one that leaves it fewest exceptions (the least such b). Decoded here and checked. */
const N = PX_TO - PX_FROM + 1;
const round = (b, px) => (b * px + 1024) >> 11;
const encode = (byPx, what) => {
  const grid = unpack(pack(byPx, PX_FROM, PX_TO, what), N);
  let base = '';
  let exceptions = '';
  for (const [k, row] of grid.entries()) {
    let best = null;
    const mid = Math.round((row[N - 1] * 2048) / PX_TO);
    for (let b = Math.max(0, mid - 256); b <= Math.min(4095, mid + 256); b++) {
      let miss = 0;
      for (let i = 0; i < N; i++) if (round(b, PX_FROM + i) !== row[i]) miss++;
      if (!best || miss < best.miss) best = { b, miss };
    }
    base += String.fromCharCode(48 + (best.b >> 6), 48 + (best.b & 63));
    for (let i = 0; i < N; i++) {
      const d = row[i] - round(best.b, PX_FROM + i);
      if (d < -32 || d > 46) throw new Error(`${what}: character ${FIRST + k} at ${PX_FROM + i}px is ${d}px off its base width`);
      if (d) exceptions += String.fromCharCode(48 + i, 80 + d);
    }
    exceptions += ' ';
  }
  if (decode(base, exceptions).join() !== grid.flat().join()) throw new Error(`${what}: doesn't decode to its widths`);
  return [base, exceptions];
};
const unpack = (s, n) => Array.from({ length: LAST - FIRST + 1 }, (_, k) => Array.from({ length: n }, (_, i) => s.charCodeAt(k * n + i) - 48));
// (The same as excelwidths.js's decode, below.)
const decode = (base, exceptions) => {
  const w = new Uint8Array((LAST - FIRST + 1) * N);
  for (let k = 0; k <= LAST - FIRST; k++) {
    const b = ((base.charCodeAt(2 * k) - 48) << 6) | (base.charCodeAt(2 * k + 1) - 48);
    for (let i = 0; i < N; i++) w[k * N + i] = round(b, PX_FROM + i);
  }
  for (let k = 0, j = 0; j < exceptions.length; j++) {
    if (exceptions[j] === ' ') { k++; continue; }
    w[k * N + exceptions.charCodeAt(j) - 48] += exceptions.charCodeAt(++j) - 80;
  }
  return [...w];
};
const lines = [];
let small = null;
// (Every font's styles in two strings, the bases and the exceptions, like with like, which gzip packs smaller; a
// style is its index in them, and two fonts or styles with the same widths share one.)
const styleIndex = new Map();
const bases = [];
const exceptionLists = [];
const indexOf = ([base, exceptions]) => {
  const key = `${base}|${exceptions}`;
  if (!styleIndex.has(key)) { styleIndex.set(key, bases.length); bases.push(base); exceptionLists.push(exceptions); }
  return styleIndex.get(key);
};
for (const [font, entry] of Object.entries(data)) {
  if (font === 'small fonts') {
    small = {};
    for (const style of STYLES) {
      const byPx = new Map(Object.entries(entry[style].px).map(([px, w]) => [Number(px), w]));
      const over = [];
      for (let px = SMALL_FROM; px <= SMALL_TO; px++) over.push(entry[style].overhang[px]);
      small[style] = { widths: pack(byPx, SMALL_FROM, SMALL_TO, `${font} ${style}`), more: over };
    }
    continue;
  }
  const packed = {};
  for (const style of STYLES) {
    const byPx = new Map();
    for (const [pt, widths] of Object.entries(entry[style])) {
      const px = Math.round((Number(pt) * 4) / 3);
      if (byPx.has(px) && byPx.get(px).join() !== widths.join()) throw new Error(`${font} ${style}: ${pt}pt differs from another size of ${px}px`);
      byPx.set(px, widths);
    }
    packed[style] = indexOf(encode(byPx, `${font} ${style}`));
  }
  // (The italic allowance by pixel size, a digit each: sizes that round to one pixel size agree, checked.)
  const name = Object.keys(italicLines.fonts).find((n) => n.toLowerCase() === font);
  if (!name) throw new Error(`${font}: no italic lines measured (test/italiclines-excel.json)`);
  const more = {};
  for (const style of ['italic', 'boldItalic']) {
    const byPx = new Map();
    italicLines.fonts[name][style].forEach((column, k) => {
      const pt = italicLines.from + k / 2;
      const px = Math.round((pt * 4) / 3);
      const zero = entry[style][String(pt)][48 - 32];
      const m = moreOf(zero, column, `${font} ${style} ${pt}pt`);
      if (byPx.has(px) && byPx.get(px) !== m) throw new Error(`${font} ${style}: ${pt}pt's allowance differs from another size of ${px}px`);
      byPx.set(px, m);
    });
    let s = '';
    for (let px = PX_FROM; px <= PX_TO; px++) { if (!byPx.has(px)) throw new Error(`${font} ${style} ${px}px: no allowance`); s += String.fromCharCode(48 + byPx.get(px)); }
    more[style] = s;
  }
  // (A style with the widths of the one it falls back to isn't written: italic to regular, bold italic to bold, bold
  // to regular.)
  const parts = [`regular: ${packed.regular}`];
  if (packed.bold !== packed.regular) parts.push(`bold: ${packed.bold}`);
  if (packed.italic !== packed.regular) parts.push(`italic: ${packed.italic}`);
  if (packed.boldItalic !== packed.bold) parts.push(`boldItalic: ${packed.boldItalic}`);
  parts.push(`more: [${JSON.stringify(more.italic)}, ${JSON.stringify(more.boldItalic)}]`);
  lines.push(`  ${JSON.stringify(font)}: { ${parts.join(', ')} },`);
}
if (!small) throw new Error('no "small fonts" in the measurement');
const out = `// Generated by tools/excel-compare/charwidths-table.mjs from tools/excel-compare/gdi-charwidths.ps1's measurement:
// the width in whole pixels Windows (GDI) gives each printable ASCII character of each font, regular, bold, italic
// and bold italic, at each pixel size from ${PX_FROM} to ${PX_TO} (8 to 36pt), as Excel lays text out on screen, and of Small Fonts
// at ${SMALL_FROM} to ${SMALL_TO}px, the bitmap font Excel lays every font out in under 8pt. Excel breaks wrapped lines at the sum of
// these (editor/rowheights.js); the browser's fractional widths miss it by a varying few percent. With them, the
// pixels an italic line takes beyond its characters, measured in Excel (test/italiclines-excel.json). Don't edit:
// measure again and run the script.

const PX_FROM = ${PX_FROM};
const PX_TO = ${PX_TO};
const SMALL_FROM = ${SMALL_FROM};
const SMALL_TO = ${SMALL_TO};
const FIRST = ${FIRST};
const LAST = ${LAST};

const N = PX_TO - PX_FROM + 1;

/** Font (lower case) -> { regular, bold, italic, boldItalic, more }: each style's widths, as its index in BASES and
 *  EXCEPTIONS, only where they aren't those of the style it falls back to (bold to regular, italic to regular, bold
 *  italic to bold). \`more\`: what a line in italic and in bold italic takes beyond its characters at each pixel size,
 *  as the character 48 + px. */
const WIDTHS = {
${lines.join('\n')}
};

/** Most of a character's widths are its base width b (1/2048ths of the size) times the pixel size, rounded. Each
 *  style's b for each character, two characters each, 48 + (b >> 6) and 48 + (b & 63), the styles one after another. */
const BASES = ${JSON.stringify(bases.join(''))};
/** Each style's other widths, the styles split by |: for each character, a size's index as 48 + i and what its
 *  width is off the base's as 80 + d, a space after each character's. */
const EXCEPTIONS = ${JSON.stringify(exceptionLists.join('|'))};

/** A style's widths, decoded once: character (from ${FIRST}) x size (from ${PX_FROM}px). */
const decoded = new Map();
let exceptionLists = null;
function decode(index) {
  let w = decoded.get(index);
  if (w) return w;
  exceptionLists = exceptionLists || EXCEPTIONS.split('|');
  const exceptions = exceptionLists[index];
  const base = BASES.slice(index * 2 * (LAST - FIRST + 1), (index + 1) * 2 * (LAST - FIRST + 1));
  w = new Uint8Array((LAST - FIRST + 1) * N);
  for (let k = 0; k <= LAST - FIRST; k++) {
    const b = ((base.charCodeAt(2 * k) - 48) << 6) | (base.charCodeAt(2 * k + 1) - 48);
    for (let i = 0; i < N; i++) w[k * N + i] = (b * (PX_FROM + i) + 1024) >> 11;
  }
  for (let k = 0, j = 0; j < exceptions.length; j++) {
    if (exceptions[j] === ' ') { k++; continue; }
    w[k * N + exceptions.charCodeAt(j) - 48] += exceptions.charCodeAt(++j) - 80;
  }
  decoded.set(index, w);
  return w;
}

/** Small Fonts in each style: per character, its width at ${SMALL_FROM} to ${SMALL_TO}px, and what a line takes beyond its characters
 *  at each (the overhang Windows gives a bold or an italic it makes up of a bitmap font). */
const SMALL = ${JSON.stringify(small)};

/** Excel's character widths for font \`name\` at \`px\` pixels (its size in points x 4/3), bold or italic or both:
 *  { width(code), digit, more, small }, or null when the table hasn't the font or the size. \`width(code)\` is a
 *  character's whole pixels, -1 for one outside printable ASCII; \`digit\` the width of 0; \`more\` the pixels a line
 *  takes beyond its characters (an italic's, and a made-up bold's under 8pt). Sizes that round to ${SMALL_TO}px or less
 *  are Small Fonts' whatever the font (\`small\`). */
export function excelWidths(name, px, bold = false, italic = false) {
  const size = Math.round(px);
  const style = bold ? (italic ? 'boldItalic' : 'bold') : italic ? 'italic' : 'regular';
  let width;
  let more = 0;
  if (size <= SMALL_TO) {
    const at = Math.max(SMALL_FROM, size) - SMALL_FROM;
    const s = SMALL[style].widths;
    const n = SMALL_TO - SMALL_FROM + 1;
    more = SMALL[style].more[at];
    width = (code) => (code >= FIRST && code <= LAST ? s.charCodeAt((code - FIRST) * n + at) - 48 : -1);
  } else {
    const table = WIDTHS[String(name).toLowerCase()];
    if (!table || size > PX_TO) return null;
    const at = size - PX_FROM;
    const w = decode(table[style] ?? (bold ? table.bold : null) ?? table.regular);
    if (italic) more = table.more[bold ? 1 : 0].charCodeAt(at) - 48;
    width = (code) => (code >= FIRST && code <= LAST ? w[(code - FIRST) * N + at] : -1);
  }
  return { width, digit: width(48), more, small: size <= SMALL_TO };
}
`;
fs.writeFileSync('src/ooxml/excelwidths.js', out);
console.log(`${lines.length} fonts, ${out.length} characters`);
