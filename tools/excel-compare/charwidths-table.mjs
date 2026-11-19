// Writes src/ooxml/excelwidths.js from gdi-charwidths.ps1's measurement: for each font, regular, bold, italic and bold
// italic, the whole-pixel width Windows gives each printable ASCII character (32 to 126) at each pixel size from 11
// to 48 (8 to 36pt; half points that round to one pixel size have one set of widths, checked here), and Small Fonts'
// at 5 to 10px, which Excel lays every font out in under 8pt. One string a font and style, character by character,
// a width w as the character 48 + w; an italic that has its upright font's widths (Windows slants a font without an
// italic of its own) isn't written twice.
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
const lines = [];
let small = null;
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
    packed[style] = pack(byPx, PX_FROM, PX_TO, `${font} ${style}`);
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
  const parts = [`regular: ${JSON.stringify(packed.regular)}`, `bold: ${JSON.stringify(packed.bold)}`];
  if (packed.italic !== packed.regular) parts.push(`italic: ${JSON.stringify(packed.italic)}`);
  if (packed.boldItalic !== packed.bold) parts.push(`boldItalic: ${JSON.stringify(packed.boldItalic)}`);
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

/** Font (lower case) -> { regular, bold, italic, boldItalic, more }: per character (${FIRST} to ${LAST}), its width at each pixel
 *  size, as the character 48 + width; italic and boldItalic only where they aren't the upright font's. \`more\`: what
 *  a line in italic and in bold italic takes beyond its characters at each pixel size, as the character 48 + px. */
const WIDTHS = {
${lines.join('\n')}
};

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
  let s;
  let n;
  let at;
  let more = 0;
  if (size <= SMALL_TO) {
    at = Math.max(SMALL_FROM, size) - SMALL_FROM;
    s = SMALL[style].widths;
    n = SMALL_TO - SMALL_FROM + 1;
    more = SMALL[style].more[at];
  } else {
    const table = WIDTHS[String(name).toLowerCase()];
    if (!table || size > PX_TO) return null;
    at = size - PX_FROM;
    s = table[style] || (bold ? table.bold : table.regular);
    n = PX_TO - PX_FROM + 1;
    if (italic) more = table.more[bold ? 1 : 0].charCodeAt(at) - 48;
  }
  const width = (code) => (code >= FIRST && code <= LAST ? s.charCodeAt((code - FIRST) * n + at) - 48 : -1);
  return { width, digit: width(48), more, small: size <= SMALL_TO };
}
`;
fs.writeFileSync('src/ooxml/excelwidths.js', out);
console.log(`${lines.length} fonts, ${out.length} characters`);
