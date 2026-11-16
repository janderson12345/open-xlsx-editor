// A table pasted from another program (the clipboard's HTML: Excel, Word, Google Sheets, a browser), with its formats,
// as Excel's Paste takes it (measured over COM: tools/excel-compare/excel-clip-paste.ps1 on clip-html-cases.json,
// giving test/clip-html-excel.json; docs/CLIPBOARD.md). The cells' CSS (htmltable.js) becomes cell formats: fonts,
// colours, fills, borders, alignment, wrapping, indents and number formats (Excel's mso-number-format, Google Sheets'
// data-sheets-numberformat); colspan and rowspan become merged ranges; each cell's text is read as typed, or its value
// taken from x:num or data-sheets-value. Returns steps as clipops.js's paste functions do.
import { readHtmlTable, cssUnescape } from './htmltable.js';
import { parseInput, formatAfterEntry } from './input.js';
import { maxOf } from './listmath.js';
import { colPad } from './digitwidth.js';

const NAMED = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', lime: '00FF00', green: '008000', blue: '0000FF', yellow: 'FFFF00',
  aqua: '00FFFF', cyan: '00FFFF', fuchsia: 'FF00FF', magenta: 'FF00FF', silver: 'C0C0C0', gray: '808080', grey: '808080',
  maroon: '800000', olive: '808000', purple: '800080', teal: '008080', navy: '000080', orange: 'FFA500',
  darkred: '8B0000', darkblue: '00008B', darkgreen: '006400', lightgray: 'D3D3D3', lightgrey: 'D3D3D3',
};
// (Colours that mean "automatic": the system's window text, as Excel's own copy writes its borders and text.)
const AUTO = new Set(['windowtext', 'auto', 'black', 'currentcolor', 'initial', 'inherit']);

/** A CSS colour as RRGGBB, 'auto' for the automatic colour, or null for none or one not understood. */
export function cssColor(text) {
  const t = String(text || '').trim().toLowerCase();
  if (!t || t === 'transparent' || t === 'none') return null;
  if (AUTO.has(t)) return 'auto';
  let m = /^#([0-9a-f]{3})$/.exec(t);
  if (m) return m[1].split('').map((x) => x + x).join('').toUpperCase();
  if ((m = /^#([0-9a-f]{6})$/.exec(t))) return m[1].toUpperCase();
  if ((m = /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*(?:,\s*([\d.]+)\s*)?\)$/.exec(t))) {
    if (m[4] != null && Number(m[4]) === 0) return null;
    return [m[1], m[2], m[3]].map((x) => Math.min(255, Number(x)).toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  return NAMED[t] || null;
}
/** A CSS value's words, split at spaces outside brackets ("1px solid rgb(0, 0, 0)": three). */
function words(text) {
  const out = [];
  let depth = 0;
  let word = '';
  for (const ch of String(text).trim()) {
    if (ch === '(') depth++;
    if (ch === ')') depth = Math.max(0, depth - 1);
    if (/\s/.test(ch) && depth === 0) { if (word) out.push(word); word = ''; } else word += ch;
  }
  if (word) out.push(word);
  return out;
}
const colorOf = (hex) => (hex === 'auto' ? null : { rgb: `FF${hex}` });

/** A CSS length in points (pt, px at 96 a inch, in, cm, mm), or null. */
function points(text) {
  const m = /^(-?[\d.]+)(pt|px|in|cm|mm)?$/i.exec(String(text).trim());
  if (!m) return null;
  const n = Number(m[1]);
  return { pt: n, px: n * 0.75, in: n * 72, cm: n * 72 / 2.54, mm: n * 72 / 25.4 }[(m[2] || 'px').toLowerCase()];
}

/** A border side from CSS ("1px solid #000", ".5pt solid windowtext", "none"), as Excel takes it (measured: .5pt is
 *  thin, 1px and 1pt medium, 1.5pt and up thick, double double): { style, color } or null for none; undefined when
 *  `text` doesn't say. */
function borderSide(text) {
  if (text == null) return undefined;
  let width = null;
  let style = null;
  let color = null;
  for (const word of words(text)) {
    const w = word.toLowerCase();
    if (['none', 'hidden', 'solid', 'double', 'dotted', 'dashed', 'hairline', 'groove', 'ridge', 'inset', 'outset', 'dot-dash', 'dot-dot-dash'].includes(w)) style = w;
    else if (w === 'thin' || w === 'medium' || w === 'thick') width = { thin: 0.75, medium: 2.25, thick: 3.75 }[w];
    else if (points(w) != null && /\d/.test(w)) width = points(w);
    else if (cssColor(w)) color = cssColor(w);
  }
  if (style == null || style === 'none' || style === 'hidden' || width === 0) return style == null && width == null && color == null ? undefined : null;
  if (width == null) width = 2.25;
  const xl = style === 'double' ? 'double' : style === 'hairline' ? 'hair' : style === 'dotted' ? 'dotted'
    : style === 'dashed' ? (width >= 1 ? 'mediumDashed' : 'dashed') : style === 'dot-dash' ? (width >= 1 ? 'mediumDashDot' : 'dashDot')
      : style === 'dot-dot-dash' ? (width >= 1 ? 'mediumDashDotDot' : 'dashDotDot') : width <= 0.5 ? 'thin' : width < 1.5 ? 'medium' : 'thick';
  return { style: xl, color: color && color !== 'auto' ? { rgb: `FF${color}` } : null };
}

/** Excel's names for number formats in mso-number-format, as Excel reads them back (Short Date measured). */
const NAMED_FORMATS = {
  general: 'General', standard: '#,##0.00', fixed: '0.00', percent: '0.00%', scientific: '0.00E+00',
  'short date': 'm/d/yyyy', 'medium date': 'd-mmm-yy', 'short time': 'h:mm', 'medium time': 'h:mm AM/PM', 'long time': 'h:mm:ss AM/PM',
};

/** The number format code a cell's CSS or attributes give (Excel's mso-number-format, Sheets' data-sheets-numberformat), or null. */
function numberFormat(css, attrs) {
  const raw = css['mso-number-format'];
  if (raw) {
    const code = cssUnescape(raw.replace(/^(["'])([\s\S]*)\1$/, '$2'));
    return NAMED_FORMATS[code.toLowerCase()] || code;
  }
  try {
    const f = attrs['data-sheets-numberformat'] && JSON.parse(attrs['data-sheets-numberformat']);
    if (f && typeof f['2'] === 'string') return f['2'];
  } catch { /* not JSON: none */ }
  return null;
}

/** The font change of a cell's CSS (only what it says). */
function fontChange(css) {
  const out = {};
  const weight = css['font-weight'];
  if (weight) out.b = /bold/i.test(weight) || Number(weight) >= 600;
  const style = css['font-style'];
  if (style) out.i = /italic|oblique/i.test(style);
  const deco = `${css['text-decoration'] || ''} ${css['text-decoration-line'] || ''}`.toLowerCase();
  if (deco.trim()) {
    out.u = deco.includes('underline') ? (/double/i.test(css['text-underline-style'] || '') ? 'double' : 'single') : null;
    out.strike = deco.includes('line-through');
  }
  const size = css['font-size'] && points(css['font-size']);
  // (Excel rounds a size in pixels to whole points: 18px, 13.5pt, became 14, measured.)
  if (size) out.size = /px$/i.test(css['font-size'].trim()) ? Math.round(size) : Math.round(size * 2) / 2;
  const family = css['font-family'] && css['font-family'].split(',').map((x) => x.trim().replace(/^(["'])(.*)\1$/, '$2'))
    .find((x) => x && !/^(serif|sans-serif|monospace|cursive|fantasy|system-ui)$/i.test(x));
  if (family) out.name = family;
  if (css.color) {
    const c = cssColor(css.color);
    if (c) out.color = colorOf(c);
  }
  return out;
}

const H_ALIGN = { left: 'left', right: 'right', center: 'center', centre: 'center', justify: 'justify', fill: 'fill', 'center-across': 'centerContinuous', start: 'left', end: 'right', general: null };
const V_ALIGN = { top: 'top', middle: 'center', center: 'center', bottom: 'bottom', 'text-top': 'top', 'text-bottom': 'bottom', baseline: 'bottom', justify: 'justify' };

/** The cell format change (stylesedit.js) a cell's CSS gives, as Excel took each (measured): only CSS, not align or
 *  valign attributes; text wraps unless white-space says not; a cell with no vertical-align is centred, as a
 *  browser draws one. */
export function styleChange(css, attrs = {}) {
  const change = {};
  const font = fontChange(css);
  if (Object.keys(font).length) change.font = font;
  const bg = css['background-color'] || (css.background && words(css.background).find((w) => cssColor(w) && cssColor(w) !== 'auto'));
  const fill = bg && cssColor(bg);
  if (fill && fill !== 'auto') change.fill = { rgb: `FF${fill}` };
  const border = {};
  const all = borderSide(css.border);
  for (const side of ['left', 'right', 'top', 'bottom']) {
    let v = borderSide(css[`border-${side}`]);
    const longhand = css[`border-${side}-style`];
    if (longhand) v = borderSide(`${longhand} ${css[`border-${side}-width`] || ''} ${css[`border-${side}-color`] || ''}`);
    if (v === undefined) v = all;
    if (v) border[side] = v;
  }
  if (Object.keys(border).length) change.border = border;
  const align = {};
  const h = (css['text-align'] || '').toLowerCase();
  if (H_ALIGN[h]) align.horizontal = H_ALIGN[h];
  const v = (css['vertical-align'] || 'middle').toLowerCase();
  if (V_ALIGN[v] && V_ALIGN[v] !== 'bottom') align.vertical = V_ALIGN[v];
  if (!/nowrap|^pre$/i.test(css['white-space'] || '')) align.wrap = true;
  const indent = Number(css['mso-char-indent-count']);
  if (indent > 0) { align.indent = Math.round(indent); align.horizontal = align.horizontal || 'left'; }
  if (Object.keys(align).length) change.align = align;
  const code = numberFormat(css, attrs);
  if (code && code !== 'General') change.numFmt = code;
  return change;
}

/** A cell's content: Google Sheets' data-sheets-value, Excel's x:num, x:str or x:bool, else its text read as typed
 *  (`code`, the format it brings; text formatted @ kept as text). { value } or { formula }, and the format typing gives. */
function contentOf(cell, code, date1904) {
  const a = cell.attrs;
  try {
    const v = a['data-sheets-value'] && JSON.parse(a['data-sheets-value']);
    if (v && v['1'] === 3 && typeof v['3'] === 'number') return { value: v['3'] };
    if (v && v['1'] === 4) return { value: !!v['4'] };
    if (v && v['1'] === 2 && typeof v['2'] === 'string') return { value: v['2'] };
  } catch { /* not JSON: the text */ }
  if (a['x:num'] && Number.isFinite(Number(a['x:num']))) return { value: Number(a['x:num']) };
  if ('x:str' in a) return { value: cell.text || null };
  if (a['x:bool']) return { value: /^true$/i.test(a['x:bool']) };
  if (!cell.text) return { value: null };
  if (code === '@') return { value: cell.text };
  const parsed = parseInput(cell.text, { date1904 });
  if (parsed.error) return { value: cell.text };
  return parsed.formula != null ? { formula: parsed.formula, typed: formatAfterEntry('General', parsed) } : { value: parsed.value, typed: formatAfterEntry('General', parsed) };
}

/** A cell's text as rich text when its pieces aren't all in one font ("Both <b>columns</b>", as Excel keeps it,
 *  measured): [{ text, font }] each font whole (strings.js richRunsXml), from the cell's format `s`; else null. */
function runsOf(wb, cell, s) {
  const changes = (cell.pieces || []).map((p) => fontChange(p.css));
  if (changes.length < 2 || changes.every((x) => JSON.stringify(x) === JSON.stringify(changes[0]))) return null;
  const f = wb.styles.fonts[wb.styles.cellXfs[s]?.fontId ?? 0];
  const base = { b: f.b, i: f.i, u: f.u, strike: f.strike, size: f.size, name: f.name, color: f.color };
  return cell.pieces.map((p, k) => ({ text: p.text, font: { ...base, ...changes[k], color: changes[k].color || base.color } }));
}

/** The HTML's table laid out from (r0, c0): its cells by place, spans taking their room, as a browser lays it out. */
export function layoutTable(table, r0 = 0, c0 = 0) {
  const taken = new Set();
  const placed = [];
  table.rows.forEach((row, dr) => {
    let dc = 0;
    for (const cell of row) {
      while (taken.has(`${dr},${dc}`)) dc++;
      for (let y = 0; y < cell.rowspan; y++) for (let x = 0; x < cell.colspan; x++) taken.add(`${dr + y},${dc + x}`);
      placed.push({ cell, r: r0 + dr, c: c0 + dc, r1: r0 + dr + cell.rowspan - 1, c1: c0 + dc + cell.colspan - 1 });
      dc += cell.colspan;
    }
  });
  return placed;
}

/**
 * HTML from the clipboard pasted at `sel`'s top-left cell of sheet si, as Excel's Paste takes it: each cell's value,
 * formats replacing the cell's own, merged ranges for its spans. `widths` (Keep Source Column Widths) gives the
 * columns the table's widths too, which Excel's Paste doesn't (measured: the widths stayed). Steps as pasteText's,
 * with `widths` [{ c, width }]; null when the HTML holds no cells.
 */
export function pasteHtml(wb, si, sel, html, { widths = false } = {}) {
  const table = readHtmlTable(html);
  if (!table || !table.rows.some((r) => r.length)) return null;
  const placed = layoutTable(table, sel.r0, sel.c0);
  const contents = [];
  const styles = [];
  const merges = [];
  for (const { cell, r, c, r1, c1 } of placed) {
    const change = styleChange(cell.css, cell.attrs);
    const content = contentOf(cell, change.numFmt, wb.date1904);
    if (!change.numFmt && content.typed) change.numFmt = content.typed;
    const s = Object.keys(change).length ? wb.styleWith(0, change) : 0;
    const { typed: _typed, ...what } = content;
    const runs = typeof what.value === 'string' ? runsOf(wb, cell, s) : null;
    contents.push({ r, c, ...what, ...(runs ? { runs } : {}) });
    for (let y = r; y <= r1; y++) {
      for (let x = c; x <= c1; x++) {
        if (y !== r || x !== c) contents.push({ r: y, c: x, value: null });
        styles.push({ r: y, c: x, s });
      }
    }
    if (r1 > r || c1 > c) merges.push({ r0: r, c0: c, r1, c1 });
  }
  const area = { r0: sel.r0, c0: sel.c0, r1: maxOf(placed.map((p) => p.r1)), c1: maxOf(placed.map((p) => p.c1)) };
  const overlaps = (m) => m.r0 <= area.r1 && area.r0 <= m.r1 && m.c0 <= area.c1 && area.c0 <= m.c1;
  const out = { si, contents, styles, merges: [...wb.sheet(si).merges.filter((m) => !overlaps(m)), ...merges], area };
  if (widths) {
    // (A column's width in pixels to Excel's characters: px = chars x mdw + padding, docs/SIZES.md, digitwidth.js.)
    const mdw = wb.mdw ?? 7;
    out.widths = table.cols.map((col, k) => ({ c: sel.c0 + k, width: col.width ? Math.round(((col.width - colPad(mdw)) / mdw) * 256) / 256 : null })).filter((x) => x.width > 0);
  }
  return [out];
}

/** The plain text a table's HTML holds, tabs between cells and a line for each row (for Paste Values when only HTML
 *  is on the clipboard). */
export function htmlText(html) {
  const table = readHtmlTable(html);
  if (!table) return '';
  return table.rows.map((row) => row.map((c) => c.text).join('\t')).join('\n');
}
