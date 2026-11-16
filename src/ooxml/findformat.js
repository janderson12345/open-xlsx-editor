// Find Format and Replace Format (Find and Replace's Format... buttons, docs/FIND.md): a cell's format resolved to what
// it looks like, compared property by property as Excel's Find compares. Measured over COM (Application.FindFormat
// with Range.Find, tools/excel-compare/excel-findformat.ps1, test/find-format2-excel.json): a theme colour and its RGB
// match each other both ways, a tint of it doesn't; automatic ink is black; "General" finds the cells with no number
// format; a font name finds the cells drawn in it whatever their size.
import { resolveColor } from './theme.js';
import { fontFamily, formatCodeOf, xfOf } from './styles.js';

const SIDES = ['left', 'right', 'top', 'bottom'];
const hex6 = (argb) => (argb ? argb.replace('#', '').slice(-6).toUpperCase() : null);

/**
 * Cell format `s` resolved: { fontName, fontSize, bold, italic, underline ('single', 'double' ... or null), strike,
 * color (RRGGBB, automatic as black), fill ('none', RRGGBB for a solid fill, or 'pattern:fg:bg'), numFmt (the code),
 * horizontal, vertical, wrap, indent, rotation, shrink, locked, hidden, border: { left, right, top, bottom: 'none' or
 * 'style:RRGGBB' } }.
 */
export function resolvedFormat(wb, s) {
  const st = wb.styles;
  const xf = xfOf(st, s || 0);
  const theme = wb.theme;
  const font = st.fonts[xf.fontId] || st.fonts[0] || {};
  const fill = st.fills[xf.fillId];
  const color = (c, auto = '000000') => hex6(resolveColor(c, theme, st.palette)) || auto;
  let fillKey = 'none';
  if (fill && fill.pattern === 'solid') fillKey = color(fill.fg, 'FFFFFF');
  else if (fill && fill.pattern && fill.pattern !== 'none') fillKey = `${fill.pattern}:${color(fill.fg)}:${color(fill.bg, 'FFFFFF')}`;
  const b = st.borders[xf.borderId] || {};
  const side = (x) => (x && x.style ? `${x.style}:${color(x.color)}` : 'none');
  const a = xf.align || {};
  return {
    fontName: fontFamily(font, theme), fontSize: font.size, bold: !!font.b, italic: !!font.i, underline: font.u || null,
    strike: !!font.strike, color: color(font.color), fill: fillKey, numFmt: formatCodeOf(st, s || 0),
    horizontal: a.horizontal || 'general', vertical: a.vertical || 'bottom', wrap: !!a.wrap, indent: a.indent || 0,
    rotation: a.rotation || 0, shrink: !!a.shrink,
    locked: xf.protection ? xf.protection.locked !== false : true, hidden: !!xf.protection?.hidden,
    border: Object.fromEntries(SIDES.map((k) => [k, side(b[k])])),
  };
}

/** A format as the Find Format and Choose Format From Cell give it, in resolvedFormat's terms: the Format... box's
 *  { bold, italic, underline, strike, color, fill (ARGB), numFmt, border: { side: { style, color } } } taken as
 *  it sets them; a format from a cell is already resolved. */
function wanted(f) {
  if (f.from) return f;
  const out = {};
  for (const k of ['bold', 'italic', 'strike']) if (f[k] != null) out[k] = !!f[k];
  if (f.underline != null) out.underline = f.underline === true ? 'single' : f.underline || null;
  if (f.color) out.color = hex6(f.color);
  if (f.fill) out.fill = f.fill === 'none' ? 'none' : hex6(f.fill);
  if (f.numFmt) out.numFmt = f.numFmt;
  for (const k of ['fontName', 'fontSize', 'horizontal', 'vertical', 'wrap', 'indent', 'rotation', 'shrink', 'locked', 'hidden']) if (f[k] != null) out[k] = f[k];
  if (f.border) out.border = Object.fromEntries(Object.entries(f.border).map(([k, x]) => [k, x ? `${x.style}${x.color ? `:${hex6(x.color)}` : ''}` : 'none']));
  return out;
}

/**
 * Whether cell `cell`'s format has every property `f` sets, the others free (Bold alone finds the bold cells whatever
 * their font, fill or size, test/find-format-excel.json); a format chosen from a cell sets them all, so only cells
 * that look the same match. A border side asked for without a colour matches that line in any colour.
 */
export function formatMatches(wb, cell, f) {
  if (!f) return true;
  const want = wanted(f);
  const have = resolvedFormat(wb, cell?.s || 0);
  for (const [k, v] of Object.entries(want)) {
    if (k === 'from') continue;
    if (k === 'border') {
      for (const [sideName, x] of Object.entries(v)) {
        const h = have.border[sideName] || 'none';
        if (x.includes(':') || x === 'none' ? h !== x : h.split(':')[0] !== x) return false;
      }
    } else if (have[k] !== v) return false;
  }
  return true;
}

/** A Replace Format as a styleWith change: the Format... box's properties, or a cell's whole format (`from`), which
 *  the cell takes as it is (its theme colours kept as theme colours). */
export function formatChange(f) {
  if (f.from) return null;
  const font = {};
  if (f.bold != null) font.b = !!f.bold;
  if (f.italic != null) font.i = !!f.italic;
  if (f.underline != null) font.u = f.underline ? 'single' : null;
  if (f.strike != null) font.strike = !!f.strike;
  if (f.color) font.color = { rgb: f.color };
  const border = f.border ? Object.fromEntries(Object.entries(f.border).map(([k, x]) => [k, x ? { style: x.style, ...(x.color ? { color: { rgb: x.color } } : {}) } : null])) : null;
  return { ...(Object.keys(font).length ? { font } : {}), ...(f.fill ? { fill: f.fill === 'none' ? null : { rgb: f.fill } } : {}),
    ...(f.numFmt ? { numFmt: f.numFmt } : {}), ...(border ? { border } : {}) };
}

/** Cell format `s` given Replace Format `f`. */
export function withFormat(wb, s, f) {
  if (f.from) return f.from.s || 0;
  return wb.styleWith(s, formatChange(f));
}

/** The format chosen from cell (r, c) of sheet si (Choose Format From Cell): every property of how it looks. */
export function formatFromCell(wb, si, r, c) {
  const s = wb.sheet(si).cells.get(r)?.get(c)?.s || 0;
  return { from: { si, r, c, s }, ...resolvedFormat(wb, s) };
}

/** What the bar shows of a format: its properties in words ("bold, italic, 0.00, font colour, fill"), or the cell it
 *  came from. */
export function describeFormat(f, cellName = '') {
  if (!f) return 'No Format Set';
  if (f.from) return `Format of ${cellName}`;
  const words = ['bold', 'italic', 'underline', 'strike'].filter((k) => f[k]);
  if (f.numFmt) words.push(f.numFmt);
  if (f.color) words.push('font colour');
  if (f.fill) words.push(f.fill === 'none' ? 'no fill' : 'fill');
  if (f.border) words.push('border');
  return words.join(', ') || 'Format set';
}
