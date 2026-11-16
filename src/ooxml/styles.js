// The workbook's styles (xl/styles.xml): fonts, fills, borders and the cell formats (cellXfs) a cell's `s` points to.
// Read once; the editor turns each cell format into CSS.
import { kids, kid, attr, onOff } from './xml.js';
import { INDEXED } from './theme.js';
import { builtinFormat } from './format.js';

/** Excel's font when styles.xml is missing or has none: the body font of the theme, 11pt. */
const FALLBACK_FONT = { name: 'Calibri', size: 11, scheme: 'minor', b: false, i: false, u: null, strike: false, color: null, vertAlign: null };
const NO_FILL = { pattern: 'none', fg: null, bg: null };
const NO_SIDE = { style: null, color: null };
const NO_BORDER = { left: NO_SIDE, right: NO_SIDE, top: NO_SIDE, bottom: NO_SIDE, diagonal: NO_SIDE, diagonalUp: false, diagonalDown: false };
const NO_ALIGN = { horizontal: 'general', vertical: 'bottom', wrap: false, indent: 0, rotation: 0, shrink: false, readingOrder: 0 };
/** A cell's protection when its format says nothing: locked, formula shown. */
const PROTECTED = { locked: true, hidden: false };

const isOn = (v) => v === '1' || v === 'true';
/** A cell format's `alignment`: horizontal (general, left, center, right, fill, justify, centerContinuous,
 *  distributed), vertical (top, center, bottom, justify, distributed), wrapping, indent (in levels), rotation (Excel's
 *  textRotation: 1 to 90 degrees up, 91 to 180 down by 90 less, 255 the letters stacked), shrink to fit and reading
 *  order (0 by context, 1 left to right, 2 right to left). */
export function readAlignment(el) {
  if (!el) return NO_ALIGN;
  return {
    horizontal: attr(el, 'horizontal') || 'general',
    vertical: attr(el, 'vertical') || 'bottom',
    wrap: isOn(attr(el, 'wrapText')),
    indent: Number(attr(el, 'indent') || 0),
    rotation: Number(attr(el, 'textRotation') || 0),
    shrink: isOn(attr(el, 'shrinkToFit')),
    readingOrder: Number(attr(el, 'readingOrder') || 0),
  };
}

/** A cell format's `protection`: locked (on unless it says 0) and hidden (its formula not shown when the sheet is
 *  protected). */
export function readProtection(el) {
  if (!el) return PROTECTED;
  const locked = attr(el, 'locked');
  return { locked: locked == null || isOn(locked), hidden: isOn(attr(el, 'hidden')) };
}

/** A colour element as { rgb, theme, indexed, auto, tint } (theme.js resolveColor turns it into RRGGBB), or null. */
export function readColor(el) {
  if (!el) return null;
  const num = (name) => (attr(el, name) != null ? Number(attr(el, name)) : null);
  const out = { rgb: attr(el, 'rgb'), theme: num('theme'), indexed: num('indexed'), auto: attr(el, 'auto') === '1' || attr(el, 'auto') === 'true', tint: num('tint') || 0 };
  // (The tint as written, kept out of sight for a colour written back: Excel's 0.79998168889431442 has a digit more
  // than JavaScript prints.)
  if (out.tint) Object.defineProperty(out, 'tintText', { value: attr(el, 'tint'), enumerable: false });
  return out;
}

/** A `u` element's underline: 'single', 'double', 'singleAccounting' or 'doubleAccounting', null for `val="none"`
 *  (no underline, as some generators write on every font), undefined with no element. */
export function underline(u) {
  if (!u) return undefined;
  const v = attr(u, 'val') || 'single';
  return v === 'none' ? null : v;
}

/** One `font` element. */
export function readFont(el, ns) {
  const val = (name) => attr(kid(el, name, ns), 'val');
  const u = kid(el, 'u', ns);
  return {
    name: val('name') || FALLBACK_FONT.name,
    size: val('sz') != null ? Number(val('sz')) : FALLBACK_FONT.size,
    scheme: val('scheme'), // 'minor' (the theme's body font), 'major' (its heading font) or null
    b: !!onOff(kid(el, 'b', ns)),
    i: !!onOff(kid(el, 'i', ns)),
    u: underline(u) ?? null,
    strike: !!onOff(kid(el, 'strike', ns)),
    color: readColor(kid(el, 'color', ns)),
    vertAlign: val('vertAlign') || null, // 'superscript', 'subscript' or null
  };
}

/** One `fill`: its pattern (solid, a hatch, or none) with foreground and background colours, or a gradient
 *  ({ pattern: 'gradient', gradient: { type: 'linear' | 'path', degree, left, right, top, bottom, stops: [{ position,
 *  color }] } }, its first stop's colour as fg). (In a solid fill the colour is the foreground's, as Excel writes it.) */
export function readFill(el, ns) {
  const g = kid(el, 'gradientFill', ns);
  if (g) {
    const n = (name) => Number(attr(g, name) || 0);
    const stops = kids(g, 'stop', ns).map((s) => ({ position: Number(attr(s, 'position') || 0), color: readColor(kid(s, 'color', ns)) }));
    const gradient = { type: attr(g, 'type') === 'path' ? 'path' : 'linear', degree: n('degree'), left: n('left'), right: n('right'), top: n('top'), bottom: n('bottom'), stops };
    return { pattern: 'gradient', fg: stops[0]?.color || null, bg: null, gradient };
  }
  const p = kid(el, 'patternFill', ns);
  if (!p) return NO_FILL;
  return { pattern: attr(p, 'patternType') || 'none', fg: readColor(kid(p, 'fgColor', ns)), bg: readColor(kid(p, 'bgColor', ns)) };
}

/** One `dxf` (the format a conditional format applies): only what it sets. A dxf's solid fill colour is its pattern's
 *  background colour, as Excel writes it (`<patternFill><bgColor rgb="FFFF0000"/>`); a property set to off is false. */
export function readDxf(el, ns) {
  const font = kid(el, 'font', ns);
  const p = kid(kid(el, 'fill', ns), 'patternFill', ns);
  const on = (name) => { const x = kid(font, name, ns); return x ? !!onOff(x) : undefined; };
  const u = kid(font, 'u', ns);
  const numFmt = kid(el, 'numFmt', ns);
  const border = kid(el, 'border', ns);
  return {
    font: font ? { b: on('b'), i: on('i'), strike: on('strike'), u: u ? underline(u) : undefined, color: readColor(kid(font, 'color', ns)) } : null,
    fill: p ? readColor(kid(p, 'bgColor', ns)) || readColor(kid(p, 'fgColor', ns)) : null,
    // (A conditional format's number format and borders, as Excel writes them in the dxf.)
    numFmt: numFmt ? attr(numFmt, 'formatCode') : null,
    border: border ? readBorder(border, ns) : null,
  };
}

/** One `border`: each side's style (thin, medium, thick, hair, dotted, dashed, double, ...) and colour. */
export function readBorder(el, ns) {
  const side = (name) => {
    const s = kid(el, name, ns);
    const style = s && attr(s, 'style');
    return style && style !== 'none' ? { style, color: readColor(kid(s, 'color', ns)) } : NO_SIDE;
  };
  // (vertical and horizontal: the lines inside a range, which only a table style's dxfs use.)
  return { left: side('left') === NO_SIDE ? side('start') : side('left'), right: side('right') === NO_SIDE ? side('end') : side('right'), top: side('top'), bottom: side('bottom'),
    diagonal: side('diagonal'), diagonalUp: isOn(attr(el, 'diagonalUp')), diagonalDown: isOn(attr(el, 'diagonalDown')), vertical: side('vertical'), horizontal: side('horizontal') };
}

/** One `xf` of cellXfs or cellStyleXfs: its font, fill, border (indexes kept inside the lists), number format,
 *  alignment, protection and cell style (`xfId`). */
export function readXf(el, ns, { fonts, fills, borders }) {
  const id = (name, list) => Math.min(Number(attr(el, name) || 0), list.length - 1);
  return { fontId: id('fontId', fonts), fillId: id('fillId', fills), borderId: id('borderId', borders), numFmtId: Number(attr(el, 'numFmtId') || 0), align: readAlignment(kid(el, 'alignment', ns)), protection: readProtection(kid(el, 'protection', ns)), xfId: Number(attr(el, 'xfId') || 0) };
}

/** The typeface a font is drawn in: the theme's for a scheme font (Excel uses the theme's font whatever `name` says),
 *  else its own name. */
export function fontFamily(font, theme) {
  if (font.scheme === 'minor') return theme.fonts.minor;
  if (font.scheme === 'major') return theme.fonts.major;
  return font.name;
}

/** { fonts, fills, borders, cellXfs: [{ fontId, fillId, borderId, numFmtId, align }], normal, palette, numFmts } from a styles part (null:
 *  Excel's defaults). `normal` is the Normal style's font (cellStyleXfs 0), the one column widths are measured in;
 *  `palette` the indexed colours (the workbook's own `indexedColors`, else the legacy palette). */
/** A table style element's stripe size as Excel allows it, 1 to 9 (a file's 0, negative or huge one kept in range, so
 *  stripes always advance). */
const stripeSize = (v) => Math.max(1, Math.min(9, Math.trunc(Number(v)) || 1));

export function readStyles(doc) {
  if (!doc) return { fonts: [FALLBACK_FONT], fills: [NO_FILL], borders: [NO_BORDER], cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, align: NO_ALIGN, protection: PROTECTED }], normal: FALLBACK_FONT, palette: INDEXED, numFmts: {}, dxfs: [] };
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const fonts = kids(kid(root, 'fonts', ns), 'font', ns).map((el) => readFont(el, ns));
  if (!fonts.length) fonts.push(FALLBACK_FONT);
  const fills = kids(kid(root, 'fills', ns), 'fill', ns).map((el) => readFill(el, ns));
  if (!fills.length) fills.push(NO_FILL);
  const borders = kids(kid(root, 'borders', ns), 'border', ns).map((el) => readBorder(el, ns));
  if (!borders.length) borders.push(NO_BORDER);
  const xf = (el) => readXf(el, ns, { fonts, fills, borders });
  const cellXfs = kids(kid(root, 'cellXfs', ns), 'xf', ns).map(xf);
  if (!cellXfs.length) cellXfs.push({ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, align: NO_ALIGN, protection: PROTECTED });
  const styleXfs = kids(kid(root, 'cellStyleXfs', ns), 'xf', ns).map(xf);
  const custom = kids(kid(kid(root, 'colors', ns), 'indexedColors', ns), 'rgbColor', ns).map((el) => (attr(el, 'rgb') || '').slice(-6));
  const palette = custom.length ? [...custom, ...INDEXED.slice(custom.length)] : INDEXED;
  const numFmts = {};
  for (const el of kids(kid(root, 'numFmts', ns), 'numFmt', ns)) numFmts[Number(attr(el, 'numFmtId'))] = attr(el, 'formatCode') || 'General';
  const dxfs = kids(kid(root, 'dxfs', ns), 'dxf', ns).map((el) => readDxf(el, ns));
  // (The workbook's own table styles: each element's dxf and stripe size; docs/TABLES.md.)
  const ts = kid(root, 'tableStyles', ns);
  // (Keyed by the file's names: no prototype, so a style named __proto__ is just a style.)
  const tableStyles = Object.create(null);
  for (const el of kids(ts, 'tableStyle', ns)) {
    if (attr(el, 'table') !== '0') {
      tableStyles[attr(el, 'name')] = Object.fromEntries(kids(el, 'tableStyleElement', ns).map((e) => [attr(e, 'type'), { dxfId: Number(attr(e, 'dxfId') ?? -1), size: stripeSize(attr(e, 'size')) }]));
    }
  }
  return { fonts, fills, borders, cellXfs, styleXfs, normal: fonts[(styleXfs[0] || cellXfs[0]).fontId], palette, numFmts, dxfs, tableStyles, defaultTableStyle: attr(ts, 'defaultTableStyle') || null };
}

/** The cell format `s` (out-of-range indexes take format 0, as Excel does). */
export const xfOf = (styles, s) => styles.cellXfs[s] || styles.cellXfs[0];
/** The format code of cell format `s`: the workbook's own (numFmts), else a built-in one, else General. */
export const formatCodeOf = (styles, s) => {
  const id = xfOf(styles, s).numFmtId;
  return styles.numFmts[id] ?? builtinFormat(id) ?? 'General';
};
/** The font of cell format `s`. */
export const fontOfXf = (styles, s) => styles.fonts[xfOf(styles, s).fontId];
