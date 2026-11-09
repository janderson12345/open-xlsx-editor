// The workbook's styles (xl/styles.xml): fonts, fills, borders and the cell formats (cellXfs) a cell's `s` points to.
// Read once; the editor turns each cell format into CSS. Alignment and number formats come next.
import { kids, kid, attr, onOff } from './xml.js';
import { INDEXED } from './theme.js';

/** Excel's font when styles.xml is missing or has none: the body font of the theme, 11pt. */
const FALLBACK_FONT = { name: 'Calibri', size: 11, scheme: 'minor', b: false, i: false, u: null, strike: false, color: null };
const NO_FILL = { pattern: 'none', fg: null, bg: null };
const NO_SIDE = { style: null, color: null };
const NO_BORDER = { left: NO_SIDE, right: NO_SIDE, top: NO_SIDE, bottom: NO_SIDE };

/** A colour element as { rgb, theme, indexed, auto, tint } (theme.js resolveColor turns it into RRGGBB), or null. */
export function readColor(el) {
  if (!el) return null;
  const num = (name) => (attr(el, name) != null ? Number(attr(el, name)) : null);
  return { rgb: attr(el, 'rgb'), theme: num('theme'), indexed: num('indexed'), auto: attr(el, 'auto') === '1' || attr(el, 'auto') === 'true', tint: num('tint') || 0 };
}

/** One `font` element. */
function readFont(el, ns) {
  const val = (name) => attr(kid(el, name, ns), 'val');
  const u = kid(el, 'u', ns);
  return {
    name: val('name') || FALLBACK_FONT.name,
    size: val('sz') != null ? Number(val('sz')) : FALLBACK_FONT.size,
    scheme: val('scheme'), // 'minor' (the theme's body font), 'major' (its heading font) or null
    b: !!onOff(kid(el, 'b', ns)),
    i: !!onOff(kid(el, 'i', ns)),
    u: u ? attr(u, 'val') || 'single' : null,
    strike: !!onOff(kid(el, 'strike', ns)),
    color: readColor(kid(el, 'color', ns)),
  };
}

/** One `fill`: its pattern (solid, a hatch, or none) with foreground and background colours. Gradient fills are read
 *  as none for now. (In a solid fill the colour is the foreground's, as Excel writes it.) */
function readFill(el, ns) {
  const p = kid(el, 'patternFill', ns);
  if (!p) return NO_FILL;
  return { pattern: attr(p, 'patternType') || 'none', fg: readColor(kid(p, 'fgColor', ns)), bg: readColor(kid(p, 'bgColor', ns)) };
}

/** One `border`: each side's style (thin, medium, thick, hair, dotted, dashed, double, ...) and colour. */
function readBorder(el, ns) {
  const side = (name) => {
    const s = kid(el, name, ns);
    const style = s && attr(s, 'style');
    return style && style !== 'none' ? { style, color: readColor(kid(s, 'color', ns)) } : NO_SIDE;
  };
  return { left: side('left') === NO_SIDE ? side('start') : side('left'), right: side('right') === NO_SIDE ? side('end') : side('right'), top: side('top'), bottom: side('bottom') };
}

/** The typeface a font is drawn in: the theme's for a scheme font (Excel uses the theme's font whatever `name` says),
 *  else its own name. */
export function fontFamily(font, theme) {
  if (font.scheme === 'minor') return theme.fonts.minor;
  if (font.scheme === 'major') return theme.fonts.major;
  return font.name;
}

/** { fonts, fills, borders, cellXfs: [{ fontId, fillId, borderId }], normal, palette } from a styles part (null:
 *  Excel's defaults). `normal` is the Normal style's font (cellStyleXfs 0), the one column widths are measured in;
 *  `palette` the indexed colours (the workbook's own `indexedColors`, else the legacy palette). */
export function readStyles(doc) {
  if (!doc) return { fonts: [FALLBACK_FONT], fills: [NO_FILL], borders: [NO_BORDER], cellXfs: [{ fontId: 0, fillId: 0, borderId: 0 }], normal: FALLBACK_FONT, palette: INDEXED };
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const fonts = kids(kid(root, 'fonts', ns), 'font', ns).map((el) => readFont(el, ns));
  if (!fonts.length) fonts.push(FALLBACK_FONT);
  const fills = kids(kid(root, 'fills', ns), 'fill', ns).map((el) => readFill(el, ns));
  if (!fills.length) fills.push(NO_FILL);
  const borders = kids(kid(root, 'borders', ns), 'border', ns).map((el) => readBorder(el, ns));
  if (!borders.length) borders.push(NO_BORDER);
  const id = (el, name, list) => Math.min(Number(attr(el, name) || 0), list.length - 1);
  const xf = (el) => ({ fontId: id(el, 'fontId', fonts), fillId: id(el, 'fillId', fills), borderId: id(el, 'borderId', borders) });
  const cellXfs = kids(kid(root, 'cellXfs', ns), 'xf', ns).map(xf);
  if (!cellXfs.length) cellXfs.push({ fontId: 0, fillId: 0, borderId: 0 });
  const styleXfs = kids(kid(root, 'cellStyleXfs', ns), 'xf', ns).map(xf);
  const custom = kids(kid(kid(root, 'colors', ns), 'indexedColors', ns), 'rgbColor', ns).map((el) => (attr(el, 'rgb') || '').slice(-6));
  const palette = custom.length ? [...custom, ...INDEXED.slice(custom.length)] : INDEXED;
  return { fonts, fills, borders, cellXfs, normal: fonts[(styleXfs[0] || cellXfs[0]).fontId], palette };
}

/** The cell format `s` (out-of-range indexes take format 0, as Excel does). */
export const xfOf = (styles, s) => styles.cellXfs[s] || styles.cellXfs[0];
/** The font of cell format `s`. */
export const fontOfXf = (styles, s) => styles.fonts[xfOf(styles, s).fontId];
