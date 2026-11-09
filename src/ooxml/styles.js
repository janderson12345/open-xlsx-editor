// The workbook's styles (xl/styles.xml): fonts and the cell formats (cellXfs) a cell's `s` points to. Read once; the
// editor turns each cell format into a CSS class. Fills, borders, alignment and number formats come next.
import { kids, kid, attr, onOff } from './xml.js';

/** Excel's font when styles.xml is missing or has none: the body font of the theme, 11pt. */
const FALLBACK_FONT = { name: 'Calibri', size: 11, scheme: 'minor', b: false, i: false, u: null, strike: false };

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
  };
}

/** The typeface a font is drawn in: the theme's for a scheme font (Excel uses the theme's font whatever `name` says),
 *  else its own name. */
export function fontFamily(font, theme) {
  if (font.scheme === 'minor') return theme.fonts.minor;
  if (font.scheme === 'major') return theme.fonts.major;
  return font.name;
}

/** { fonts, cellXfs: [{ fontId }], normal } from a styles part (null: Excel's defaults). `normal` is the Normal
 *  style's font (cellStyleXfs 0), the one column widths are measured in. */
export function readStyles(doc) {
  if (!doc) return { fonts: [FALLBACK_FONT], cellXfs: [{ fontId: 0 }], normal: FALLBACK_FONT };
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const fonts = kids(kid(root, 'fonts', ns), 'font', ns).map((el) => readFont(el, ns));
  if (!fonts.length) fonts.push(FALLBACK_FONT);
  const xf = (el) => ({ fontId: Math.min(Number(attr(el, 'fontId') || 0), fonts.length - 1) });
  const cellXfs = kids(kid(root, 'cellXfs', ns), 'xf', ns).map(xf);
  if (!cellXfs.length) cellXfs.push({ fontId: 0 });
  const styleXfs = kids(kid(root, 'cellStyleXfs', ns), 'xf', ns).map(xf);
  return { fonts, cellXfs, normal: fonts[(styleXfs[0] || cellXfs[0]).fontId] };
}

/** The font of cell format `s` (out-of-range indexes take format 0, as Excel does). */
export const fontOfXf = (styles, s) => styles.fonts[(styles.cellXfs[s] || styles.cellXfs[0]).fontId];
