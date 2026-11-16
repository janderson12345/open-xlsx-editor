// The workbook's theme (xl/theme/theme1.xml): its heading and body fonts and its colour scheme. The one place theme
// fonts and colours are looked up.
import { kids, kid, NS } from './xml.js';

/** The theme a workbook without a theme part gets: Excel 365's current default, whose body font is Aptos Narrow
 *  (Excel's PDF of test/fixtures/basic.xlsx, a workbook with no theme part and fonts of scheme "minor", is set in
 *  Aptos Narrow). Older Excel versions used Calibri here; the heading font is Excel's, not yet measured. */
export const DEFAULT_THEME = {
  fonts: { major: 'Aptos Display', minor: 'Aptos Narrow' },
  // (The 2023 Office theme's colours, in scheme order; not yet measured on Excel's output.)
  colors: ['000000', 'FFFFFF', '0E2841', 'E8E8E8', '156082', 'E97132', '196B24', '0F9ED5', 'A02B93', '4EA72E', '467886', '96607D'],
};

/** The colour scheme in the theme part's order. */
const SCHEME = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

/** The legacy palette an `indexed` colour counts in (ECMA-376 Part 1, 18.8.27), unless styles.xml gives its own;
 *  64 is the system foreground (black), 65 the system background (white). */
export const INDEXED = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF',
  'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080',
  'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF',
  'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC',
  'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966',
  '003300', '333300', '993300', '993366', '333399', '333333', '000000', 'FFFFFF'];

// Tints work in Windows' integer HLS, hue, luminance and saturation from 0 to 240 (Microsoft KB Q29240, "Converting
// colors between RGB and HLS"): reproduces Excel's printed tints exactly (test/fixtures/fills.xlsx: accent 1 darker
// 50% and 25%, lighter 40%, 60% and 80%, black lighter 50%, white darker 15%, accent 6 darker 25%), where floating-point
// HLS misses 5 of the 8 by a step.
const HLS = 240;
const idiv = (a, b) => Math.trunc(a / b);
function rgbToHls(r, g, b) {
  const mx = Math.max(r, g, b);
  const mn = Math.min(r, g, b);
  const l = idiv((mx + mn) * HLS + 255, 510);
  if (mx === mn) return [idiv(HLS * 2, 3), l, 0];
  const d = mx - mn;
  const sum = mx + mn;
  const s = l <= HLS / 2 ? idiv(d * HLS + idiv(sum, 2), sum) : idiv(d * HLS + idiv(510 - sum, 2), 510 - sum);
  const norm = (v) => idiv((mx - v) * (HLS / 6) + idiv(d, 2), d);
  let h = r === mx ? norm(b) - norm(g) : g === mx ? HLS / 3 + norm(r) - norm(b) : (HLS * 2) / 3 + norm(g) - norm(r);
  if (h < 0) h += HLS;
  else if (h > HLS) h -= HLS;
  return [h, l, s];
}
function hueToChannel(h, m1, m2) {
  if (h > HLS) h -= HLS;
  else if (h < 0) h += HLS;
  if (h > 160) return m1;
  if (h > 120) h = 160 - h;
  else if (h > 40) return m2;
  return idiv(h * (m2 - m1) + 20, 40) + m1;
}
function hlsToRgb(h, l, s) {
  if (!s) return [0, 0, 0].map(() => idiv(l * 255 + 120, HLS));
  const m2 = l > 120 ? s + l - idiv(s * l + 120, HLS) : idiv((s + HLS) * l + 120, HLS);
  const m1 = l * 2 - m2;
  return [h + 80, h, h - 80].map((x) => idiv(hueToChannel(x, m1, m2) * 255 + 120, HLS));
}

/** An RRGGBB colour with a tint applied: luminance darkened (tint < 0) or lightened towards white (tint > 0). */
export function applyTint(rgb, tint) {
  if (!tint) return rgb;
  // (A file's tint is Excel's to four places: 0.79998168889431442 is its lighter 80%, -0.249977111117893 darker 25%.
  // Rounded so, 220 of 220 colours as Excel gives them over COM, test/tints-excel.json, and every cell of the 60 table
  // styles as its DisplayFormat shows them; taken as written, 0.79998 lands a step darker: TableStyleLight7's stripe D9F1CF for Excel's DAF2D0.)
  tint = Math.round(tint * 10000) / 10000;
  const [h, l, s] = rgbToHls(...[0, 2, 4].map((i) => parseInt(rgb.slice(i, i + 2), 16)));
  const l2 = Math.trunc(tint < 0 ? l * (1 + tint) : l * (1 - tint) + HLS * tint);
  return hlsToRgb(h, l2, s).map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
}

/** A SpreadsheetML colour ({ rgb, theme, indexed, auto, tint }, styles.js readColor) as '#RRGGBB', or null for
 *  automatic or unknown (the caller's default). Theme indexes swap the first two pairs: 0 is lt1, 1 dk1, 2 lt2, 3 dk2,
 *  then the accents (Excel's PDF of fills.xlsx: theme 0 white, 1 black, 2 E7E6E6, 3 1F3864 with our theme). */
export function resolveColor(color, theme, palette = INDEXED) {
  if (!color || color.auto) return null;
  let rgb = null;
  if (color.rgb) rgb = color.rgb.length === 8 ? color.rgb.slice(2) : color.rgb;
  else if (color.theme != null) rgb = theme.colors[color.theme < 4 ? color.theme ^ 1 : color.theme] || null;
  else if (color.indexed != null) rgb = palette[color.indexed] || null;
  return rgb && /^[0-9A-Fa-f]{6}$/.test(rgb) ? '#' + applyTint(rgb.toUpperCase(), color.tint || 0) : null;
}

/** { fonts: { major, minor }, colors: ['RRGGBB', ...] } from a theme part (null: the default theme). */
export function readTheme(doc) {
  if (!doc) return DEFAULT_THEME;
  const a = NS.a;
  const el = kid(doc.documentElement, 'themeElements', a);
  const fontScheme = kid(el, 'fontScheme', a);
  const latin = (name) => kid(kid(fontScheme, name, a), 'latin', a)?.getAttribute('typeface') || null;
  const clr = kid(el, 'clrScheme', a);
  const colors = SCHEME.map((name) => {
    const c = kid(clr, name, a);
    const srgb = kid(c, 'srgbClr', a);
    const sys = kid(c, 'sysClr', a);
    return srgb ? srgb.getAttribute('val') : sys ? sys.getAttribute('lastClr') : null;
  });
  // (The line styles' widths in EMU, which shapes' lnRef idx 1, 2, 3 name: drawpaint.js.)
  const lines = kids(kid(kid(el, 'fmtScheme', a), 'lnStyleLst', a), 'ln', a).map((l) => Number(l.getAttribute('w') || 0));
  return {
    ...(lines.length ? { lines } : {}),
    fonts: { major: latin('majorFont') || DEFAULT_THEME.fonts.major, minor: latin('minorFont') || DEFAULT_THEME.fonts.minor },
    colors: kids(clr).length ? colors : DEFAULT_THEME.colors,
  };
}

/** A DrawingML colour as charts write it ({ rgb } or { scheme: 'accent1', 'tx1', ... }, with lumMod and lumOff, the
 *  fractions of luminance kept and added) as '#RRGGBB', or null. Luminance in Windows' integer HLS, as tints are:
 *  black at 65% and 35% is #595959, at 15% and 85% #D9D9D9, Excel's chart text and line greys in its PDFs. */
export function resolveDrawingColor(c, theme) {
  if (!c || c.none) return null;
  const at = c.scheme != null ? SCHEME.indexOf({ tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' }[c.scheme] || c.scheme) : -1;
  const rgb = c.rgb || (at >= 0 ? theme.colors[at] : null);
  if (!rgb || !/^[0-9A-Fa-f]{6}$/.test(rgb)) return null;
  if (c.lumMod == null && c.lumOff == null) return `#${rgb.toUpperCase()}`;
  const [h, l, s] = rgbToHls(...[0, 2, 4].map((i) => parseInt(rgb.slice(i, i + 2), 16)));
  const l2 = Math.max(0, Math.min(HLS, Math.round(l * (c.lumMod ?? 1) + HLS * (c.lumOff ?? 0))));
  return `#${hlsToRgb(h, l2, s).map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase()}`;
}
