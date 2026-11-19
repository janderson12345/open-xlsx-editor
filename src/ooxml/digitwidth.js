// The maximum digit width (mdw): the widest of 0-9 in the Normal style's font, in whole screen pixels at 96 dpi,
// which column widths are counted in (ECMA-376 Part 1, 18.3.1.13). Measured from Excel 365 over COM
// (tools/excel-compare/excel-digitwidth.ps1, test/digitwidth-excel.json): the Normal style given each font and size,
// columns set to 1, 2, 5, 10 and 20 characters, their widths read back in points. Excel gets its widths from Windows'
// hinted glyphs, which a browser can't read; the font's own digits at the size in pixels, rounded, give the same for
// 32 of the 36 installed fonts measured (Calibri 8 and 14, Arial 14 and Times New Roman 10 are a pixel off), so the
// measured table comes first and the browser's measure (setDigitMeasure) for any other font.
import { excelFontMetric } from './excelfonts.js';

/** Excel's maximum digit widths by typeface (lower case) and size in points (measured). */
export const MEASURED_MDW = {
  calibri: { 8: 6, 9: 6, 10: 7, 11: 7, 12: 8, 14: 10, 16: 11, 18: 12, 20: 14, 24: 16 },
  'aptos narrow': { 11: 7 },
  aptos: { 10: 7, 11: 8, 12: 9 },
  arial: { 8: 6, 9: 7, 10: 7, 11: 8, 12: 9, 14: 11, 16: 12 },
  'times new roman': { 10: 6, 11: 7, 12: 8, 14: 9 },
  cambria: { 10: 7, 11: 8, 12: 9 },
  'segoe ui': { 9: 6, 10: 7, 11: 8 },
  verdana: { 8: 7, 10: 8, 12: 10 },
  'courier new': { 10: 8, 11: 9, 12: 10 },
  tahoma: { 10: 7 },
  georgia: { 11: 9 },
  consolas: { 11: 8 },
  'ms sans serif': { 10: 7 },
};

/** The digits' width as a fraction of the font size, for typefaces measured, when there is no browser to ask
 *  (Node, tests): the font's widest digit advance over its em (from the font files). Others take Calibri's. */
const DIGIT_EM = {
  calibri: 0.5068, arial: 0.5562, 'times new roman': 0.5, cambria: 0.5537, 'segoe ui': 0.5391, verdana: 0.6358,
  'courier new': 0.6001, tahoma: 0.5459, georgia: 0.6138, consolas: 0.5498,
};

let measure = null;
/** Give the measure of a typeface's widest digit at a size in CSS pixels (fractional): (family, px) => width. The
 *  editor gives a canvas one (src/editor/digitmeasure.js); null takes it away. */
export function setDigitMeasure(fn) {
  measure = fn;
  cache.clear();
}

const cache = new Map();
/** The maximum digit width in whole pixels of `family` at `size` points: Excel's measured one, else the browser's
 *  measure of 0-9 at the size in pixels rounded as Excel rounds (to the nearest pixel), else an estimate. At least 1. */
export function digitWidth(family, size = 11) {
  const name = String(family || 'Calibri').toLowerCase();
  const pt = Number(size) || 11;
  // (Excel's table of the fonts workbooks use at every half point, excelfonts.js, then the older one, then derived.)
  return excelFontMetric(name, pt, 'mdw') || MEASURED_MDW[name]?.[pt] || deriveDigitWidth(family, pt);
}

/** The maximum digit width worked out from the font alone (not Excel's table): the browser's measure of 0-9 at the
 *  size in pixels, rounded to the nearest pixel, else the font's digit advance (DIGIT_EM; Calibri's for others). */
export function deriveDigitWidth(family, size = 11) {
  const name = String(family || 'Calibri').toLowerCase();
  const pt = Number(size) || 11;
  const key = `${name}|${pt}`;
  if (!cache.has(key)) {
    const px = (pt * 96) / 72;
    let w = null;
    try { w = measure ? measure(family, px) : null; } catch { w = null; }
    if (!(w > 0)) w = px * (DIGIT_EM[name] ?? DIGIT_EM.calibri);
    cache.set(key, Math.max(1, Math.round(w)));
  }
  return cache.get(key);
}

/** The padding Excel puts in a column, in pixels, besides its characters: 5 up to an mdw of 8, 7 to 12, 9 to 16
 *  (measured: a column of n characters is n x mdw + this; the default 8.43 of Calibri 11 is 8 x 7 + 5 = 61, rounded
 *  up to 64). Two pixels each side and one for the gridline at Calibri 11's size, as ECMA-376 says, growing a pixel
 *  each side for every 4 pixels of digit. */
export const colPad = (mdw) => 2 * Math.ceil(mdw / 4) + 1;
