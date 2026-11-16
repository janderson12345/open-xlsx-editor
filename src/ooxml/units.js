import { colPad } from './digitwidth.js';

// Office Open XML's units in CSS pixels (96 per inch): drawings measure in EMUs (914,400 an inch), row heights and font
// sizes in points, column widths in characters of the default font's maximum digit width.

/** EMUs in a pixel. */
export const EMU_PER_PX = 9525;
/** Twips in a pixel. */
export const TWIPS_PER_PX = 15;
/** Points in a pixel (72 points an inch, 96 pixels an inch). */
export const PT_PER_PX = 0.75;

/** EMUs (a number or an attribute's string) to pixels, unrounded. */
export const emuToPx = (v) => Number(v) / EMU_PER_PX;
/** Pixels to whole EMUs. */
export const pxToEmu = (px) => Math.round(px * EMU_PER_PX);
/** Twips (a number or an attribute's string) to pixels, unrounded. */
export const twipsToPx = (v) => Number(v) / TWIPS_PER_PX;
/** Points (a number or an attribute's string) to pixels, unrounded. */
export const ptToPx = (v) => Number(v) / PT_PER_PX;
/** Pixels to points, unrounded. */
export const pxToPt = (px) => px * PT_PER_PX;

// Column widths (ECMA-376 Part 1, 18.3.1.13 `col@width`). A stored width is characters of the maximum digit width
// (mdw: the widest of 0-9 in the Normal style's font, 7px for Calibri 11 at 96 dpi, digitwidth.js) plus its padding
// (5px up to an mdw of 8, more for larger fonts: digitwidth.js colPad), in 1/256ths: a column shown as "8.43" is stored
// as 9.140625. The formulas are the specification's with Excel's padding, measured over COM for 41 fonts and sizes
// (Excel's stored widths and its widths in points, test/digitwidth-excel.json).

/** A stored column width (characters, padding included) to whole pixels. */
export const colWidthToPx = (width, mdw) => Math.trunc(((256 * Number(width) + Math.trunc(128 / mdw)) / 256) * mdw);
/** Whole pixels to the stored column width that gives them. */
export const pxToColWidth = (px, mdw) => Math.trunc((px / mdw) * 256) / 256;
/** Characters as the user sees them ("8.43") to the stored width (9.140625). */
export const charsToColWidth = (chars, mdw) => Math.trunc(((chars * mdw + colPad(mdw)) / mdw) * 256) / 256;
/** The stored width to the characters the user sees, to Excel's two decimals. */
export const colWidthToChars = (width, mdw) => Math.round(((colWidthToPx(width, mdw) - colPad(mdw)) / mdw) * 100) / 100;
/** A column's pixels to the characters Excel shows for it (the width box, the drag tip): unrounded. */
export const pxToChars = (px, mdw) => Math.max(0, px - colPad(mdw)) / mdw;

/** How many characters of the maximum digit width fit in a column, its padding taken off: the room Excel gives
 *  a General number (64px holds 8, 22px holds 2; fitted with numfmt.js formatGeneral on test/fixtures/general.xlsx). */
export const colChars = (px, mdw) => Math.max(0, Math.floor((px - colPad(mdw)) / mdw));

/** A sheet's default column width in pixels: `defaultColWidth` when given, else `baseColWidth` (8) characters plus
 *  padding, rounded up to a multiple of 8 pixels, which is how Excel's untouched columns come to be 64px (8.43) in
 *  Calibri 11 rather than 61 ([MS-OI29500]'s note on `baseColWidth`; measured for 41 Normal fonts, from 56px for an
 *  mdw of 6 to 144px for 16: excel-digitwidth.ps1's StandardWidth and default column). */
export function defaultColPx({ defaultColWidth, baseColWidth } = {}, mdw = 7) {
  if (defaultColWidth != null) return colWidthToPx(defaultColWidth, mdw);
  const base = baseColWidth != null ? Number(baseColWidth) : 8;
  return Math.ceil(colWidthToPx(charsToColWidth(base, mdw), mdw) / 8) * 8;
}

/** Excel prints a column wider than it shows it: printed width = screen pixels x 1.075 (fitted on Excel's PDF of
 *  test/fixtures/widths.xlsx: 30 columns of 8 to 285 screen pixels, every one within 0.15px; 64px prints as 68.8,
 *  145 as 155.9). Measured with Excel 365 exporting through the default printer's driver (a 600 dpi inkjet); the
 *  factor may belong to the printer rather than to Excel, which a second machine would show. */
export const PRINT_COL_SCALE = 1.075;
/** A column's printed width in pixels at 96 dpi, from its width on screen (unrounded). */
export const printColPx = (screenPx) => screenPx * PRINT_COL_SCALE;

/** Excel prints rows at a scale that depends on the workbook's Normal style font (measured, not derived):
 *  - Aptos Narrow (Excel 365's default theme): whole device pixels of a 600 dpi printer (6.25 to a CSS pixel) of
 *    pt x 1.3565, truncated. Excel's PDFs of test/fixtures/widths.xlsx (21 rows of 8 to 60pt, baselines within 0.12px of
 *    one offset from their row's bottom) and large.xlsx (43 steps of 15pt rows, 20.32px each = 127 device pixels).
 *  - Calibri: the same, of pt x 1.291. Excel's PDF of test/fixtures/fills.xlsx gives 15pt rows exactly from its fill
 *    boundaries (19.36px = 121 device pixels; plain pt x 1.288, fitted earlier on text baselines, drifted 0.04px a row
 *    and lost fills further down the page), and fonts.xlsx's rows of 8 to 60pt step within about 0.1px of it (10pt
 *    12.8, 18pt 23.2, 30pt 38.72, 60pt 77.44; 8pt and 11pt 0.16 and 0.12 short).
 *  Other Normal fonts are printed at their height in points until measured. Same printer caveat as the columns. */
export const PRINT_ROW_SCALE = { 'aptos narrow': { pxPerPt: 1.3565, device: true }, calibri: { pxPerPt: 1.291, device: true } };
const PRINT_DEVICE_PER_PX = 600 / 96;
/** A row's printed height in pixels at 96 dpi, from its height in points and the Normal style's typeface. */
export function printRowPx(pt, normalFamily = 'Aptos Narrow') {
  const scale = PRINT_ROW_SCALE[String(normalFamily).toLowerCase()];
  if (!scale) return ptToPx(pt);
  const px = Number(pt) * scale.pxPerPt;
  return scale.device ? Math.trunc(px * PRINT_DEVICE_PER_PX) / PRINT_DEVICE_PER_PX : px;
}

/** A row height in points to whole pixels (15pt, the Calibri 11 default, is 20px). */
export const rowHeightToPx = (pt) => Math.round(ptToPx(pt));
