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
// (mdw: the widest of 0-9 in the Normal style's font, 7px for Calibri 11 at 96 dpi) plus 5px of padding, in 1/256ths:
// a column shown as "8.43" is stored as 9.140625. The formulas are the specification's; they still have to be checked
// against Excel's own PDFs (tools/excel-compare) before anything is laid out on top of them (LESSONS.md section 3).

/** A stored column width (characters, padding included) to whole pixels. */
export const colWidthToPx = (width, mdw) => Math.trunc(((256 * Number(width) + Math.trunc(128 / mdw)) / 256) * mdw);
/** Whole pixels to the stored column width that gives them. */
export const pxToColWidth = (px, mdw) => Math.trunc((px / mdw) * 256) / 256;
/** Characters as the user sees them ("8.43") to the stored width (9.140625). */
export const charsToColWidth = (chars, mdw) => Math.trunc(((chars * mdw + 5) / mdw) * 256) / 256;
/** The stored width to the characters the user sees, to Excel's two decimals. */
export const colWidthToChars = (width, mdw) => Math.round(((colWidthToPx(width, mdw) - 5) / mdw) * 100) / 100;

/** How many characters of the maximum digit width fit in a column, its 5px of padding taken off: the room Excel gives
 *  a General number (64px holds 8, 22px holds 2; fitted with numfmt.js formatGeneral on test/fixtures/general.xlsx). */
export const colChars = (px, mdw) => Math.max(0, Math.floor((px - 5) / mdw));

/** A sheet's default column width in pixels: `defaultColWidth` when given, else `baseColWidth` (8) characters plus
 *  padding, rounded up to a multiple of 8 pixels, which is how Excel's untouched columns come to be 64px (8.43) in
 *  Calibri 11 rather than 61. (The multiple of 8 is [MS-OI29500]'s note on `baseColWidth`, not yet measured.) */
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

/** Excel prints a row a little taller than its height in points: whole device pixels of a 600 dpi printer (6.25 to a
 *  CSS pixel) of pt x 1.3565 px, truncated. Fitted on Excel's PDF of test/fixtures/widths.xlsx (21 rows of 8 to 60pt:
 *  text baselines all within 0.12px of one offset from their row's bottom) and of large.xlsx (43 steps of 15pt rows,
 *  20.32px each = 127 device pixels). Same printer caveat as the columns. */
export const PRINT_ROW_PX_PER_PT = 1.3565;
const PRINT_DEVICE_PER_PX = 600 / 96;
/** A row's printed height in pixels at 96 dpi, from its height in points. */
export const printRowPx = (pt) => Math.trunc(Number(pt) * PRINT_ROW_PX_PER_PT * PRINT_DEVICE_PER_PX) / PRINT_DEVICE_PER_PX;

/** A row height in points to whole pixels (15pt, the Calibri 11 default, is 20px). */
export const rowHeightToPx = (pt) => Math.round(ptToPx(pt));
