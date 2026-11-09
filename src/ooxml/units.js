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

/** A sheet's default column width in pixels: `defaultColWidth` when given, else `baseColWidth` (8) characters plus
 *  padding, rounded up to a multiple of 8 pixels, which is how Excel's untouched columns come to be 64px (8.43) in
 *  Calibri 11 rather than 61. (The multiple of 8 is [MS-OI29500]'s note on `baseColWidth`, not yet measured.) */
export function defaultColPx({ defaultColWidth, baseColWidth } = {}, mdw = 7) {
  if (defaultColWidth != null) return colWidthToPx(defaultColWidth, mdw);
  const base = baseColWidth != null ? Number(baseColWidth) : 8;
  return Math.ceil(colWidthToPx(charsToColWidth(base, mdw), mdw) / 8) * 8;
}

/** A row height in points to whole pixels (15pt, the Calibri 11 default, is 20px). */
export const rowHeightToPx = (pt) => Math.round(ptToPx(pt));
