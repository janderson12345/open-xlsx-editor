// A sheet's rows and columns as the grid and print lay them out: sizes in pixels, on screen or printed.
import { Axis } from './axis.js';
import { MAX_ROWS, MAX_COLS } from '../ooxml/refs.js';
import { colWidthToPx, defaultColPx, rowHeightToPx, printColPx, printRowPx } from '../ooxml/units.js';

/** Calibri 11's maximum digit width at 96 dpi; it will come from the Normal style's font once styles are read. */
export const MDW = 7;
/** How far past the used range the grid reaches (Excel's goes on to row 1,048,576; this one grows with the content). */
export const PAD_ROWS = 100;
export const PAD_COLS = 10;

/** The grid's rows and columns for a sheet: sizes from its rows, cols and defaults, hidden ones of no size. With
 *  `print`, columns and rows at their printed sizes (units.js printColPx, printRowPx, which needs the Normal style's
 *  typeface) instead of their sizes on screen, and with `scale` (a fraction) rows as they print at that scale. With
 *  `auto` (rowheights.js), rows without customHeight take Excel's own heights instead of their stored `ht`. */
export function sheetGeometry(sheet, { print = false, normalFamily, auto = null, scale = 1 } = {}) {
  // (Columns scaled are whole printer dots too: 430 at 100% is 323 at 75%, measured at 40% to 90%.)
  const colPx = !print ? (px) => px : scale === 1 ? printColPx : (px) => Math.round(printColPx(px) * 6.25 * scale) / 6.25 / scale;
  // (Printed at a scale, a row's height is the printer's whole dots of its height at that scale, docs/PRINT.md; here
  // before the scale, which the page's zoom applies.)
  const rowPx = !print ? rowHeightToPx : scale === 1 ? (pt) => printRowPx(pt, normalFamily) : (pt) => Math.round(printRowPx(pt, normalFamily) * 6.25 * scale) / 6.25 / scale;
  const rowCount = Math.min(MAX_ROWS, Math.max(sheet.maxRow + 1 + PAD_ROWS, 200));
  const colCount = Math.min(MAX_COLS, Math.max(sheet.maxCol + 1 + PAD_COLS, 26));
  const rowSizes = new Map();
  for (const [r, info] of sheet.rows) {
    if (r >= rowCount) continue;
    if (info.hidden) rowSizes.set(r, 0);
    else if (info.ht != null && (!auto || info.customHeight)) rowSizes.set(r, rowPx(info.ht));
  }
  if (auto) for (const [r, pt] of auto) if (r < rowCount && !rowSizes.has(r)) rowSizes.set(r, rowPx(pt));
  const colSizes = new Map();
  for (const col of sheet.cols) {
    for (let c = col.min; c <= Math.min(col.max, colCount - 1); c++) {
      if (col.hidden) colSizes.set(c, 0);
      else if (col.width != null) colSizes.set(c, colPx(colWidthToPx(col.width, MDW)));
    }
  }
  // (Show Formulas draws every column twice as wide, its width as saved unchanged: Excel's 54.75 points to 109.5.)
  if (sheet.showFormulas) for (const [c, px] of colSizes) colSizes.set(c, px * 2);
  const rows = new Axis(rowCount, rowPx(sheet.format.defaultRowHeight), rowSizes);
  const cols = new Axis(colCount, colPx(defaultColPx(sheet.format, MDW)) * (sheet.showFormulas ? 2 : 1), colSizes);
  return { rows, cols, rowHeadW: Math.max(3, String(rowCount).length) * MDW + 12 };
}

/** The column headings' height and the room each row heading's digits need (view chrome, not workbook content). */
export const COLHEAD_H = 20;
/** Rows and columns drawn beyond the edges of the view, so a small scroll shows no gap before the redraw. */
export const OVERSCAN = 3;
/** How many columns back to look for text that runs on into the view from the left. */
export const SPILL_BACK = 64;
