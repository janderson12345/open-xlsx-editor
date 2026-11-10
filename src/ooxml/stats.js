// What Excel's status bar says about the selection: Average, Count and Sum (its default three), and Numerical Count,
// Minimum and Maximum. From Excel's documentation and its status bar's behaviour (COM doesn't reach the status bar):
// shown for a selection of two or more cells holding something; Count counts the cells that aren't empty, the others
// take numbers only (not text, nor TRUE and FALSE); cells in hidden rows and columns are left out; with an error in
// the selection only Count is given. Numbers are shown in the number format of the first cell with one.
import { formatCodeOf } from './styles.js';
import { cellText } from './numfmt.js';

/**
 * The status bar's figures for selection `sel` of sheet si, or null when Excel shows none (one cell, or fewer than
 * two filled): { count, numbers, sum, average, min, max, text: { sum, average, min, max } } (no sums when the
 * selection holds an error or no number).
 */
export function selectionStats(wb, si, sel) {
  if (sel.r0 === sel.r1 && sel.c0 === sel.c1) return null;
  const sheet = wb.sheet(si);
  const hiddenCol = (c) => sheet.cols.some((x) => x.hidden && c >= x.min && c <= x.max);
  let count = 0;
  let numbers = 0;
  let sum = 0;
  let min = Infinity;
  let max = -Infinity;
  let error = false;
  let first = null;
  for (const [r, line] of sheet.cells) {
    if (r < sel.r0 || r > sel.r1 || sheet.rows.get(r)?.hidden) continue;
    for (const [c, cell] of line) {
      if (c < sel.c0 || c > sel.c1 || cell.v == null || cell.v === '' || hiddenCol(c)) continue;
      count++;
      if (cell.t === 'e') error = true;
      else if (typeof cell.v === 'number') {
        numbers++;
        sum += cell.v;
        min = Math.min(min, cell.v);
        max = Math.max(max, cell.v);
        if (!first || r < first.r || (r === first.r && c < first.c)) first = { r, c, s: cell.s || 0 };
      }
    }
  }
  if (count < 2) return null;
  if (error || !numbers) return { count, numbers };
  const code = formatCodeOf(wb.styles, first.s);
  const show = (v) => cellText({ t: 'n', v }, Infinity, code, wb.date1904).trim();
  const average = sum / numbers;
  return { count, numbers, sum, average, min, max, text: { sum: show(sum), average: show(average), min: show(min), max: show(max) } };
}
