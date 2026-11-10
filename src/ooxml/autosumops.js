// AutoSum (and its Average, Count Numbers, Max and Min) as Excel does it, measured over COM with the ribbon's command
// in a visible Excel (docs/AUTOSUM.md). A selection holding numbers gets its totals in the blank row under it (and
// the blank column beside it), or in the row below; blank cells selected under or beside numbers sum them up to the
// first blank or text, the selection's first column deciding the rows for all; SUM subtotals above are summed
// alone. One cell, or nothing found to sum: the formula Excel would propose, for the cell to be edited and Enter
// pressed, as Excel leaves it.
import { refName } from './refs.js';
import { formatCodeOf } from './styles.js';
import { isDateFormat } from './format.js';

export const AUTOSUM_FUNCTIONS = { sum: 'SUM', average: 'AVERAGE', count: 'COUNT', max: 'MAX', min: 'MIN' };

const blank = (cell) => !cell || ((cell.f == null || cell.f === '') && cell.v == null);

/**
 * AutoSum's function `fn` ('sum', 'average', 'count', 'max', 'min') on `sel` of sheet si: a step ({ si, contents,
 * styles, area }: the formulas, the number formats they take, the selection after), or { edit: { r, c, text } } when
 * Excel leaves a formula being edited (one cell selected, or nothing found to sum).
 */
export function autoSum(wb, si, sel, fn = 'sum') {
  const sheet = wb.sheet(si);
  const name = AUTOSUM_FUNCTIONS[fn];
  const at = (r, c) => sheet.cells.get(r)?.get(c);
  // (A date or time isn't one of the numbers AutoSum looks for: over dates Excel was seen to propose =AVERAGE().)
  const isNumber = (cell) => !!cell && typeof cell.v === 'number' && cell.t !== 'e' && !isDateFormat(formatCodeOf(wb.styles, cell.s || 0));
  const isSubtotal = (cell) => /^SUM\(/i.test(cell?.f || '');
  /** The run of numbers before (r, c), up (dr) or left (dc): blanks next to the cell taken in, then numbers to the
   *  first blank, text or sheet edge. { from, to } in that direction's index, the SUM subtotals in it, or null. */
  const run = (r, c, up) => {
    const cellAt = (i) => (up ? at(i, c) : at(r, i));
    const end = (up ? r : c) - 1;
    let i = end;
    while (i >= 0 && blank(cellAt(i))) i--;
    if (i < 0 || !isNumber(cellAt(i))) return null;
    while (i >= 0 && isNumber(cellAt(i))) i--;
    const subtotals = [];
    for (let k = end; k > i; k--) if (isSubtotal(cellAt(k))) subtotals.push(k);
    return { from: i + 1, to: end, subtotals, up };
  };
  /** The formula over a run, put in column c (row r for a run along a row). */
  const over = (span, r, c) => {
    const ref = (i) => (span.up ? refName(i, c) : refName(r, i));
    const args = span.subtotals.length ? span.subtotals.map(ref).join(',') : span.from === span.to ? ref(span.from) : `${ref(span.from)}:${ref(span.to)}`;
    return `${name}(${args})`;
  };
  const one = sel.r0 === sel.r1 && sel.c0 === sel.c1;
  const filled = [];
  for (let r = sel.r0; r <= sel.r1; r++) for (let c = sel.c0; c <= sel.c1; c++) if (!blank(at(r, c))) filled.push({ r, c });

  const out = { si, contents: [], styles: [], area: { ...sel } };
  // (The totals take the number format of the numbers summed, the first column's: seen even over a column of
  // percentages, beside one of currency.)
  const formatFrom = (cell) => formatCodeOf(wb.styles, cell?.s || 0);
  const put = (r, c, formula, code) => {
    out.contents.push({ r, c, formula });
    const s = at(r, c)?.s || 0;
    if (code !== 'General' && formatCodeOf(wb.styles, s) === 'General') out.styles.push({ r, c, s: wb.styleWith(s, { numFmt: code }) });
  };

  if (one || !filled.length) {
    // Blank cells: each row of the selection summed up from its first cell if it can be, else from the left.
    for (let r = sel.r0; r <= sel.r1; r++) {
      const span = run(r, sel.c0, true) || run(r, sel.c0, false);
      if (!span) return { edit: { r: sel.r0, c: sel.c0, text: `=${name}()` } };
      if (one) return { edit: { r, c: sel.c0, text: `=${over(span, r, sel.c0)}` } };
      const code = formatFrom(span.up ? at(span.from, sel.c0) : at(r, span.from));
      for (let c = sel.c0; c <= sel.c1; c++) put(r, c, over(span, r, c), code);
    }
    return out;
  }

  // Cells with numbers: totals in the selection's blank last row and column, or in the row below (the column to the
  // right of a single row), the first one free.
  const rowBlank = (r, c0, c1) => { for (let c = c0; c <= c1; c++) if (!blank(at(r, c))) return false; return true; };
  const colBlank = (c, r0, r1) => { for (let r = r0; r <= r1; r++) if (!blank(at(r, c))) return false; return true; };
  const lastRow = sel.r1 > sel.r0 && rowBlank(sel.r1, sel.c0, sel.c1);
  const lastCol = sel.c1 > sel.c0 && colBlank(sel.c1, sel.r0, sel.r1);
  const r1 = lastRow ? sel.r1 - 1 : sel.r1;
  const c1 = lastCol ? sel.c1 - 1 : sel.c1;
  let totalRow = lastRow ? sel.r1 : null;
  let totalCol = lastCol ? sel.c1 : null;
  if (totalRow == null && totalCol == null) {
    if (sel.r0 === sel.r1) {
      totalCol = sel.c1 + 1;
      while (!colBlank(totalCol, sel.r0, r1)) totalCol++;
      out.area.c1 = sel.c1 + 1;
    } else {
      totalRow = sel.r1 + 1;
      while (!rowBlank(totalRow, sel.c0, c1)) totalRow++;
      out.area.r1 = sel.r1 + 1;
    }
  }
  const code = formatFrom(at(sel.r0, sel.c0));
  const range = (a, b) => (a === b ? a : `${a}:${b}`);
  if (totalCol != null) for (let r = sel.r0; r <= r1; r++) put(r, totalCol, `${name}(${range(refName(r, sel.c0), refName(r, c1))})`, code);
  if (totalRow != null) for (let c = sel.c0; c <= c1; c++) put(totalRow, c, `${name}(${range(refName(sel.r0, c), refName(r1, c))})`, code);
  // (The corner of both: the total row's totals.)
  if (totalRow != null && totalCol != null) put(totalRow, totalCol, `${name}(${range(refName(totalRow, sel.c0), refName(totalRow, c1))})`, code);
  return out;
}
