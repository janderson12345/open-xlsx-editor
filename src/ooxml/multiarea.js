// What Excel's commands do with a selection of several areas (measured over COM, tools/excel-compare/
// excel-multirange.ps1, test/multirange-excel.json; docs/GOTOSPECIAL.md): most act on each area, some on them all
// at once, and some refuse with Excel's words. Areas are { r0, c0, r1, c1 }, 0-based, in the order they were selected.

/** Excel's refusal for Sort, Filter, Text to Columns, Remove Duplicates, Fill Justify, Group and Subtotal. */
export const MULTI_RANGE_MESSAGE = "This can't be done on a multiple range selection.\n\nSelect a single range and try again.";
/** Excel's refusal for inserting or deleting cells over areas that overlap. */
export const OVERLAP_MESSAGE = 'Cannot use that command on overlapping selections.';

const meets = (a, b) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;

/** Whether any two of the areas share a cell. */
export function overlapping(areas) {
  for (let i = 0; i < areas.length; i++) for (let j = i + 1; j < areas.length; j++) if (meets(areas[i], areas[j])) return true;
  return false;
}

/**
 * The lines (rows, axis 'row', or columns, 'col') that Insert or Delete Sheet Rows and Columns take over areas, as
 * [first, last] spans from the last one to the first, so that each can be done in turn where the sheet was before
 * (Excel's: the rows of every area, 2:3 and 6 inserted as blank rows above 2 and above 6; overlapping areas taken
 * together). Deleting: the lines of all the areas, joined where they touch. Inserting: the same, broken where an
 * area starts (rows 2:4 and 3:5 insert one row above 2 and three above 3, as Excel did).
 */
export function lineSpans(areas, axis, insert = false) {
  const [lo, hi] = axis === 'row' ? ['r0', 'r1'] : ['c0', 'c1'];
  const union = [];
  for (const a of [...areas].sort((x, y) => x[lo] - y[lo])) {
    const last = union.at(-1);
    if (last && a[lo] <= last[1] + 1) last[1] = Math.max(last[1], a[hi]);
    else union.push([a[lo], a[hi]]);
  }
  if (!insert) return union.reverse();
  const starts = [...new Set(areas.map((a) => a[lo]))].sort((x, y) => x - y);
  const spans = [];
  for (const [a, b] of union) {
    const cuts = starts.filter((s) => s > a && s <= b);
    let from = a;
    for (const s of cuts) { spans.push([from, s - 1]); from = s; }
    spans.push([from, b]);
  }
  return spans.reverse();
}

/** The order to insert or delete each area's cells in, shifting them along `axis` ('row': up or down, 'col': left or
 *  right), so that one done doesn't move another yet to be done: the lowest (or rightmost) first. Excel's results
 *  were the same as each area done where it was (A2:B3 and B6:C6 deleted up: column B lost rows 2, 3 and 6). */
export function shiftOrder(areas, axis) {
  const k = axis === 'row' ? 'r0' : 'c0';
  return [...areas].sort((a, b) => b[k] - a[k]);
}
