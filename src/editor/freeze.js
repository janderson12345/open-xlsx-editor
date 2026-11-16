// What View > Freeze Panes freezes, as Excel's ribbon commands do (measured over COM: CommandBars.ExecuteMso
// FreezePanes, FreezeTopRow, FreezeFirstColumn on a scrolled, split or partly hidden window;
// tools/excel-compare/excel-freeze.ps1, test/freeze-excel.json, docs/VIEW.md). Everything counts from the cell first in
// view, not from A1: Freeze Panes freezes the rows and columns in view above and left of the active cell, Freeze Top
// Row the row at the top of the window, Freeze First Column the column at its left.

/** The first row (or column) from `i` on that isn't hidden. */
const shownFrom = (axis, i) => { let k = i; while (axis.size(k) === 0 && k < i + 16384) k++; return k; };

/**
 * The freeze for `kind` ('panes', 'topRow', 'firstColumn', 'none'): { rows, cols, top, left, split }, for
 * workbook.setView. `env` = { sheet (its freeze, viewTopLeft and split, sheet.js readView), frozen (whether it is
 * frozen now), grid ({ view() with topLeft, cellsInView() { rows, cols }, geometry() with rows, cols, rowHeadW,
 * colHeadH }), active ({ r, c }) }.
 */
export function freezeFor(kind, { sheet, frozen, grid, active }) {
  const off = { rows: 0, cols: 0 };
  if (kind === 'none' || (kind === 'panes' && frozen)) return off;
  const g = grid.geometry();
  // (Frozen already, the window starts where the freeze does; Top Row then freezes that row: Excel's C5 frozen, then
  // Top Row, gave row 1.)
  const at = frozen ? (sheet.viewTopLeft || { r: 0, c: 0 }) : grid.view().topLeft;
  const vt = { r: shownFrom(g.rows, at.r), c: shownFrom(g.cols, at.c) };
  if (kind === 'topRow') return { rows: 1, cols: 0, top: vt.r, left: vt.c };
  if (kind === 'firstColumn') return { rows: 0, cols: 1, top: vt.r, left: vt.c };
  // (A split window: frozen where it is split, from the split's own first cell: Excel's split after 3 rows and 2
  // columns of E7, ySplit 1200 and xSplit 2310 twentieths of a point with the headings, froze rows 7 to 9 and columns
  // E and F, as frozenSplit.)
  if (sheet.split && (sheet.split.x || sheet.split.y)) {
    const from = sheet.viewTopLeft || { r: 0, c: 0 };
    const count = (axis, start, twips, headPx) => {
      let room = twips / 20 - headPx * 0.75;
      let k = 0;
      while (room > 0 && k < 1000) { const pt = axis.size(start + k) * 0.75; if (room < pt / 2) break; room -= pt; k++; }
      return k;
    };
    const rows = count(g.rows, from.r, sheet.split.y, g.colHeadH);
    const cols = count(g.cols, from.c, sheet.split.x, g.rowHeadW);
    if (rows || cols) return { rows, cols, top: from.r, left: from.c, split: true };
  }
  const n = grid.cellsInView();
  const inView = active.r >= vt.r && active.c >= vt.c && active.r < vt.r + n.rows && active.c < vt.c + n.cols;
  // (The active cell first in view, or out of view: the middle of the window, as Excel freezes it with A1 active,
  // 32 rows and 26 columns of its 64 by 53.)
  if (!inView || (active.r === vt.r && active.c === vt.c)) return { rows: Math.floor(n.rows / 2), cols: Math.floor(n.cols / 2), top: vt.r, left: vt.c };
  // (A hidden active row or column: the freeze goes on to the next one shown, as Excel's did with C5 in hidden rows 4
  // to 6: rows 1 to 6.)
  const r = active.r > vt.r ? shownFrom(g.rows, active.r) : active.r;
  const c = active.c > vt.c ? shownFrom(g.cols, active.c) : active.c;
  return { rows: r - vt.r, cols: c - vt.c, top: vt.r, left: vt.c };
}
