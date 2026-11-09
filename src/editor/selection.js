// The selection as the grid shows it: the selected range shaded, the active cell outlined, both in an overlay above
// the cells and below the sticky headings. View state only: nothing here reaches the workbook or its undo history.
import { setVars, div } from './grid.js';

/** A selection: the active cell (ar, ac) and the range (r0..r1, c0..c1) it extends over, 0-based. */
export const cellSelection = (r, c) => ({ ar: r, ac: c, r0: r, c0: c, r1: r, c1: c });

/** The range between an anchor cell and another, with the anchor kept active. */
export function extendSelection(sel, r, c) {
  return { ...sel, r0: Math.min(sel.ar, r), r1: Math.max(sel.ar, r), c0: Math.min(sel.ac, c), c1: Math.max(sel.ac, c) };
}

/** Draw the selection into the grid's overlay: a shaded box over a range of more than one cell, an outline round the
 *  active cell (or the merged range it is in). */
export function drawSelection(st) {
  const sel = st.sel;
  if (!st.overlay) return;
  if (!sel) return st.overlay.replaceChildren();
  const box = (r0, c0, r1, c1, cls) => {
    const el = div(cls);
    setVars(el, { '--oxe-x': st.cols.start(c0), '--oxe-y': st.rows.start(r0), '--oxe-w': st.cols.start(c1 + 1) - st.cols.start(c0), '--oxe-h': st.rows.start(r1 + 1) - st.rows.start(r0) });
    return el;
  };
  const kids = [];
  if (sel.r0 !== sel.r1 || sel.c0 !== sel.c1) kids.push(box(sel.r0, sel.c0, sel.r1, sel.c1, 'oxe-sel-range'));
  const m = st.sheet.merges.find((x) => sel.ar >= x.r0 && sel.ar <= x.r1 && sel.ac >= x.c0 && sel.ac <= x.c1);
  kids.push(m ? box(m.r0, m.c0, m.r1, m.c1, 'oxe-sel-active') : box(sel.ar, sel.ac, sel.ar, sel.ac, 'oxe-sel-active'));
  st.overlay.replaceChildren(...kids);
}
