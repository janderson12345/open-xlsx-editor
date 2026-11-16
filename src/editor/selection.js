// The selection as the grid shows it: the selected range shaded, the active cell outlined, both in an overlay above
// the cells and below the sticky headings. View state only: nothing here reaches the workbook or its undo history.
import { setVars, div } from './dom.js';
import { MAX_ROWS, MAX_COLS, colName } from '../ooxml/refs.js';
import { sheetPrefix } from '../calc/tokens.js';

/** A selection: the active cell (ar, ac) and the range (r0..r1, c0..c1) it extends over, 0-based. With Ctrl, more
 *  than one range (Excel's areas): `more`, the others in the order they were added, the active range last. */
export const cellSelection = (r, c) => ({ ar: r, ac: c, r0: r, c0: c, r1: r, c1: c });

/** Ctrl+click: cell (r, c) the active cell and a new range, the ranges selected before kept. */
export function addArea(sel, r, c) {
  return { ...cellSelection(r, c), more: [...(sel.more || []), { r0: sel.r0, c0: sel.c0, r1: sel.r1, c1: sel.c1 }] };
}

/** Every range of a selection, as Excel's Areas: the others first, the active range last. */
export const areasOf = (sel) => [...(sel.more || []), { r0: sel.r0, c0: sel.c0, r1: sel.r1, c1: sel.c1 }];

/** A selection of ranges (as Go To Special gives them), the first range's top-left cell active. */
export function areasSelection(areas, active = null) {
  const last = areas[areas.length - 1];
  const at = active || { r: areas[0].r0, c: areas[0].c0 };
  const sel = { ar: at.r, ac: at.c, r0: last.r0, c0: last.c0, r1: last.r1, c1: last.c1 };
  if (areas.length > 1) sel.more = areas.slice(0, -1).map((a) => ({ r0: a.r0, c0: a.c0, r1: a.r1, c1: a.c1 }));
  return sel;
}

/** The areas Find and Replace look in ({ si, areas }): the selection's, when it is more than one cell (a merged
 *  range alone counts as one cell, as in Excel); null otherwise, for the whole sheet. */
export function withinSelection(sheet, sel, si) {
  const one = (a) => (a.r0 === a.r1 && a.c0 === a.c1) || sheet.merges.some((m) => m.r0 === a.r0 && m.c0 === a.c0 && m.r1 === a.r1 && m.c1 === a.c1);
  return sel.more?.length || !one(sel) ? { si, areas: areasOf(sel).map(({ r0, c0, r1, c1 }) => ({ r0, c0, r1, c1 })) } : null;
}

/** What a name for selection `sel` of sheet si refers to, as Excel's New Name box and Name Box write it: each range
 *  absolute, with its sheet, several apart by commas (=Sheet1!$A$2:$B$3,Sheet1!$D$6:$E$7). */
export function selectionRefers(wb, si, sel) {
  const abs = (r, c) => `$${colName(c)}$${r + 1}`;
  const one = (a) => `${sheetPrefix(wb.sheets[si].name)}${abs(a.r0, a.c0)}${a.r0 === a.r1 && a.c0 === a.c1 ? '' : `:${abs(a.r1, a.c1)}`}`;
  return `=${areasOf(sel).map(one).join(',')}`;
}

/** Whether cell (r, c) is in any range of the selection. */
export const inSelection = (sel, r, c) => areasOf(sel).some((a) => r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1);

/** The range between an anchor cell and another, with the anchor kept active. */
export function extendSelection(sel, r, c) {
  return { ...sel, r0: Math.min(sel.ar, r), r1: Math.max(sel.ar, r), c0: Math.min(sel.ac, c), c1: Math.max(sel.ac, c) };
}

/** Draw the selection into the grid's overlay: a shaded box over a range of more than one cell, an outline round the
 *  active cell (or the merged range it is in), the fill handle at the range's bottom-right corner (when it can be
 *  used), and while the handle is dragged, the cells it will fill outlined. */
export function drawSelection(st) {
  const sel = st.sel;
  if (!st.selLayer) return;
  if (!sel) return st.selLayer.replaceChildren();
  const box = (r0, c0, r1, c1, cls) => rangeBox(st, r0, c0, r1, c1, cls);
  const kids = [];
  for (const a of sel.more || []) kids.push(box(a.r0, a.c0, a.r1, a.c1, 'oxe-sel-range'));
  if (sel.more?.length || sel.r0 !== sel.r1 || sel.c0 !== sel.c1) kids.push(box(sel.r0, sel.c0, sel.r1, sel.c1, 'oxe-sel-range'));
  const m = st.sheet.merges.find((x) => sel.ar >= x.r0 && sel.ar <= x.r1 && sel.ac >= x.c0 && sel.ac <= x.c1);
  kids.push(m ? box(m.r0, m.c0, m.r1, m.c1, 'oxe-sel-active') : box(sel.ar, sel.ac, sel.ar, sel.ac, 'oxe-sel-active'));
  const corner = handleCorner(st);
  if (corner) {
    const el = div('oxe-fill-handle');
    setVars(el, { '--oxe-x': st.cols.start(corner.c + 1), '--oxe-y': st.rows.start(corner.r + 1) });
    kids.push(el);
  }
  if (st.filling) kids.push(box(st.filling.r0, st.filling.c0, st.filling.r1, st.filling.c1, 'oxe-fill-drag'));
  st.selLayer.replaceChildren(...kids);
}

/** The cell whose bottom-right corner has the fill handle: the selection's (the merged range's, when the selection is
 *  the active cell in one), or null when it has none (turned off, or whole rows or columns, where Excel's handle sits
 *  off the sheet's end). */
export function handleCorner(st) {
  const sel = st.sel;
  if (!sel || !st.fillHandle || sel.more?.length) return null; // (none on more than one range, as in Excel)
  if (sel.r1 >= MAX_ROWS - 1 || sel.c1 >= MAX_COLS - 1) return null;
  const m = st.sheet.merges.find((x) => x.r0 === sel.r0 && x.c0 === sel.c0 && x.r1 >= sel.r1 && x.c1 >= sel.c1);
  return m ? { r: m.r1, c: m.c1 } : { r: sel.r1, c: sel.c1 };
}

/** A box over cells r0..r1, c0..c1 in the overlay's coordinates. */
export function rangeBox(st, r0, c0, r1, c1, cls) {
  const el = div(cls);
  setVars(el, { '--oxe-x': st.cols.start(c0), '--oxe-y': st.rows.start(r0), '--oxe-w': st.cols.start(c1 + 1) - st.cols.start(c0), '--oxe-h': st.rows.start(r1 + 1) - st.rows.start(r0) });
  return el;
}

/** The ranges a formula being typed refers to, as Excel's range finder shows them: each outlined in its colour, the
 *  one being pointed at dashed (marks: [{ r0, c0, r1, c1, color, dashed }]). */
export function drawMarks(st) {
  if (!st.markLayer) return;
  st.markLayer.replaceChildren(...(st.marks || []).map((m) => {
    const el = rangeBox(st, m.r0, m.c0, m.r1, m.c1, m.dashed ? 'oxe-mark oxe-mark-point' : 'oxe-mark');
    el.style.setProperty('--oxe-mark', m.color);
    return el;
  }));
}
