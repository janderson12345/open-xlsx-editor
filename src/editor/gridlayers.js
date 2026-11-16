// The grid's layers drawn for the part of the sheet in view (grid.js draws them in turn): the column headings, the
// frozen panes, the gridlines and the rows of cells.
import { setVars, div } from './dom.js';
import { cellElement, mergeAt, fitNumbers } from './cells.js';
import { decorations } from './decor.js';
import { drawSelection } from './selection.js';
import { drawFilterButtons } from './filterbuttons.js';
import { drawNoteMarks } from './notemarks.js';
import { colName } from '../ooxml/refs.js';
import { SPILL_BACK, MDW } from './geometry.js';

export function drawColumnHeadings(st, win) {
  const heading = (c, role) => {
    const el = div('oxe-colhead', role);
    if (role) el.setAttribute('aria-colindex', String(c + 2));
    el.textContent = colName(c);
    setVars(el, { '--oxe-x': st.cols.start(c), '--oxe-w': st.cols.size(c) });
    return el;
  };
  const kids = [st.corner];
  // (Frozen columns' headings stay where they are, in a pinned strip of their own.)
  const fc = st.freeze?.cols || 0;
  if (fc && st.pinHeads) {
    const inner = st.pinHeads.firstChild;
    const left = st.freeze.left || 0;
    inner.replaceChildren(...Array.from({ length: fc }, (_, c) => left + c).filter((c) => st.cols.size(c)).map((c) => heading(c, null)));
    kids.push(st.pinHeads);
  }
  for (let c = win.c0; c <= win.c1; c++) if (st.cols.size(c)) kids.push(heading(c, 'columnheader'));
  st.colheads.replaceChildren(...kids);
}

/** The frozen panes (rows above and columns left of the freeze, and the corner where they meet), each drawn with the
 *  layers the sheet has (gridlines, fills, cells, borders, selection) over the part of the sheet they keep in place. */
export function drawPanes(st, win) {
  const fr = st.freeze?.rows || 0;
  const fc = st.freeze?.cols || 0;
  const paint = (pane, w) => {
    if (!w) { for (const k of ['lines', 'fills', 'body', 'edges', 'selLayer', 'filterLayer', 'noteLayer']) pane[k].replaceChildren(); return; }
    // (The pane's own layers; no cell ids, which belong to the sheet's own cells.)
    const sub = Object.assign(Object.create(st), { lines: pane.lines, fills: pane.fills, body: pane.body, edges: pane.edges, selLayer: pane.selLayer, filterLayer: pane.filterLayer, noteLayer: pane.noteLayer, noteTexts: null, idBase: null });
    drawGridlines(sub, w);
    const { fills, edges } = decorations(sub, w);
    pane.fills.replaceChildren(...fills);
    drawRows(sub, w);
    pane.edges.replaceChildren(...edges);
    drawSelection(sub);
    drawFilterButtons(sub, w);
    drawNoteMarks(sub, w);
  };
  // (From where the freeze starts: the window's first row and column when it was frozen.)
  const top = st.freeze?.top || 0;
  const left = st.freeze?.left || 0;
  paint(st.panes.top, fr ? { r0: top, r1: top + fr - 1, c0: win.c0, c1: win.c1 } : null);
  paint(st.panes.left, fc ? { r0: win.r0, r1: win.r1, c0: left, c1: left + fc - 1 } : null);
  paint(st.panes.corner, fr && fc ? { r0: top, r1: top + fr - 1, c0: left, c1: left + fc - 1 } : null);
}

export function drawGridlines(st, win) {
  if (!st.sheet.showGridLines) return st.lines.replaceChildren();
  const top = st.rows.start(win.r0);
  const left = st.cols.start(win.c0);
  const height = st.rows.start(win.r1 + 1) - top;
  const width = st.cols.start(win.c1 + 1) - left;
  const kids = [];
  // (A cell's gridlines are its right and bottom pixel, inside its own width and height.)
  for (let c = win.c0; c <= win.c1; c++) {
    if (!st.cols.size(c)) continue;
    const el = div('oxe-vline');
    setVars(el, { '--oxe-x': st.cols.start(c + 1) - 1, '--oxe-y': top, '--oxe-h': height });
    kids.push(el);
  }
  for (let r = win.r0; r <= win.r1; r++) {
    if (!st.rows.size(r)) continue;
    const el = div('oxe-hline');
    setVars(el, { '--oxe-x': left, '--oxe-y': st.rows.start(r + 1) - 1, '--oxe-w': width });
    kids.push(el);
  }
  st.lines.replaceChildren(...kids);
}

export function drawRows(st, win) {
  const kids = [];
  const sheet = st.sheet;
  // Merged ranges that start above the view are drawn with the first row in view.
  const pending = sheet.merges.filter((m) => m.r0 < win.r0 && m.r1 >= win.r0 && m.c1 >= win.c0 && m.c0 <= win.c1);
  for (let r = win.r0; r <= win.r1; r++) {
    const h = st.rows.size(r);
    if (!h) continue;
    const row = div('oxe-row', 'row');
    row.setAttribute('aria-rowindex', String(r + 2));
    const top = st.rows.start(r);
    setVars(row, { '--oxe-y': top, '--oxe-h': h });
    const head = div('oxe-rowhead', 'rowheader');
    head.textContent = String(r + 1);
    row.appendChild(head);
    const line = sheet.cells.get(r);
    const cells = [];
    for (const m of pending.splice(0)) cells.push([m.r0, m.c0]);
    if (line) {
      // Text in a cell left of the view may run on into it.
      for (let c = win.c0 - 1; c >= Math.max(0, win.c0 - SPILL_BACK); c--) {
        const cell = line.get(c);
        if (cell?.v == null) continue;
        if (cell.t === 's' || cell.t === 'str') cells.push([r, c]);
        break;
      }
      for (let c = win.c0; c <= win.c1; c++) if (line.get(c) && st.cols.size(c)) cells.push([r, c]);
    }
    // (A merged range is drawn even when its top-left cell is empty: it hides the gridlines inside it.)
    for (const m of sheet.merges) if (m.r0 === r && m.c1 >= win.c0 && m.c0 <= win.c1) cells.push([m.r0, m.c0]);
    const seen = new Set();
    for (const [cr, cc] of cells) {
      if (seen.has(`${cr},${cc}`)) continue;
      seen.add(`${cr},${cc}`);
      const merge = mergeAt(st, cr, cc);
      // (Inside a merged range only the top-left cell shows; the others are kept in the file but not drawn.)
      if (merge && (merge.r0 !== cr || merge.c0 !== cc)) continue;
      const cell = sheet.cells.get(cr)?.get(cc) || { t: 'n', v: null };
      if (cell.v == null && !merge && !(st.sel && st.sel.ar === cr && st.sel.ac === cc)) continue;
      row.appendChild(cellElement(st, cr, cc, cell, sheet.cells.get(cr) || new Map(), win, top));
    }
    // (The active cell is drawn even when it is empty, so that the grid can name it as its active descendant.)
    const sel = st.sel;
    if (sel && sel.ar === r && sel.ac >= win.c0 && sel.ac <= win.c1 && !seen.has(`${r},${sel.ac}`) && !mergeAt(st, r, sel.ac)) {
      row.appendChild(cellElement(st, r, sel.ac, { t: 'n', v: null }, line || new Map(), win, top));
    }
    kids.push(row);
  }
  st.body.replaceChildren(...kids);
  fitNumbers(st.body);
}

/** Redraw if the window in view changed (or `force`). */

/** While a heading edge is dragged: a line where the edge will be and Excel's tip ("Width: 20.00 (145 pixels)",
 *  "Height: 30.00 (40 pixels)"); no axis takes them away. */
export function drawGuide(st, axis, index, size) {
  const view = st.view;
  st.canvas.querySelectorAll(':scope > .oxe-guide, :scope > .oxe-guide-tip').forEach((el) => el.remove());
  if (axis == null) return;
  const line = div('oxe-guide');
  const tip = div('oxe-guide-tip');
  const z = st.zoom / 100;
  const top = view.scrollTop / z;
  const left = view.scrollLeft / z;
  if (axis === 'col') {
    // (A frozen column's edge doesn't scroll.)
    const x = st.rowHeadW + st.cols.start(index) + size + (index >= st.freeze.left && index < st.freeze.left + st.freeze.cols ? left - st.paneX0 : 0);
    setVars(line, { '--oxe-x': x, '--oxe-y': top, '--oxe-w': 1, '--oxe-h': view.clientHeight / z });
    setVars(tip, { '--oxe-x': x + 4, '--oxe-y': top + 2 });
    tip.textContent = `Width: ${(Math.max(0, size - 5) / MDW).toFixed(2)} (${Math.round(size)} pixels)`;
  } else {
    const y = st.colHeadH + st.rows.start(index) + size + (index >= st.freeze.top && index < st.freeze.top + st.freeze.rows ? top - st.paneY0 : 0);
    setVars(line, { '--oxe-x': left, '--oxe-y': y, '--oxe-w': view.clientWidth / z, '--oxe-h': 1 });
    setVars(tip, { '--oxe-x': left + st.rowHeadW + 4, '--oxe-y': y + 4 });
    tip.textContent = `Height: ${(Math.round(size) * 0.75).toFixed(2)} (${Math.round(size)} pixels)`;
  }
  st.canvas.append(line, tip);
}
