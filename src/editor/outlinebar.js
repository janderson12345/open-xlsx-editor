// The outline bar, as Excel draws it beside the row headings and over the column headings (docs/OUTLINE.md): a slot
// 19px wide for each level (Excel's cells start 43px further in with one level grouped, 19px more for each level after,
// measured over COM in excel-outline.ps1), the level buttons 1, 2, 3... in the corner, and for each group a line along
// its lines with a − button at its summary line, a + when its lines are hidden. Clicked, the grid's onOutline hears
// { axis, level } (a level button) or { axis, line } (a group's button). View chrome, hidden from assistive technology:
// the same commands are on the Data tab. Frozen rows and columns keep their part of it in place (pinned strips).
import { div, setVars } from './dom.js';
import { OUTLINE_SLOT } from './geometry.js';
import { outlineOf } from '../ooxml/outline.js';

/** The button's size and its place in a slot (not measured: Excel's glyphs aren't in its PDF or COM). */
const BUTTON = 15;
const INSET = 4;

/** A strip's clicks: a button's to onOutline, all of them kept from the grid's own mouse (no heading selected). */
function strip(st, cls) {
  const el = div(cls);
  el.setAttribute('aria-hidden', 'true');
  el.addEventListener('mousedown', (e) => { e.stopPropagation(); e.preventDefault(); });
  el.addEventListener('click', (e) => {
    const b = e.target.closest('[data-axis]');
    if (!b) return;
    e.stopPropagation();
    const axis = b.dataset.axis;
    st.onOutline?.(b.dataset.level ? { axis, level: Number(b.dataset.level) } : { axis, line: Number(b.dataset.line) });
  });
  return el;
}

const button = (cls, text, data, vars) => {
  const b = div(`oxe-ol-button ${cls}`);
  b.textContent = text;
  for (const [k, v] of Object.entries(data)) b.dataset[k] = String(v);
  b.title = data.level ? `Show level ${data.level}` : text === '+' ? 'Show Detail' : 'Hide Detail';
  setVars(b, vars);
  return b;
};

/** One axis's groups in view: a line along each group shown, its button at the summary line. */
function groupsOf(st, axis, groups, from, to) {
  const lines = axis === 'row' ? st.rows : st.cols;
  const kids = [];
  for (const g of groups) {
    if (g.end < from - 1 || g.start > to + 1) continue;
    let shown = false;
    for (let i = g.start; i <= g.end && !shown; i++) if (lines.size(i)) shown = true;
    const slot = (g.level - 1) * OUTLINE_SLOT + INSET;
    const mid = slot + Math.round(BUTTON / 2);
    if (shown) {
      const a = lines.start(g.start);
      const b = lines.start(g.end + 1);
      const line = div(`oxe-ol-line oxe-ol-${axis} ${g.summary != null && g.summary < g.start ? 'oxe-ol-end' : 'oxe-ol-start'}`);
      setVars(line, axis === 'row' ? { '--oxe-x': mid, '--oxe-y': a + 2, '--oxe-h': Math.max(1, b - a - 4) } : { '--oxe-x': a + 2, '--oxe-y': mid, '--oxe-w': Math.max(1, b - a - 4) });
      kids.push(line);
    }
    if (g.summary == null || !lines.size(g.summary)) continue;
    const at = lines.start(g.summary) + Math.round((lines.size(g.summary) - BUTTON) / 2);
    kids.push(button('oxe-ol-toggle', shown ? '−' : '+', { axis, line: g.summary }, axis === 'row' ? { '--oxe-x': slot, '--oxe-y': at } : { '--oxe-x': at, '--oxe-y': slot }));
  }
  return kids;
}

/** Draw the outline bar for the part of the sheet in view (grid.js draw, after the column headings). */
export function drawOutline(st, win) {
  if (!st.outlineRows) {
    st.outlineRows = strip(st, 'oxe-outline-rows');
    st.colheads.after(st.outlineRows);
    st.outlineCols = strip(st, 'oxe-outline-cols');
    st.outlineLevels = strip(st, 'oxe-outline-levels');
    st.corner.append(st.outlineLevels);
    // (Frozen rows and columns keep their part of the bar in place, over the frozen panes: pinned strips of their own.)
    st.outlineRowsPin = strip(st, 'oxe-outline-rows oxe-outline-rows-pin');
    st.panes.top.wrap.after(st.outlineRowsPin);
    st.outlineColsPin = strip(st, 'oxe-outline-cols oxe-outline-cols-pin');
  }
  const o = st.outlineW || st.outlineH ? outlineOf(st.sheet) : null;
  st.outlineRows.replaceChildren(...(o && st.outlineW ? groupsOf(st, 'row', o.rows.groups, win.r0, win.r1) : []));
  st.outlineCols.replaceChildren(...(o && st.outlineH ? groupsOf(st, 'col', o.cols.groups, win.c0, win.c1) : []));
  if (st.outlineH) st.colheads.append(st.outlineCols); else st.outlineCols.remove();
  const fr = st.freeze?.rows || 0;
  const fc = st.freeze?.cols || 0;
  st.outlineRowsPin.replaceChildren(...(o && st.outlineW && fr ? groupsOf(st, 'row', o.rows.groups, st.freeze.top || 0, (st.freeze.top || 0) + fr - 1) : []));
  st.outlineColsPin.replaceChildren(...(o && st.outlineH && fc ? groupsOf(st, 'col', o.cols.groups, st.freeze.left || 0, (st.freeze.left || 0) + fc - 1) : []));
  if (st.outlineH && fc) st.pinHeads.firstChild.append(st.outlineColsPin); else st.outlineColsPin.remove();
  // (The level buttons: the rows' along the column headings' band, the columns' down the row headings'.)
  const levels = [];
  const headsH = st.colHeadH - st.outlineH;
  const headsW = st.rowHeadW - st.outlineW;
  if (o && st.outlineW) for (let k = 0; k <= o.rows.level; k++) levels.push(button('oxe-ol-level', String(k + 1), { axis: 'row', level: k + 1 }, { '--oxe-x': k * OUTLINE_SLOT + INSET, '--oxe-y': st.outlineH + Math.max(0, Math.round((headsH - BUTTON) / 2)) }));
  if (o && st.outlineH) for (let k = 0; k <= o.cols.level; k++) levels.push(button('oxe-ol-level', String(k + 1), { axis: 'col', level: k + 1 }, { '--oxe-x': st.outlineW + Math.max(0, Math.round((headsW - BUTTON) / 2)), '--oxe-y': k * OUTLINE_SLOT + INSET }));
  st.outlineLevels.replaceChildren(...levels);
}
