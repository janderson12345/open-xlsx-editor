// The outline's automatic styles, as Excel applies them (tools/excel-compare/excel-outline.ps1: Outline.AutomaticStyles,
// Range.ApplyOutlineStyles, AutoOutline and Subtotal with the styles on, 9 cases; docs/OUTLINE.md). A group's summary
// row takes the built-in style RowLevel_L of the group's level L (builtinId 1, iLevel L-1), a summary column ColLevel_L
// (builtinId 2): level 1 bold, level 2 italic, deeper the Normal font. Where a styled row and column meet, the lower
// level wins, and the row on a tie (ApplyOutlineStyles: E5, row level 2 and column level 1, ColLevel_1; E8, both level
// 1, RowLevel_1). Apply Styles styles the cells of a range (empty ones made); Group and Subtotal with Automatic styles
// on give the summary rows (columns) the style as their own format, the cells there too; Create (Auto Outline) styles
// the cells of the used range.
// (The styles themselves, RowLevel_n and ColLevel_n, are declared with Excel's other built-in ones: stylesedit.js.)
import { outlineOf } from './outline.js';

/** For each summary line, the lowest level of the groups it sums up: Map(index => level). */
function summaries(groups) {
  const out = new Map();
  for (const g of groups) if (g.summary != null && !(out.get(g.summary) <= g.level)) out.set(g.summary, g.level);
  return out;
}

/** The level style of cell (r, c), or null: the lower level of its row's and column's, the row's on a tie. */
function styleAt(rows, cols, r, c) {
  const a = rows.get(r);
  const b = cols.get(c);
  if (a == null && b == null) return null;
  return b == null || (a != null && a <= b) ? `RowLevel_${a}` : `ColLevel_${b}`;
}

/** Apply Styles over `area` of sheet si ({ r0, c0, r1, c1 }): each cell on a summary row or column given its level
 *  style. Returns how many cells were styled. */
export function applyOutlineStyles(wb, si, area) {
  const sheet = wb.sheet(si);
  const o = outlineOf(sheet);
  const rows = summaries(o.rows.groups);
  const cols = summaries(o.cols.groups);
  const cells = [];
  const visit = (r, c) => {
    const name = styleAt(rows, cols, r, c);
    if (name) cells.push({ r, c, s: wb.styleWith(sheet.cells.get(r)?.get(c)?.s || 0, { style: name }) });
  };
  // (Only the summary lines in the area: a whole row or column selected isn't walked cell by cell.)
  for (const r of rows.keys()) if (r >= area.r0 && r <= area.r1) for (let c = area.c0; c <= Math.min(area.c1, Math.max(sheet.maxCol, ...cols.keys())); c++) visit(r, c);
  for (const c of cols.keys()) if (c >= area.c0 && c <= area.c1) for (let r = area.r0; r <= Math.min(area.r1, Math.max(sheet.maxRow, ...rows.keys())); r++) if (!rows.has(r)) visit(r, c);
  if (cells.length) wb.setStyles(si, cells);
  return cells.length;
}

/** Automatic styles after Group or Subtotal: the summary lines of `axis` among `lines` (indexes) given their level
 *  style as their own format, and the cells already there. */
export function styleSummaryLines(wb, si, axis, lines) {
  const sheet = wb.sheet(si);
  const o = outlineOf(sheet);
  const levels = summaries(axis === 'row' ? o.rows.groups : o.cols.groups);
  const own = [];
  const cells = [];
  for (const i of lines) {
    const level = levels.get(i);
    if (level == null) continue;
    const name = `${axis === 'row' ? 'Row' : 'Col'}Level_${level}`;
    const lineS = axis === 'row' ? sheet.rows.get(i)?.s || 0 : sheet.cols.find((x) => i >= x.min && i <= x.max)?.style || 0;
    own.push(axis === 'row' ? { r: i, s: wb.styleWith(lineS, { style: name }) } : { c: i, s: wb.styleWith(lineS, { style: name }) });
    if (axis === 'row') for (const [c, cell] of sheet.cells.get(i) || []) cells.push({ r: i, c, s: wb.styleWith(cell.s || 0, { style: name }) });
    else for (const [r, line] of sheet.cells) if (line.has(i)) cells.push({ r, c: i, s: wb.styleWith(line.get(i).s || 0, { style: name }) });
  }
  if (cells.length) wb.setStyles(si, cells);
  if (own.length) wb.setLineStyles(si, axis === 'row' ? { rows: own } : { cols: own });
}

/** The summary lines of the groups that lines a to b are part of (Group with Automatic styles on). */
export function summariesOf(sheet, axis, a, b) {
  const o = outlineOf(sheet);
  return (axis === 'row' ? o.rows.groups : o.cols.groups).filter((g) => g.end >= a && g.start <= b && g.summary != null).map((g) => g.summary);
}
