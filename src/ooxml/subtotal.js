// Data > Subtotal and its Remove All, as Excel does them (tools/excel-compare/excel-outline.ps1, Range.Subtotal and
// RemoveSubtotal, 12 cases, docs/OUTLINE.md). At each change in a column of the list (its first row the headings), a row
// is inserted under the group (over it with Summary below data off) holding "<value> Total" in that column, bold, and
// =SUBTOTAL(n,C2:C3) over the group's own rows in each column chosen; a Grand Total row under the list (over it) over
// all its rows. The list is outlined: its rows a level deeper than the subtotal rows, the first time level 2 under level
// 1 (the grand total at none). Subtotals added to subtotals (Replace current subtotals off) go inside the ones there:
// next to their group's rows, the rows one level deeper again, and their grand total nearest the list's rows. Page break
// between groups puts a manual break after each group's subtotal, the last after the grand total. Remove All deletes
// the rows holding SUBTOTAL and the list's outline and page breaks.
import { insertBreak } from './pagesetupedit.js';
import { clearOutline, lineLevels, setSummary } from './outline.js';
import { colName, refName, parseRef } from './refs.js';
import { kids, kid, attr } from './xml.js';
import { pinRefs, forgetRows } from './sheetedit.js';
import { styleSummaryLines } from './outlinestyles.js';

/** Excel's Use function list: SUBTOTAL's number and the word its rows are labelled with. */
export const SUBTOTAL_FUNCTIONS = {
  sum: [9, 'Total'], count: [3, 'Count'], average: [1, 'Average'], max: [4, 'Max'], min: [5, 'Min'], product: [6, 'Product'],
  countNums: [2, 'Count'], stdDev: [7, 'StdDev'], stdDevp: [8, 'StdDevp'], var: [10, 'Var'], varp: [11, 'Varp'],
};
const IS_SUBTOTAL = /^\s*(_xlfn\.)?SUBTOTAL\s*\(/i;

/** The rows of `area` (sheet si's list) holding a SUBTOTAL in one of its columns. */
function subtotalRows(sheet, area) {
  const out = [];
  for (let r = area.r0 + 1; r <= area.r1; r++) {
    const line = sheet.cells.get(r);
    if (!line) continue;
    for (let c = area.c0; c <= area.c1; c++) if (IS_SUBTOTAL.test(line.get(c)?.f || '')) { out.push(r); break; }
  }
  return out;
}

/** Remove All: the list's subtotal rows deleted, its rows' outline cleared and its manual page breaks taken away.
 *  Returns the list's range after it. */
export function removeSubtotals(wb, si, area) {
  const rows = subtotalRows(wb.sheet(si), area);
  // (The list's breaks, taken before its rows go: deleting rows leaves a sheet's breaks where they are.)
  const breaks = wb.sheet(si).page.breaks.rows.filter((id) => id > area.r0 && id <= area.r1 + 1);
  // (Bottom up, runs of neighbouring rows deleted at once.)
  for (let k = rows.length - 1; k >= 0;) {
    let a = k;
    while (a > 0 && rows[a - 1] === rows[a] - 1) a--;
    wb.changeLines(si, 'row', rows[a], -(k - a + 1));
    k = a - 1;
  }
  const r1 = area.r1 - rows.length;
  clearOutline(wb, si, 'row', area.r0, r1);
  if (breaks.length) {
    const part = wb._part(si);
    const root = part.doc.documentElement;
    const el = [...root.childNodes].find((n) => n.localName === 'rowBreaks');
    for (const b of [...el.childNodes].filter((n) => n.localName === 'brk' && breaks.includes(Number(n.getAttribute('id'))))) el.removeChild(b);
    const left = [...el.childNodes].filter((n) => n.localName === 'brk');
    if (!left.length) root.removeChild(el);
    else { el.setAttribute('count', String(left.length)); el.setAttribute('manualBreakCount', String(left.filter((n) => n.getAttribute('man') === '1').length)); }
    wb._reread(si, part);
  }
  return { ...area, r1 };
}

/** Subtotal's rows over their groups (Summary below data off) take the formats of the row under them, as Excel's do
 *  ("subtotal formatted cells above", "subtotal formatted row above": the Grand row italic as the first row of data,
 *  not bold as the headings): its own format and height, and the cells whose format isn't the row's, as an inserted
 *  row takes the row above's (rowcol.js). `rows`, the new rows (empty yet); bottom up, so each copies a finished one. */
function formatFromBelow(wb, si, rows) {
  const part = wb._part(si);
  const doc = part.doc;
  const ns = doc.documentElement.namespaceURI;
  const sheetData = kid(doc.documentElement, 'sheetData', ns);
  pinRefs(sheetData, ns);
  forgetRows(doc);
  const rowEl = (r) => kids(sheetData, 'row', ns).find((x) => Number(attr(x, 'r')) - 1 === r);
  const make = (name) => doc.createElementNS(ns, doc.documentElement.prefix ? `${doc.documentElement.prefix}:${name}` : name);
  for (const r of [...rows].sort((a, b) => b - a)) {
    const source = rowEl(r + 1);
    let row = rowEl(r);
    if (!row && !source) continue;
    if (!row) { row = make('row'); row.setAttribute('r', String(r + 1)); sheetData.insertBefore(row, source); }
    for (const name of ['s', 'customFormat', 'ht', 'customHeight']) { if (attr(source, name) != null) row.setAttribute(name, attr(source, name)); else row.removeAttribute(name); }
    for (const c of kids(row, 'c', ns)) row.removeChild(c);
    const own = attr(source, 'customFormat') === '1' ? Number(attr(source, 's') || 0) : 0;
    for (const c of kids(source, 'c', ns)) {
      if (Number(attr(c, 's') || 0) === own) continue;
      const el = make('c');
      el.setAttribute('r', refName(r, parseRef(attr(c, 'r')).c));
      el.setAttribute('s', attr(c, 's'));
      row.appendChild(el);
    }
  }
  wb._reread(si, part);
}

/** The text a cell's value is labelled with ("East" for "East Total"). */
const labelOf = (cell) => (cell?.v == null ? '' : typeof cell.v === 'boolean' ? (cell.v ? 'TRUE' : 'FALSE') : String(cell.v));

/**
 * Subtotal over the list `area` of sheet si (its first row the headings): at each change in column `by`, `fn` (a key of
 * SUBTOTAL_FUNCTIONS) of columns `cols`; `replace` the subtotals there, `pageBreaks` between groups, `below` the data
 * (else above). Returns the list's range after it, or null when it has no rows under its headings.
 */
export function subtotal(wb, si, area, { by, fn = 'sum', cols, replace = true, pageBreaks = false, below = true }) {
  if (replace) area = removeSubtotals(wb, si, area);
  const sheet = wb.sheet(si);
  const totals = new Set(subtotalRows(sheet, area));
  const levelOf = new Map(lineLevels(wb, si, 'row'));
  const data = [];
  for (let r = area.r0 + 1; r <= area.r1; r++) if (!totals.has(r)) data.push(r);
  if (!data.length) return null;
  // (Groups: neighbouring rows alike in column `by`, case aside; a subtotal row between rows parts them too.)
  const groups = [];
  for (const r of data) {
    const key = labelOf(sheet.cells.get(r)?.get(by)).toLowerCase();
    const last = groups[groups.length - 1];
    if (last && last.key === key && last.rows[last.rows.length - 1] === r - 1) last.rows.push(r);
    else groups.push({ key, label: labelOf(sheet.cells.get(r)?.get(by)), rows: [r] });
  }
  // (The grand totals there: subtotal rows outside the outline, at the list's end, or its start above.)
  const grands = [...totals].filter((r) => !levelOf.get(r));
  // The list as it will be, row by row: the rows there ({ r }) and the new ones ({ group } or { grand }).
  const order = [];
  const firstGrandAt = below ? (grands.length ? Math.min(...grands) : area.r1 + 1) : (grands.length ? Math.max(...grands) + 1 : area.r0 + 1);
  for (let r = area.r0 + 1; r <= area.r1 + 1; r++) {
    if (r === firstGrandAt) order.push({ grand: true });
    for (const g of groups) {
      if (!below && g.rows[0] === r) order.push({ group: g });
      if (below && g.rows[g.rows.length - 1] === r - 1) order.push({ group: g });
    }
    if (r <= area.r1) order.push({ r });
  }
  // (`below`: a group's row comes after its last row, before the grand total when they meet at the list's end.)
  if (below) {
    const gi = order.findIndex((x) => x.grand);
    const lastGroup = order.findIndex((x) => x.group === groups[groups.length - 1]);
    if (gi >= 0 && lastGroup > gi) { const [gr] = order.splice(gi, 1); order.splice(lastGroup, 0, gr); }
  }
  const at = (k) => area.r0 + 1 + k;
  order.forEach((x, k) => { if (!('r' in x)) wb.changeLines(si, 'row', at(k), 1); });
  // (Summaries above: each new row formatted as the row under it, not the one over it, bottom up.)
  if (!below) formatFromBelow(wb, si, order.map((x, k) => ('r' in x ? null : at(k))).filter((r) => r != null));
  // Each kept row's place after, the new rows written.
  const place = new Map();
  order.forEach((x, k) => { if ('r' in x) place.set(x.r, at(k)); else x.at = at(k); });
  const [num, word] = SUBTOTAL_FUNCTIONS[fn];
  const firstData = place.get(data[0]);
  const lastData = place.get(data[data.length - 1]);
  const contents = [];
  const bold = [];
  for (const x of order) {
    if ('r' in x) continue;
    const [a, b] = x.grand ? [firstData, lastData] : [place.get(x.group.rows[0]), place.get(x.group.rows[x.group.rows.length - 1])];
    contents.push({ r: x.at, c: by, value: x.grand ? `Grand ${word}` : `${x.group.label} ${word}` });
    bold.push({ r: x.at, c: by });
    for (const c of cols) contents.push({ r: x.at, c, formula: `SUBTOTAL(${num},${colName(c)}${a + 1}:${colName(c)}${b + 1})` });
  }
  wb.apply(si, contents);
  const now = wb.sheet(si);
  // (Automatic styles on, the rows take their level's style instead of the bold label: "subtotal with automatic styles".)
  const styled = !!now.outline?.styles;
  if (!styled) wb.setStyles(si, bold.map(({ r, c }) => ({ r, c, s: wb.styleWith(now.cells.get(r)?.get(c)?.s || 0, { font: { b: true } }) })));
  // The outline: the list's rows a level deeper than their new subtotal rows (level 2 under 1 the first time).
  const levels = [];
  for (const g of groups) {
    const d = levelOf.get(g.rows[0]) || 0;
    const [sub, rows] = d ? [d, d + 1] : [1, 2];
    for (const r of g.rows) levels.push([place.get(r), rows]);
    levels.push([order.find((x) => x.group === g).at, sub]);
  }
  // (The grand total outside the outline, though inserted rows take the level of the row above.)
  for (const x of order) if (x.grand) levels.push([x.at, 0]);
  lineLevels(wb, si, 'row', levels);
  setSummary(wb, si, { below });
  if (styled) styleSummaryLines(wb, si, 'row', order.filter((x) => !('r' in x)).map((x) => x.at));
  if (pageBreaks) {
    const ends = groups.map((g) => (below ? order.find((x) => x.group === g).at : place.get(g.rows[g.rows.length - 1])));
    const grand = order.find((x) => x.grand).at;
    if (below) ends[ends.length - 1] = grand;
    for (const end of ends) insertBreak(wb, si, end + 1, 0);
  }
  wb.resetCalc?.();
  return { ...area, r1: area.r1 + order.length - (area.r1 - area.r0) };
}
