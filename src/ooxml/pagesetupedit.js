// Page Layout > Page Setup: Print Area (Set, Add to, Clear) and Breaks (Insert, Remove, Reset All), written as Excel
// writes them. Measured over COM in a hidden Excel (the ribbon's commands through ExecuteMso where it allows, else
// PageSetup.PrintArea; docs/PRINT.md):
// - Insert Page Break at C5: a manual row break before row 5 and a column break before C, <brk id="4" max="16383"
//   man="1"/> in rowBreaks and <brk id="2" max="1048575" man="1"/> in colBreaks, each with count and
//   manualBreakCount; in row 1 only the column's, in column A only the row's; at A1 nothing (Excel refused).
// - Remove Page Break: the break above the active row and the one left of its column, those there (E3, with breaks at
//   rows 5 and 9 and columns C and E: E's alone).
// - Reset All Page Breaks: every manual break gone, the elements too.
// - The print area is the sheet's own name _xlnm.Print_Area (localSheetId), its ranges $-absolute with the sheet's name
//   ('My Data'!$B$2:$D$5, two joined by a comma; one cell 'My Data'!$C$3); cleared, the name goes.
import { kids, kid } from './xml.js';
import { rangeName, parseRange, colIndex, colName } from './refs.js';
import { sheetPrefix } from '../calc/tokens.js';
import { rangesOf } from './printarea.js';
import { findName, defineName, editName, deleteName } from './names.js';

const AFTER_ROW = ['colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
const AFTER_COL = AFTER_ROW.slice(1);
const PRINT_AREA = '_xlnm.Print_Area';

/** The sheet's rowBreaks or colBreaks (`axis` 'row' or 'col') as their ids, manual ones; `make`: the element made in
 *  its place when missing. */
function breaksEl(doc, axis, make) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const tag = axis === 'row' ? 'rowBreaks' : 'colBreaks';
  let el = kid(root, tag, ns);
  if (!el && make) {
    el = doc.createElementNS(ns, root.prefix ? `${root.prefix}:${tag}` : tag);
    const after = axis === 'row' ? AFTER_ROW : AFTER_COL;
    root.insertBefore(el, [...root.childNodes].find((n) => n.nodeType === 1 && after.includes(n.localName)) || null);
  }
  return el;
}

/** The breaks element given these brk ids (added or taken away), sorted, counted as Excel counts them; gone when none
 *  are left. */
function writeBreaks(doc, axis, change) {
  const el = breaksEl(doc, axis, true);
  const ns = el.namespaceURI;
  const p = el.prefix ? `${el.prefix}:` : '';
  const brks = kids(el, 'brk', ns);
  change(brks, (id) => {
    const b = doc.createElementNS(ns, `${p}brk`);
    b.setAttribute('id', String(id));
    b.setAttribute('max', axis === 'row' ? '16383' : '1048575');
    b.setAttribute('man', '1');
    return b;
  }, el);
  const all = kids(el, 'brk', ns).sort((a, b) => Number(a.getAttribute('id')) - Number(b.getAttribute('id')));
  for (const b of all) el.appendChild(b);
  if (!all.length) { el.parentNode.removeChild(el); return; }
  el.setAttribute('count', String(all.length));
  el.setAttribute('manualBreakCount', String(all.filter((b) => ['1', 'true'].includes(b.getAttribute('man'))).length));
}

const hasBreak = (doc, axis, id) => kids(breaksEl(doc, axis, false), 'brk', doc.documentElement.namespaceURI).some((b) => Number(b.getAttribute('id')) === id);

/** Insert Page Break at cell (r, c) of sheet si: false at A1, where Excel adds none. */
export function insertBreak(wb, si, r, c) {
  if (r === 0 && c === 0) return false;
  const part = wb._part(si);
  const doc = part.doc;
  if (r > 0 && !hasBreak(doc, 'row', r)) writeBreaks(doc, 'row', (brks, make, el) => el.appendChild(make(r)));
  if (c > 0 && !hasBreak(doc, 'col', c)) writeBreaks(doc, 'col', (brks, make, el) => el.appendChild(make(c)));
  wb._reread(si, part);
  return true;
}

/** Remove Page Break at cell (r, c) of sheet si: the breaks above its row and left of its column, those there; false
 *  when there are none. */
export function removeBreak(wb, si, r, c) {
  const part = wb._part(si);
  const doc = part.doc;
  const row = hasBreak(doc, 'row', r);
  const col = hasBreak(doc, 'col', c);
  if (!row && !col) return false;
  const drop = (id) => (brks) => { for (const b of brks) if (Number(b.getAttribute('id')) === id) b.parentNode.removeChild(b); };
  if (row) writeBreaks(doc, 'row', drop(r));
  if (col) writeBreaks(doc, 'col', drop(c));
  wb._reread(si, part);
  return true;
}

/** Reset All Page Breaks on sheet si: false when it has none. */
export function resetBreaks(wb, si) {
  const part = wb._part(si);
  const doc = part.doc;
  const els = ['row', 'col'].map((axis) => breaksEl(doc, axis, false)).filter(Boolean);
  if (!els.length) return false;
  for (const el of els) el.parentNode.removeChild(el);
  wb._reread(si, part);
  return true;
}

/** Sheet si's print area as ranges ({ r0, c0, r1, c1 }), or null. */
export function printAreaOf(wb, si) {
  const d = findName(wb, PRINT_AREA, si);
  const out = d ? rangesOf(d.text).filter((a) => !a.rows && !a.cols).map(({ r0, c0, r1, c1 }) => ({ r0, c0, r1, c1 })) : [];
  return out.length ? out : null;
}

/** Set Print Area of sheet si to `areas` (the selection's ranges), or with `add` Add to Print Area (its ranges after
 *  those there; Excel's joining of neighbours not measured). */
export function setPrintArea(wb, si, areas, { add = false } = {}) {
  const all = [...(add ? printAreaOf(wb, si) || [] : []), ...areas];
  const prefix = sheetPrefix(wb.sheets[si].name);
  const abs = (a) => rangeName(a).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2');
  const parts = all.map((a) => `${prefix}${abs(a)}`);
  // (Defined with its first range, then given them all: Excel's list of ranges joined by commas, which the name
  // formula reader here doesn't take at the top level.)
  if (findName(wb, PRINT_AREA, si)) editName(wb, { name: PRINT_AREA, sheet: si }, { refersTo: `=${parts[0]}` });
  else defineName(wb, { name: PRINT_AREA, refersTo: `=${parts[0]}`, sheet: si });
  const d = findName(wb, PRINT_AREA, si);
  d.text = parts.join(',');
  d.el.textContent = d.text;
  wb.bookDirty = true;
}

/** Clear Print Area of sheet si: false when it has none. */
export function clearPrintArea(wb, si) {
  if (!findName(wb, PRINT_AREA, si)) return false;
  deleteName(wb, { name: PRINT_AREA, sheet: si });
  return true;
}

const PRINT_TITLES = '_xlnm.Print_Titles';

/** Sheet si's print titles: { rows: [r0, r1] or null, cols: [c0, c1] or null }. */
export function printTitlesOf(wb, si) {
  const d = findName(wb, PRINT_TITLES, si);
  const out = { rows: null, cols: null };
  for (const a of d ? rangesOf(d.text) : []) {
    if (a.rows) out.rows = [a.r0, a.r1];
    if (a.cols) out.cols = [a.c0, a.c1];
  }
  return out;
}

/** What is typed for Rows to repeat at top (`axis` 'rows') or Columns to repeat at left ('cols'), as Excel takes it:
 *  [first, last], '' as null, or undefined when it isn't one. Measured: "1:2" and "$1:$2" rows 1 to 2, a range
 *  ("B2:C4") its rows (2 to 4), a lone "$3" refused; columns the same ("A:B", "$A:$A", a range's columns). */
export function readTitles(text, axis) {
  const t = String(text || '').trim().replace(/^(?:'(?:[^']|'')+'|[^!]+)!/, '').replace(/\$/g, '').toUpperCase();
  if (!t) return null;
  const lines = axis === 'rows' ? /^(\d+):(\d+)$/.exec(t) : /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(t);
  if (lines) {
    const [a, b] = axis === 'rows' ? [Number(lines[1]) - 1, Number(lines[2]) - 1] : [colIndex(lines[1]), colIndex(lines[2])];
    return a >= 0 && b >= 0 ? [Math.min(a, b), Math.max(a, b)] : undefined;
  }
  const r = /^[A-Z]{1,3}\d+:[A-Z]{1,3}\d+$/.test(t) ? parseRange(t) : null;
  if (!r) return undefined;
  return axis === 'rows' ? [r.r0, r.r1] : [r.c0, r.c1];
}

/** Sheet si's print titles set to `titles` ({ rows, cols }: [first, last] or null), as Excel saves them: columns
 *  first, then rows, each with the sheet's name ('Q1 Plan'!$A:$B,'Q1 Plan'!$1:$2); neither, the name gone. */
export function setPrintTitles(wb, si, { rows = null, cols = null }) {
  const prefix = sheetPrefix(wb.sheets[si].name);
  const parts = [];
  if (cols) parts.push(`${prefix}$${colName(cols[0])}:$${colName(cols[1])}`);
  if (rows) parts.push(`${prefix}$${rows[0] + 1}:$${rows[1] + 1}`);
  const had = findName(wb, PRINT_TITLES, si);
  if (!parts.length) { if (had) deleteName(wb, { name: PRINT_TITLES, sheet: si }); return; }
  if (had) editName(wb, { name: PRINT_TITLES, sheet: si }, { refersTo: `=${parts[0]}` });
  else defineName(wb, { name: PRINT_TITLES, refersTo: `=${parts[0]}`, sheet: si });
  const d = findName(wb, PRINT_TITLES, si);
  d.text = parts.join(',');
  d.el.textContent = d.text;
  wb.bookDirty = true;
}
