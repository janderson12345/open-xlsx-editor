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
import { rangeName } from './refs.js';
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
