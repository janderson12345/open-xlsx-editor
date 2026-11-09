// Inserting and deleting whole rows and columns, as Excel does it (measured over COM, docs/ROWCOL.md). The sheet's
// rows and cells move in its XML, with its column widths, merged ranges, array formulas and the ranges its hyperlinks,
// conditional formats, data validations, filter and view hold; every formula of the workbook that reads the sheet
// follows (tokens.js shiftRefs), defined names and chart series too. New rows take the format of the row above, new
// columns the format and width of the column to their left. Afterwards the sheet is read again and the workbook worked
// out again. A sheet with pictures, charts, comments, tables or pivot tables is refused for now: their positions
// aren't moved yet.
import { kids, kid, attr } from './xml.js';
import { parseRef, refName, parseRange, rangeName, MAX_ROWS, MAX_COLS } from './refs.js';
import { readSheet } from './sheet.js';
import { pinRefs, forgetRows } from './sheetedit.js';
import { shiftLine, shiftSpan, shiftRefs } from '../calc/tokens.js';
import { SheetError, mapFormulas } from './sheetops.js';

export const ROWCOL_MESSAGES = {
  offSheet: 'To prevent possible loss of data, Excel cannot shift nonblank cells off of the worksheet. Select another location in which to insert new cells, or delete data from the end of your worksheet.',
  array: "You can't change part of an array.",
  parts: "Rows and columns can't be inserted or deleted yet on a sheet with pictures, charts, comments, tables or pivot tables.",
};
/** Relationships of a sheet whose parts hold positions this doesn't move yet. */
const ANCHORED = /\/(drawing|vmlDrawing|comments|table|pivotTable|threadedComment|slicer|timeline)$/;

const make = (doc, ns, name) => doc.createElementNS(ns, doc.documentElement.prefix ? `${doc.documentElement.prefix}:${name}` : name);

/** A range after the change, or null when all of it was deleted (shiftSpan on one axis). */
function shiftRange(range, axis, at, n) {
  if (axis === 'row') {
    const s = shiftSpan(range.r0, range.r1, at, n);
    return s && { ...range, r0: s[0], r1: Math.min(s[1], MAX_ROWS - 1) };
  }
  const s = shiftSpan(range.c0, range.c1, at, n);
  return s && { ...range, c0: s[0], c1: Math.min(s[1], MAX_COLS - 1) };
}
/** A space-separated list of ranges (sqref) after the change; '' when none is left. */
const shiftSqref = (text, axis, at, n) => (text || '').split(/\s+/).filter(Boolean).map(parseRange).filter(Boolean)
  .map((r) => shiftRange(r, axis, at, n)).filter(Boolean).map(rangeName).join(' ');

/** Throws Excel's refusal when the change can't be made on sheet si. */
function check(wb, si, axis, at, n) {
  const info = wb.sheets[si];
  if (wb.pkg.readRels(info.path).some((r) => ANCHORED.test(r.type))) throw new SheetError(ROWCOL_MESSAGES.parts);
  const model = wb.sheet(si);
  const lines = axis === 'row' ? (a) => [a.r0, a.r1] : (a) => [a.c0, a.c1];
  for (const a of model.arrays) {
    const [a0, a1] = lines(a);
    // (An insert inside an array, or a delete of part of one, would split it.)
    const cut = n > 0 ? a0 < at && at <= a1 : a0 <= at - n - 1 && a1 >= at && !(a0 >= at && a1 <= at - n - 1);
    if (cut) throw new SheetError(ROWCOL_MESSAGES.array);
  }
  if (n > 0) {
    const last = axis === 'row' ? model.maxRow : model.maxCol;
    const limit = axis === 'row' ? MAX_ROWS : MAX_COLS;
    const filled = [...model.cells].some(([r, line]) => [...line].some(([c, cell]) => (axis === 'row' ? r : c) >= limit - n && (cell.v != null || cell.f != null)));
    if (last >= at && filled) throw new SheetError(ROWCOL_MESSAGES.offSheet);
  }
}

/** The cells of sheetData moved (and those deleted taken out); new lines formatted as the one before them. */
function moveCells(doc, axis, at, n) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const sheetData = kid(root, 'sheetData', ns);
  if (!sheetData) return;
  pinRefs(sheetData, ns);
  forgetRows(doc);
  const rowsEls = kids(sheetData, 'row', ns);
  // (The line before an insert, whose formats the new lines take.)
  const before = n > 0 && at > 0 ? at - 1 : -1;
  const rowStyle = (rowEl) => (attr(rowEl, 'customFormat') === '1' || attr(rowEl, 'customFormat') === 'true' ? Number(attr(rowEl, 's') || 0) : 0);
  const colStyleOf = colStyles(root, ns);
  const copies = [];
  if (axis === 'row') {
    const source = before >= 0 ? rowsEls.find((el) => Number(attr(el, 'r')) - 1 === before) : null;
    for (const rowEl of rowsEls) {
      const r = Number(attr(rowEl, 'r')) - 1;
      const next = shiftLine(r, at, n);
      if (next == null) { sheetData.removeChild(rowEl); continue; }
      if (next !== r) {
        rowEl.setAttribute('r', String(next + 1));
        for (const c of kids(rowEl, 'c', ns)) c.setAttribute('r', refName(next, parseRef(attr(c, 'r')).c));
      }
    }
    if (source) {
      // (Excel's copy of the row above: its height and format, and the cells whose format isn't the row's own.)
      const style = rowStyle(source);
      const cells = kids(source, 'c', ns).filter((c) => Number(attr(c, 's') || 0) !== style);
      const after = kids(sheetData, 'row', ns).find((el) => Number(attr(el, 'r')) - 1 >= at + n);
      for (let k = 0; k < n; k++) {
        const row = make(doc, ns, 'row');
        row.setAttribute('r', String(at + k + 1));
        for (const name of ['s', 'customFormat', 'ht', 'customHeight']) if (attr(source, name) != null) row.setAttribute(name, attr(source, name));
        for (const c of cells) {
          const el = make(doc, ns, 'c');
          el.setAttribute('r', refName(at + k, parseRef(attr(c, 'r')).c));
          el.setAttribute('s', attr(c, 's'));
          row.appendChild(el);
        }
        if (row.attributes.length > 1 || row.firstChild) copies.push([row, after]);
      }
    }
  } else {
    for (const rowEl of rowsEls) {
      const r = Number(attr(rowEl, 'r')) - 1;
      let source = null;
      for (const c of kids(rowEl, 'c', ns)) {
        const col = parseRef(attr(c, 'r')).c;
        if (col === before) source = c;
        const next = shiftLine(col, at, n);
        if (next == null) { rowEl.removeChild(c); continue; }
        if (next !== col) c.setAttribute('r', refName(r, next));
      }
      // (Excel's copy of the column to the left: its cells whose format isn't what a new cell there gets anyway.)
      const own = rowStyle(rowEl) || colStyleOf(before);
      if (source && Number(attr(source, 's') || 0) !== own) {
        const after = kids(rowEl, 'c', ns).find((c) => parseRef(attr(c, 'r')).c >= at + n);
        for (let k = 0; k < n; k++) {
          const el = make(doc, ns, 'c');
          el.setAttribute('r', refName(r, at + k));
          el.setAttribute('s', attr(source, 's'));
          rowEl.insertBefore(el, after || null);
        }
      }
      if (attr(rowEl, 'spans')) rowEl.removeAttribute('spans'); // (a loading hint; dropped rather than left wrong)
    }
  }
  for (const [row, after] of copies) sheetData.insertBefore(row, after || null);
}

/** The style a column's cells take (its `col` element's), by column index. */
function colStyles(root, ns) {
  const cols = kids(kid(root, 'cols', ns), 'col', ns).map((el) => ({ min: Number(attr(el, 'min')) - 1, max: Number(attr(el, 'max')) - 1, s: Number(attr(el, 'style') || 0) }));
  return (c) => cols.find((x) => c >= x.min && c <= x.max)?.s || 0;
}

/** Column widths and formats (`cols`) after a column insert or delete: new columns take the left one's. */
function moveCols(doc, at, n) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const cols = kid(root, 'cols', ns);
  if (!cols) return;
  for (const el of kids(cols, 'col', ns)) {
    const min = Number(attr(el, 'min')) - 1;
    const max = Number(attr(el, 'max')) - 1;
    // (A column element ending just before an insert takes in the new columns: they get its width and format.)
    const span = n > 0 && max === at - 1 ? [min, max + n] : shiftSpan(min, max, at, n);
    if (!span) { cols.removeChild(el); continue; }
    el.setAttribute('min', String(span[0] + 1));
    el.setAttribute('max', String(Math.min(span[1], MAX_COLS - 1) + 1));
  }
  if (!kids(cols, 'col', ns).length) root.removeChild(cols);
}

/** The sheet's own ranges: dimension, merges, array formulas, hyperlinks, conditional formats, data validations,
 *  the filter, ignored errors, protected ranges and the view's selection and frozen pane. */
function moveRanges(doc, axis, at, n) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const one = (el, name) => {
    const r = parseRange(attr(el, name));
    if (!r) return true;
    const next = shiftRange(r, axis, at, n);
    if (next) el.setAttribute(name, rangeName(next));
    return !!next;
  };
  const dim = kid(root, 'dimension', ns);
  if (dim && !one(dim, 'ref')) dim.setAttribute('ref', 'A1');
  const merges = kid(root, 'mergeCells', ns);
  if (merges) {
    for (const el of kids(merges, 'mergeCell', ns)) {
      const r = parseRange(attr(el, 'ref'));
      const next = r && shiftRange(r, axis, at, n);
      if (!next || (next.r0 === next.r1 && next.c0 === next.c1)) merges.removeChild(el);
      else el.setAttribute('ref', rangeName(next));
    }
    const left = kids(merges, 'mergeCell', ns).length;
    if (left) merges.setAttribute('count', String(left)); else root.removeChild(merges);
  }
  for (const f of [...doc.getElementsByTagNameNS(ns, 'f')]) if (attr(f, 't') === 'array' && attr(f, 'ref')) one(f, 'ref');
  const listed = [['hyperlinks', 'hyperlink', 'ref'], ['conditionalFormatting', null, 'sqref'], ['dataValidations', 'dataValidation', 'sqref'],
    ['ignoredErrors', 'ignoredError', 'sqref'], ['protectedRanges', 'protectedRange', 'sqref']];
  for (const [outer, inner, name] of listed) {
    const holders = inner ? kids(kid(root, outer, ns), inner, ns) : kids(root, outer, ns);
    for (const el of holders) {
      const next = name === 'ref' ? (one(el, 'ref') ? attr(el, 'ref') : '') : shiftSqref(attr(el, name), axis, at, n);
      if (next) { if (name !== 'ref') el.setAttribute(name, next); } else el.parentNode.removeChild(el);
    }
    const wrap = inner && kid(root, outer, ns);
    if (wrap && !kids(wrap, inner, ns).length) root.removeChild(wrap);
    else if (wrap && attr(wrap, 'count') != null) wrap.setAttribute('count', String(kids(wrap, inner, ns).length));
  }
  const filter = kid(root, 'autoFilter', ns);
  if (filter && !one(filter, 'ref')) root.removeChild(filter);
  for (const view of kids(kid(root, 'sheetViews', ns), 'sheetView', ns)) {
    for (const sel of kids(view, 'selection', ns)) {
      const cell = parseRef(attr(sel, 'activeCell'));
      const moved = cell && shiftRange({ r0: cell.r, c0: cell.c, r1: cell.r, c1: cell.c }, axis, at, n);
      if (cell) sel.setAttribute('activeCell', moved ? refName(moved.r0, moved.c0) : 'A1');
      if (attr(sel, 'sqref') != null) sel.setAttribute('sqref', shiftSqref(attr(sel, 'sqref'), axis, at, n) || attr(sel, 'activeCell') || 'A1');
    }
  }
}

/**
 * Insert `n` rows (axis 'row') or columns ('col') before index `at` of sheet si (n > 0), or delete -n of them from `at`
 * (n < 0). Throws a SheetError with Excel's words when it can't be done.
 */
export function changeLines(wb, si, axis, at, n) {
  check(wb, si, axis, at, n);
  const info = wb.sheets[si];
  const { doc } = wb._part(si);
  // Formulas first, while every cell is where it was (shared ones on this sheet written out cell by cell).
  mapFormulas(wb, (text, here) => shiftRefs(text, { here, sheet: info.name, axis, at, n }), { name: info.name, target: si, shared: true });
  moveCells(doc, axis, at, n);
  if (axis === 'col') moveCols(doc, at, n);
  moveRanges(doc, axis, at, n);
  // The sheet read again from its moved XML, and everything worked out again.
  wb._parsed.set(info, { doc, model: readSheet(doc, wb.sharedStrings) });
  wb.dirty.add(info.path);
  wb.edited = true;
  wb._dropCalcChain();
  wb.resetCalc();
  wb.calculate();
}
