// Inserting and deleting whole rows and columns, as Excel does it (measured over COM, docs/ROWCOL.md). The sheet's
// rows and cells move in its XML, with its column widths, merged ranges, array formulas and the ranges its hyperlinks,
// conditional formats, data validations, filter and view hold; every formula of the workbook that reads the sheet
// follows (tokens.js shiftRefs), defined names and chart series too. New rows take the format of the row above, new
// columns the format and width of the column to their left. Afterwards the sheet is read again and the workbook worked
// out again. Notes and threaded comments go with their cells, pictures, charts and shapes with whole rows and
// columns, tables grown, narrowed, moved or deleted, pivot tables moved or deleted and pivot caches' sources followed
// (anchors.js). A sheet with slicers or timelines is refused for now: their positions aren't moved yet.
// Data validation rules and sparklines follow as Excel moves them (rulemoves.js).
//
// Inserting and deleting cells (the rest of their columns shifting down or up, or of their rows right or left) is the
// same change limited to a band of columns (or rows): changeLines's `band`, docs/ROWCOL.md.
import { kids, kid, attr } from './xml.js';
import { parseRef, refName, parseRange, rangeName, MAX_ROWS, MAX_COLS } from './refs.js';
import { readSheet } from './sheet.js';
import { pinRefs, forgetRows, setCellXml } from './sheetedit.js';
import { shiftLine, shiftSpan, shiftRefs, bandArea, tokenize } from '../calc/tokens.js';
import { SheetError, mapFormulas } from './sheetops.js';
import { moveNotes, moveDrawings, lineSizes, planTables, applyTables, deadTableRefs, keepTableRow, planPivots, applyPivots, movePivotSources } from './anchors.js';
import { prepareRules, moveExtRanges, extendRules, dropDeadSparklines } from './rulemoves.js';
import { readPivots } from './pivot.js';
import { insertedLines, openJoined } from './outline.js';
import { pxWidth } from './sizes.js';
import { defaultColPx } from './units.js';

export const ROWCOL_MESSAGES = {
  offSheet: 'To prevent possible loss of data, Excel cannot shift nonblank cells off of the worksheet. Select another location in which to insert new cells, or delete data from the end of your worksheet.',
  array: "You can't change part of an array.",
  parts: "Rows and columns can't be inserted or deleted yet on a sheet with slicers or timelines.",
  cellParts: "Cells can't be inserted or deleted yet on a sheet with slicers or timelines.",
  unmerge: 'This operation will cause some merged cells to unmerge. Do you wish to continue?',
};
/** Relationships of a sheet whose parts hold positions this doesn't move yet. */
const ANCHORED = /\/(slicer|timeline)$/;

const make = (doc, ns, name) => doc.createElementNS(ns, doc.documentElement.prefix ? `${doc.documentElement.prefix}:${name}` : name);

/** A range after the change, or null when all of it was deleted (shiftSpan on one axis; bandArea for cells). */
function shiftRange(range, axis, at, n, band = null) {
  if (band) {
    const next = bandArea(range, axis, at, n, band);
    return next === undefined ? range : next && { ...range, r0: next.r0, c0: next.c0, r1: next.r1, c1: next.c1 };
  }
  if (axis === 'row') {
    const s = shiftSpan(range.r0, range.r1, at, n);
    return s && { ...range, r0: s[0], r1: Math.min(s[1], MAX_ROWS - 1) };
  }
  const s = shiftSpan(range.c0, range.c1, at, n);
  return s && { ...range, c0: s[0], c1: Math.min(s[1], MAX_COLS - 1) };
}
/** A space-separated list of ranges (sqref) after the change; '' when none is left. */
const shiftSqref = (text, axis, at, n, band) => (text || '').split(/\s+/).filter(Boolean).map(parseRange).filter(Boolean)
  .map((r) => shiftRange(r, axis, at, n, band)).filter(Boolean).map(rangeName).join(' ');

/** Whether a range has cells that a shift of cells moves or deletes (the band's, from line `at` on) without being
 *  within the band: the shift would split it. */
function splitBy(range, axis, at, band) {
  if (!band) return false;
  const [p0, p1, s1] = axis === 'row' ? [range.c0, range.c1, range.r1] : [range.r0, range.r1, range.c1];
  return s1 >= at && p0 <= band[1] && p1 >= band[0] && !(p0 >= band[0] && p1 <= band[1]);
}

/** Throws Excel's refusal when the change can't be made on sheet si. Returns the merged ranges a shift of cells would
 *  split (Excel unmerges them, asking first). */
function check(wb, si, axis, at, n, band) {
  const info = wb.sheets[si];
  if (wb.pkg.readRels(info.path).some((r) => ANCHORED.test(r.type))) throw new SheetError(band ? ROWCOL_MESSAGES.cellParts : ROWCOL_MESSAGES.parts);
  const model = wb.sheet(si);
  const lines = axis === 'row' ? (a) => [a.r0, a.r1] : (a) => [a.c0, a.c1];
  const inBand = (r, c) => !band || ((axis === 'row' ? c : r) >= band[0] && (axis === 'row' ? c : r) <= band[1]);
  for (const a of model.arrays) {
    if (splitBy(a, axis, at, band)) throw new SheetError(ROWCOL_MESSAGES.array);
    // (An array outside the band isn't moved; one across it was refused above.)
    if (!inBand(a.r0, a.c0)) continue;
    const [a0, a1] = lines(a);
    // (An insert inside an array, or a delete of part of one, would split it.)
    const cut = n > 0 ? a0 < at && at <= a1 : a0 <= at - n - 1 && a1 >= at && !(a0 >= at && a1 <= at - n - 1);
    if (cut) throw new SheetError(ROWCOL_MESSAGES.array);
  }
  if (n > 0) {
    const last = axis === 'row' ? model.maxRow : model.maxCol;
    const limit = axis === 'row' ? MAX_ROWS : MAX_COLS;
    const filled = [...model.cells].some(([r, line]) => [...line].some(([c, cell]) => (axis === 'row' ? r : c) >= limit - n && inBand(r, c) && (cell.v != null || cell.f != null)));
    if (last >= at && filled) throw new SheetError(ROWCOL_MESSAGES.offSheet);
  }
  return model.merges.filter((m) => splitBy(m, axis, at, band));
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
        // (And its outline level, shown and never collapsed: "insert row inside a group", excel-outline.ps1.)
        for (const name of ['s', 'customFormat', 'ht', 'customHeight', 'outlineLevel']) if (attr(source, name) != null) row.setAttribute(name, attr(source, name));
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

/** The cells of a band moved (and those deleted taken out): for axis 'row', the band's columns from row `at` down,
 *  for 'col' its rows from column `at` right. New cells take the format of the cell above (or to the left). */
function moveBandCells(doc, axis, at, n, [b0, b1]) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const sheetData = kid(root, 'sheetData', ns);
  if (!sheetData) return;
  pinRefs(sheetData, ns);
  forgetRows(doc);
  const rowOf = (el) => Number(attr(el, 'r')) - 1;
  const colOf = (el) => parseRef(attr(el, 'r')).c;
  const styled = (el) => el && attr(el, 's') && attr(el, 's') !== '0';
  const rows = new Map(kids(sheetData, 'row', ns).map((el) => [rowOf(el), el]));
  /** The row element of row r, made in its place if there isn't one. */
  const rowEl = (r) => {
    if (rows.has(r)) return rows.get(r);
    const el = make(doc, ns, 'row');
    el.setAttribute('r', String(r + 1));
    sheetData.insertBefore(el, kids(sheetData, 'row', ns).find((x) => rowOf(x) > r) || null);
    rows.set(r, el);
    return el;
  };
  /** Cell element `el` put in row `row` in column order. */
  const place = (row, el) => row.insertBefore(el, kids(row, 'c', ns).find((x) => colOf(x) > colOf(el)) || null);
  const newCell = (r, c, s) => {
    const el = make(doc, ns, 'c');
    el.setAttribute('r', refName(r, c));
    el.setAttribute('s', s);
    return el;
  };
  if (axis === 'row') {
    const above = n > 0 && at > 0 ? kids(sheetData, 'row', ns).find((el) => rowOf(el) === at - 1) : null;
    const sources = above ? kids(above, 'c', ns).filter((c) => colOf(c) >= b0 && colOf(c) <= b1 && styled(c)).map((c) => [colOf(c), attr(c, 's')]) : [];
    const moving = [];
    for (const row of kids(sheetData, 'row', ns)) {
      const r = rowOf(row);
      if (r < at) continue;
      for (const c of kids(row, 'c', ns)) {
        const col = colOf(c);
        if (col < b0 || col > b1) continue;
        row.removeChild(c);
        const next = shiftLine(r, at, n);
        if (next != null) moving.push([next, c]);
      }
      if (attr(row, 'spans')) row.removeAttribute('spans');
    }
    for (const [next, c] of moving) {
      c.setAttribute('r', refName(next, colOf(c)));
      place(rowEl(next), c);
    }
    for (let k = 0; k < n; k++) for (const [col, s] of sources) place(rowEl(at + k), newCell(at + k, col, s));
  } else {
    for (const row of kids(sheetData, 'row', ns)) {
      const r = rowOf(row);
      if (r < b0 || r > b1) continue;
      let source = null;
      for (const c of kids(row, 'c', ns)) {
        const col = colOf(c);
        if (col === at - 1 && n > 0) source = c;
        if (col < at) continue;
        const next = shiftLine(col, at, n);
        if (next == null) row.removeChild(c); else c.setAttribute('r', refName(r, next));
      }
      if (styled(source)) for (let k = 0; k < n; k++) place(row, newCell(r, at + k, attr(source, 's')));
      if (attr(row, 'spans')) row.removeAttribute('spans');
    }
  }
}

/** The style a column's cells take (its `col` element's), by column index. */
function colStyles(root, ns) {
  const cols = kids(kid(root, 'cols', ns), 'col', ns).map((el) => ({ min: Number(attr(el, 'min')) - 1, max: Number(attr(el, 'max')) - 1, s: Number(attr(el, 'style') || 0) }));
  return (c) => cols.find((x) => c >= x.min && c <= x.max)?.s || 0;
}

/** Column widths and formats (`cols`) after a column insert or delete: new columns take the left one's. */
function moveCols(doc, at, n, defaultWidth) {
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
    // (Its outline level too, but shown and not collapsed, as Excel inserts rows: outline.js insertedLines.)
    if (n > 0 && max === at - 1 && (attr(el, 'hidden') === '1' || attr(el, 'collapsed') === '1')) {
      el.parentNode.insertBefore(insertedLines(el, at, n, defaultWidth), el.nextSibling);
      span[1] = max;
    }
    el.setAttribute('min', String(span[0] + 1));
    el.setAttribute('max', String(Math.min(span[1], MAX_COLS - 1) + 1));
  }
  if (!kids(cols, 'col', ns).length) root.removeChild(cols);
}

/** The sheet's own ranges: dimension, merges, array formulas, hyperlinks, conditional formats, data validations,
 *  the filter, ignored errors, protected ranges and the view's selection and frozen pane. */
function moveRanges(doc, axis, at, n, band = null, split = []) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const one = (el, name) => {
    const r = parseRange(attr(el, name));
    if (!r) return true;
    const next = shiftRange(r, axis, at, n, band);
    if (next) el.setAttribute(name, rangeName(next));
    return !!next;
  };
  const dim = kid(root, 'dimension', ns);
  // (A shift of cells only ever grows the used range: an insert as if whole lines, a delete leaving it.)
  const dimRange = dim && parseRange(attr(dim, 'ref'));
  if (band) {
    if (n > 0 && dimRange) {
      const grown = axis === 'row' ? { ...dimRange, r1: Math.min(dimRange.r1 + n, MAX_ROWS - 1) } : { ...dimRange, c1: Math.min(dimRange.c1 + n, MAX_COLS - 1) };
      dim.setAttribute('ref', rangeName(grown));
    }
  } else if (dim && !one(dim, 'ref')) dim.setAttribute('ref', 'A1');
  const merges = kid(root, 'mergeCells', ns);
  if (merges) {
    for (const el of kids(merges, 'mergeCell', ns)) {
      const r = parseRange(attr(el, 'ref'));
      const next = r && !split.some((m) => m.r0 === r.r0 && m.c0 === r.c0 && m.r1 === r.r1 && m.c1 === r.c1) && shiftRange(r, axis, at, n, band);
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
      const next = name === 'ref' ? (one(el, 'ref') ? attr(el, 'ref') : '') : shiftSqref(attr(el, name), axis, at, n, band);
      if (next) { if (name !== 'ref') el.setAttribute(name, next); } else el.parentNode.removeChild(el);
    }
    const wrap = inner && kid(root, outer, ns);
    if (wrap && !kids(wrap, inner, ns).length) root.removeChild(wrap);
    else if (wrap && attr(wrap, 'count') != null) wrap.setAttribute('count', String(kids(wrap, inner, ns).length));
  }
  const filter = kid(root, 'autoFilter', ns);
  if (filter && !one(filter, 'ref')) root.removeChild(filter);
  // (The saved sort, the sheet's or the autoFilter's: its range and each level's follow the lines as Excel moves them,
  // a level whose column is deleted going, and the sortState once its rows or all its levels are (excel-sortstate.ps1,
  // test/sortstate-excel.json, the then-* cases).)
  const kept = kid(root, 'autoFilter', ns);
  for (const st of [kid(root, 'sortState', ns), kept && kid(kept, 'sortState', ns)].filter(Boolean)) {
    const rowsLeft = one(st, 'ref');
    const levels = kids(st, 'sortCondition', ns).filter((sc) => (one(sc, 'ref') ? true : (st.removeChild(sc), false)));
    if (!rowsLeft || !levels.length) st.parentNode.removeChild(st);
  }
  for (const view of kids(kid(root, 'sheetViews', ns), 'sheetView', ns)) {
    for (const sel of kids(view, 'selection', ns)) {
      const cell = parseRef(attr(sel, 'activeCell'));
      const moved = cell && shiftRange({ r0: cell.r, c0: cell.c, r1: cell.r, c1: cell.c }, axis, at, n, band);
      if (cell) sel.setAttribute('activeCell', moved ? refName(moved.r0, moved.c0) : 'A1');
      if (attr(sel, 'sqref') != null) sel.setAttribute('sqref', shiftSqref(attr(sel, 'sqref'), axis, at, n, band) || attr(sel, 'activeCell') || 'A1');
    }
  }
}

/**
 * Insert `n` rows (axis 'row') or columns ('col') before index `at` of sheet si (n > 0), or delete -n of them from `at`
 * (n < 0). Throws a SheetError with Excel's words when it can't be done.
 *
 * With `band` ([first, last] columns for axis 'row', rows for 'col'), only the cells of the band move: Excel's Insert
 * with Shift cells down (axis 'row') or right ('col'), and Delete with Shift cells up or left. A merged range the shift
 * would split is unmerged when `unmerge` is set; otherwise a SheetError with `ask` set says Excel's question.
 */
export function changeLines(wb, si, axis, at, n, band = null, { unmerge = false } = {}) {
  // (All of a table's data rows deleted: Excel keeps the first, anchors.js keepTableRow.)
  [at, n] = keepTableRow(wb, si, axis, at, n, band);
  if (!n) return;
  const split = check(wb, si, axis, at, n, band);
  if (split.length && !unmerge) throw Object.assign(new SheetError(ROWCOL_MESSAGES.unmerge), { ask: true });
  // (Tables first: a change Excel refuses for one is refused before anything moves.)
  const tables = planTables(wb, si, axis, at, n, band, (message) => { throw new SheetError(message); });
  const pivots = planPivots(wb, si, axis, at, n, band, (message) => { throw new SheetError(message); }, readPivots(wb));
  const info = wb.sheets[si];
  const { doc } = wb._part(si);
  // (Data validation rules split where cells shifting cut them, their formulas read from their first cell to be.)
  prepareRules(doc, axis, at, n, band, (range) => shiftRange(range, axis, at, n, band));
  // Formulas first, while every cell is where it was (shared ones on this sheet written out cell by cell).
  mapFormulas(wb, (text, here) => shiftRefs(text, { here, sheet: info.name, axis, at, n, band }), { name: info.name, target: si, shared: true });
  if (band) moveBandCells(doc, axis, at, n, band); else moveCells(doc, axis, at, n);
  const format = wb.sheet(si).format;
  if (axis === 'col' && !band) moveCols(doc, at, n, pxWidth(defaultColPx(format, 7)));
  moveRanges(doc, axis, at, n, band, split);
  moveExtRanges(doc, axis, at, n, band, (text) => shiftSqref(text, axis, at, n, band));
  extendRules(doc, axis, at, n, band);
  // (Sparklines whose data is all deleted go, on any sheet.)
  for (const [other, part] of wb._parsed) if (dropDeadSparklines(part.doc)) wb.dirty.add(other.path);
  // (Lines inserted into a collapsed group open it: its summary's collapsed goes, as Excel's does.)
  if (n > 0 && !band) openJoined(doc, axis, at, n, format);
  // Notes with their cells; pictures, charts and shapes with whole lines (not when cells shift, as in Excel).
  moveNotes(wb, si, (r, c) => {
    const next = shiftRange({ r0: r, c0: c, r1: r, c1: c }, axis, at, n, band);
    return next && { r: next.r0, c: next.c0 };
  });
  if (!band) moveDrawings(wb, si, axis, at, n, lineSizes(wb.sheet(si), axis, at, n));
  // Tables: their ranges, columns and header names; references to what was deleted of them made #REF!.
  const dead = applyTables(wb, si, axis, tables, doc, setCellXml);
  for (const name of new Set(dead.map((d) => d.table))) mapFormulas(wb, (text) => deadTableRefs(text, dead, tokenize), { name });
  // Pivot tables on the sheet moved (or gone); every pivot cache reading the sheet, its source range as a reference.
  applyPivots(wb, si, axis, pivots);
  movePivotSources(wb, info.name, (area) => shiftRange(area, axis, at, n, band));
  // The sheet read again from its moved XML, and everything worked out again.
  wb._parsed.set(info, { doc, model: readSheet(doc, wb.sharedStrings) });
  // (Counted by sheet for what a view holds by position and Excel drops on any insert or delete there: Circle Invalid
  // Data's ovals, excel-refmoves.ps1, all gone after every step on their sheet, kept after one on another sheet.)
  wb.lineChanges.set(info.sheetId, (wb.lineChanges.get(info.sheetId) || 0) + 1);
  wb.dirty.add(info.path);
  wb.edited = true;
  wb._dropCalcChain();
  wb.resetCalc();
  wb.calculate();
}

/** The way Excel shifts cells when it isn't told (Range.Insert and Delete without Shift): a selection taller than it
 *  is wide shifts right (left on a delete), 'col'; any other shifts down (up), 'row'. */
export const shapeShift = (sel) => (sel.r1 - sel.r0 > sel.c1 - sel.c0 ? 'col' : 'row');

/** Insert (dir 1) or delete (-1) the cells of `sel` on sheet si, the rest of their columns shifting down or up (axis
 *  'row') or of their rows right or left ('col'): changeLines with the selection's band. */
export function shiftCells(wb, si, sel, dir, axis, opts = {}) {
  if (axis === 'row') return changeLines(wb, si, 'row', sel.r0, dir * (sel.r1 - sel.r0 + 1), [sel.c0, sel.c1], opts);
  return changeLines(wb, si, 'col', sel.c0, dir * (sel.c1 - sel.c0 + 1), [sel.r0, sel.r1], opts);
}
