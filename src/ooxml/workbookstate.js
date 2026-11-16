// A workbook's state taken and put back (Undo's snapshots of the whole workbook), and a formula's result written into
// its cell: Workbook's snapshot, restore and writeResult (workbook.js), here to keep that file short.
import { kids, kid, attr, serializeXml, parseXml } from './xml.js';
import { readSheet } from './sheet.js';
import { withEl } from './sheetops.js';
import { setResultXml } from './sheetedit.js';
import { RICH_ERRORS } from './spillxml.js';
import { readStyles } from './styles.js';
import { forgetFormats } from './stylesedit.js';

/** The workbook as it is now, to go back to (restore): the parts, the sheets read so far and the workbook part; with
 *  `styles`, the styles part too (a cell style changed in place, cellstyledefine.js, which cells' formats don't undo). */
export function snapshotOf(wb, { styles = false } = {}) {
  return {
    ...(styles && wb.stylesDoc ? { styles: serializeXml(wb.stylesDoc) } : {}),
    files: { ...wb.pkg.files }, rels: wb.pkg.rels, book: serializeXml(wb.doc), dirty: new Set(wb.dirty), bookDirty: wb.bookDirty,
    // (The sheets as they were, names, order and all, and the parsed ones by their ids: sheets added, renamed, moved
    // or deleted since are put back too.)
    sheets: wb.sheets.map((info) => ({ ...info })),
    parsed: [...wb._parsed].map(([info, part]) => [info.sheetId, serializeXml(part.doc)]),
  };
}

/** Back to a snapshot: its parts, its sheets (their list read again) and its defined names. Cell formats added
 *  since stay (no cell of the snapshot uses them). */
export function restoreTo(wb, snap) {
  wb.pkg.files = { ...snap.files };
  wb.pkg.rels = snap.rels;
  wb.doc = parseXml(snap.book);
  const root = wb.doc.documentElement;
  const els = kids(kid(root, 'sheets', wb.ns), 'sheet', wb.ns);
  wb.sheets = snap.sheets.map((info) => withEl({ ...info }, els.find((el) => Number(attr(el, 'sheetId')) === info.sheetId) || null));
  wb._parsed = new Map(snap.parsed.map(([id, text]) => {
    const doc = parseXml(text);
    return [wb.sheets.find((s) => s.sheetId === id), { doc, model: readSheet(doc, wb.sharedStrings) }];
  }).filter(([info]) => info));
  wb.definedNames = kids(kid(root, 'definedNames', wb.ns), 'definedName', wb.ns).map((el) => withEl({
    name: attr(el, 'name'),
    sheet: attr(el, 'localSheetId') == null ? null : Number(attr(el, 'localSheetId')),
    text: el.textContent,
  }, el));
  wb.dirty = new Set(snap.dirty);
  if (snap.styles) {
    wb.stylesDoc = parseXml(snap.styles);
    Object.assign(wb.styles, readStyles(wb.stylesDoc), { palette: wb.styles.palette });
    forgetFormats(wb);
    wb.dirty.add(wb.stylesPath);
  }
  wb._dynamicCm = null;
  wb.bookDirty = true;
  wb.edited = true;
  wb.resetCalc();
}

/** Write a formula's result into a cell (from the engine); nothing changes if it is the value the cell has. */
export function writeResultTo(wb, i, r, c, value) {
  const part = wb._part(i);
  const model = part.model;
  const cell = model.cells.get(r)?.get(c);
  const t = typeof value === 'number' ? 'n' : typeof value === 'boolean' ? 'b' : typeof value === 'string' ? 'str' : 'e';
  const v = t === 'e' ? value.code : value;
  if (cell && cell.t === t && cell.v === v) return;
  // (#SPILL! and #CALC! are rich values in Excel's file: #VALUE! there, worked out again when Excel opens it.)
  if (t === 'e' && RICH_ERRORS.has(v)) wb._incomplete = true;
  setResultXml(part.doc, r, c, t === 'e' ? (RICH_ERRORS.has(v) ? { code: '#VALUE!' } : value) : v);
  if (cell) { cell.t = t; cell.v = v; }
  else {
    if (!model.cells.has(r)) model.cells.set(r, new Map());
    model.cells.get(r).set(c, { t, v, s: 0, f: null });
  }
  model.maxRow = Math.max(model.maxRow, r);
  model.maxCol = Math.max(model.maxCol, c);
  wb.dirty.add(wb.sheets[i].path);
}
