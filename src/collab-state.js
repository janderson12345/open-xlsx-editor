// What a co-edited workbook's sheets hold, in the terms collab.js shares (a digest), and the writers that make a sheet
// hold it: each cell's contents and format, each row's and column's own attributes (height, width, hidden, format,
// outline), the sheet's defaults. Everything else in the workbook (merges, rules, tables, names, drawings, the styles
// part's other lists, other parts) is summed up in a fingerprint: a change there isn't shared piece by piece, and
// collab.js shares the whole workbook instead.
import { frameOf } from './ooxml/sheetcells.js';
import { kids, kid, attr, serializeXml } from './ooxml/xml.js';
import { pinRefs, forgetRows } from './ooxml/sheetedit.js';
import { setColumns, setSheetFormat } from './ooxml/sizes.js';
import { MAX_COLS } from './ooxml/refs.js';

const attrsOf = (el, skip) => {
  const out = {};
  for (const a of [...el.attributes]) if (!skip.has(a.name)) out[a.name] = a.value;
  return out;
};
const ROW_SKIP = new Set(['r', 'spans']);
const COL_SKIP = new Set(['min', 'max']);
export const sameAttrs = (a, b) => {
  const ka = Object.keys(a || {});
  return ka.length === Object.keys(b || {}).length && ka.every((k) => a[k] === b[k]);
};

/** A cell's contents as the room keeps them: { f, legacy } for a formula, { v } for a value; null for a cell whose
 *  contents aren't its own (part of an array formula's range, or spilled into). */
export function contentsOf(wb, si, model, r, c, cell) {
  if (model.arrays.some((a) => r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1 && (r !== a.r0 || c !== a.c0))) return null;
  if (cell && cell.f != null && cell.f !== '') return { f: cell.f, ...(cell.dynamic || cell.array ? {} : { legacy: true }) };
  if (wb.isSpilled(si, r, c)) return null;
  return { v: cell && cell.f == null ? cell.v ?? null : null };
}
export const sameContents = (a, b) => (a?.f ?? null) === (b?.f ?? null) && !!a?.legacy === !!b?.legacy && (a?.v ?? null) === (b?.v ?? null);

/** Sheet si's digest: cells Map('r,c' -> { c (contents, or null), s }), rows Map(r -> attributes), cols Map(c ->
 *  attributes), format (sheetFormatPr's attributes). */
export function digestSheet(wb, si) {
  const { doc, model } = wb._part(si);
  const cells = new Map();
  for (const [r, row] of model.cells) {
    for (const [c, cell] of row) {
      const contents = contentsOf(wb, si, model, r, c, cell);
      if ((contents?.v == null && contents?.f == null) && !cell.s) continue;
      cells.set(`${r},${c}`, { c: contents, s: cell.s || 0 });
    }
  }
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const rows = new Map();
  let r = -1;
  for (const el of kids(kid(root, 'sheetData', ns), 'row', ns)) {
    r = attr(el, 'r') ? Number(attr(el, 'r')) - 1 : r + 1;
    const a = attrsOf(el, ROW_SKIP);
    if (Object.keys(a).length) rows.set(r, a);
  }
  const cols = new Map();
  for (const el of kids(kid(root, 'cols', ns), 'col', ns)) {
    const a = attrsOf(el, COL_SKIP);
    const max = Math.min(Number(attr(el, 'max')), MAX_COLS);
    for (let c = Number(attr(el, 'min')) - 1; c < max; c++) cols.set(c, a);
  }
  const fmt = kid(root, 'sheetFormatPr', ns);
  return { cells, rows, cols, format: fmt ? attrsOf(fmt, new Set()) : {} };
}

/** What a sheet part holds besides its digest (and its view, each editor's own): its XML, for comparing. */
const SHEET_SKIP = new Set(['sheetData', 'cols', 'sheetFormatPr', 'dimension', 'sheetViews']);
export function sheetPrint(wb, si) {
  const root = frameOf(wb._part(si)).documentElement;
  let out = '';
  for (const n of root.childNodes) {
    if (n.nodeType !== 1 || SHEET_SKIP.has(n.localName)) continue;
    if (n.localName === 'sheetPr') {
      // (The tab colour goes with the sheet list.)
      const pr = n.cloneNode(true);
      for (const t of [...pr.childNodes]) if (t.nodeType === 1 && t.localName === 'tabColor') pr.removeChild(t);
      out += serializeXml(pr);
    } else out += serializeXml(n);
  }
  return out;
}

const BOOK_SKIP = new Set(['sheets', 'bookViews']);
const STYLES_SKIP = new Set(['numFmts', 'fonts', 'fills', 'borders', 'cellXfs', 'colors']);
/** What the workbook holds besides its sheets' digests and its sheet list: the workbook part (names, calculation),
 *  the styles part's other lists (cell styles, conditional formats' dxfs, table styles), and every other part. */
export function bookPrint(wb) {
  let text = '';
  for (const n of wb.doc.documentElement.childNodes) {
    if (n.nodeType !== 1 || BOOK_SKIP.has(n.localName)) continue;
    if (n.localName === 'calcPr') {
      // (fullCalcOnLoad: written by any save of a workbook with a formula not worked out here.)
      const pr = n.cloneNode(true);
      pr.removeAttribute('fullCalcOnLoad');
      text += serializeXml(pr);
    } else text += serializeXml(n);
  }
  if (wb.stylesDoc) for (const n of wb.stylesDoc.documentElement.childNodes) if (n.nodeType === 1 && !STYLES_SKIP.has(n.localName)) text += serializeXml(n);
  const skip = new Set([...wb.sheets.map((s) => s.path), wb.stylesPath, wb.pkg.mainPath, '[Content_Types].xml']);
  const parts = new Map();
  for (const [path, bytes] of Object.entries(wb.pkg.files)) {
    // (Parts every edit may rewrite: shared strings, the calculation chain, the workbook's relationships, properties.)
    if (skip.has(path) || /sharedStrings|calcChain|^docProps\/|_rels\/workbook\.xml\.rels$/.test(path)) continue;
    parts.set(path, bytes);
  }
  return { text, parts };
}
export const samePrint = (a, b) => a.text === b.text && a.parts.size === b.parts.size && [...a.parts].every(([k, v]) => b.parts.get(k) === v);

/** Rows of sheet si given exactly the attributes of `list` ([{ r, attrs }]; {} for none): the row element made when
 *  it has none, taken away when it has nothing left. The sheet read again once. */
export function writeRows(wb, si, list) {
  if (!list.length) return;
  const part = wb._part(si);
  const root = part.doc.documentElement;
  const ns = root.namespaceURI;
  const sheetData = kid(root, 'sheetData', ns);
  pinRefs(sheetData, ns);
  forgetRows(part.doc);
  const byRow = new Map(kids(sheetData, 'row', ns).map((el) => [Number(attr(el, 'r')) - 1, el]));
  for (const { r, attrs } of list) {
    let el = byRow.get(r);
    const empty = !Object.keys(attrs).length;
    if (!el) {
      if (empty) continue;
      el = part.doc.createElementNS(ns, root.prefix ? `${root.prefix}:row` : 'row');
      el.setAttribute('r', String(r + 1));
      const after = kids(sheetData, 'row', ns).find((x) => Number(attr(x, 'r')) - 1 > r);
      sheetData.insertBefore(el, after || null);
      byRow.set(r, el);
    }
    for (const a of [...el.attributes]) if (!ROW_SKIP.has(a.name)) el.removeAttribute(a.name);
    for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, v);
    if (empty && !kids(el, 'c', ns).length) { sheetData.removeChild(el); byRow.delete(r); }
  }
  wb._reread(si, part);
  wb.resetCalc();
}

/** Columns of sheet si given exactly the attributes of `list` ([{ c, attrs }]; {} for none): their col elements split
 *  out (sizes.js), set, and joined again with neighbours that have the same. The sheet read again once. */
export function writeCols(wb, si, list) {
  if (!list.length) return;
  const part = wb._part(si);
  const root = part.doc.documentElement;
  const ns = root.namespaceURI;
  for (const { c, attrs } of list) {
    setColumns(part.doc, c, c, {}, part.model.format);
    const cols = kid(root, 'cols', ns);
    const el = kids(cols, 'col', ns).find((x) => Number(attr(x, 'min')) === c + 1);
    for (const a of [...el.attributes]) if (!COL_SKIP.has(a.name)) el.removeAttribute(a.name);
    for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, v);
    if (!Object.keys(attrs).length) cols.removeChild(el);
  }
  const cols = kid(root, 'cols', ns);
  let last = null;
  for (const el of kids(cols, 'col', ns)) {
    if (last && Number(attr(last, 'max')) + 1 === Number(attr(el, 'min')) && sameAttrs(attrsOf(last, COL_SKIP), attrsOf(el, COL_SKIP))) {
      last.setAttribute('max', attr(el, 'max'));
      cols.removeChild(el);
    } else last = el;
  }
  if (!kids(cols, 'col', ns).length) root.removeChild(cols);
  wb._reread(si, part);
  wb.resetCalc();
}

/** Sheet si's sheetFormatPr given exactly `attrs`. */
export function writeFormat(wb, si, attrs) {
  const part = wb._part(si);
  const el = kid(part.doc.documentElement, 'sheetFormatPr', part.doc.documentElement.namespaceURI);
  const now = el ? attrsOf(el, new Set()) : {};
  setSheetFormat(part.doc, { ...Object.fromEntries(Object.keys(now).map((k) => [k, null])), ...attrs });
  wb._reread(si, part);
  wb.resetCalc();
}
