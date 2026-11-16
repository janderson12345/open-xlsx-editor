// Column widths, row heights and hidden rows and columns, written as Excel writes them (measured over COM,
// docs/SIZES.md): a width is stored as its pixels over the maximum digit width, cut to 1/256 (typed in characters,
// it is first made whole pixels: 12.5 is 93 pixels, 13.28515625); a height in points with customHeight; a hidden
// column width 0, a hidden row hidden="1". AutoFit Row Height takes the fixed height away (Excel fits such rows when
// it opens the file); AutoFit Column Width stores the width found, with bestFit.
import { kids, kid, attr } from './xml.js';
import { pxToColWidth, defaultColPx } from './units.js';
import { pinRefs, forgetRows } from './sheetedit.js';

/** The maximum digit width of the Normal font in pixels (Calibri 11 and Aptos Narrow 11: 7, as units.js has it). */
const MDW = 7;

/** A width typed in characters as Excel stores it: whole pixels (the characters' width rounded, plus 5 of padding). */
export const typedWidth = (chars) => pxToColWidth(Math.round(chars * MDW) + 5, MDW);
/** A width in pixels (a drag, an AutoFit) as Excel stores it. */
export const pxWidth = (px) => pxToColWidth(px, MDW);

const make = (doc, ns, name) => doc.createElementNS(ns, doc.documentElement.prefix ? `${doc.documentElement.prefix}:${name}` : name);

/**
 * Columns c0 to c1 (0-based) changed: { width } (stored units; 0 hides), { bestFit: true }, { hidden: true/false },
 * { style } (the columns' own format).
 * Existing col elements are split where the range cuts them; a column without one gets one (a hidden one width 0, a
 * shown one the sheet's default width).
 */
export function setColumns(doc, c0, c1, change, format = {}) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  let cols = kid(root, 'cols', ns);
  if (!cols) {
    cols = make(doc, ns, 'cols');
    root.insertBefore(cols, kid(root, 'sheetData', ns));
  }
  const defaultWidth = pxWidth(defaultColPx(format, MDW));
  const pieces = [];
  const piece = (el, min, max) => {
    const p = el.cloneNode(true);
    p.setAttribute('min', String(min + 1));
    p.setAttribute('max', String(max + 1));
    return p;
  };
  let covered = [];
  for (const el of kids(cols, 'col', ns)) {
    const min = Number(attr(el, 'min')) - 1;
    const max = Number(attr(el, 'max')) - 1;
    if (max < c0 || min > c1) { pieces.push(el); continue; }
    if (min < c0) pieces.push(piece(el, min, c0 - 1));
    const inner = piece(el, Math.max(min, c0), Math.min(max, c1));
    pieces.push(inner);
    covered.push(inner);
    if (max > c1) pieces.push(piece(el, c1 + 1, max));
  }
  // (Columns of the range without an element get one.)
  const has = (c) => pieces.some((p) => c >= Number(attr(p, 'min')) - 1 && c <= Number(attr(p, 'max')) - 1);
  for (let c = c0; c <= c1; c++) {
    if (has(c)) continue;
    let end = c;
    while (end < c1 && !has(end + 1)) end++;
    const el = make(doc, ns, 'col');
    el.setAttribute('min', String(c + 1));
    el.setAttribute('max', String(end + 1));
    el.setAttribute('width', String(defaultWidth));
    el.oxeNew = true;
    pieces.push(el);
    covered.push(el);
    c = end;
  }
  covered = covered.filter(Boolean);
  for (const el of covered) {
    if (change.width != null) {
      el.setAttribute('width', String(change.width));
      el.setAttribute('customWidth', '1');
      if (change.width === 0) el.setAttribute('hidden', '1'); else el.removeAttribute('hidden');
      if (change.bestFit) el.setAttribute('bestFit', '1'); else el.removeAttribute('bestFit');
    }
    // (A column's own format: `style`, as Excel writes a formatted column; 0 takes it away.)
    if ('style' in change) { if (change.style) el.setAttribute('style', String(change.style)); else el.removeAttribute('style'); }
    if (change.hidden === true) {
      // (A column hidden without a width of its own is written with width 0, as Excel writes it.)
      if (el.oxeNew) el.setAttribute('width', '0');
      el.setAttribute('hidden', '1');
      el.setAttribute('customWidth', '1');
    } else if (change.hidden === false) {
      el.removeAttribute('hidden');
      // (Shown again with no width of its own: the sheet's default.)
      if (Number(attr(el, 'width') || 0) === 0) el.setAttribute('width', String(defaultWidth));
    }
  }
  pieces.sort((a, b) => Number(attr(a, 'min')) - Number(attr(b, 'min')));
  while (cols.firstChild) cols.removeChild(cols.firstChild);
  for (const p of pieces) cols.appendChild(p);
}

/**
 * Rows r0 to r1 (0-based) changed: { height } (points; 0 hides), { height: null } (AutoFit: the fixed height taken
 * away), { fitted } (AutoFit's height for wrapped text, points, without customHeight), { hidden: true/false }, { style } (the rows' own format). Rows without an element get one when something must be written.
 */
export function setRows(doc, r0, r1, change) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const sheetData = kid(root, 'sheetData', ns);
  pinRefs(sheetData, ns);
  forgetRows(doc);
  const byRow = new Map(kids(sheetData, 'row', ns).map((el) => [Number(attr(el, 'r')) - 1, el]));
  const writes = change.height != null || change.hidden === true || !!change.style;
  for (let r = r0; r <= r1; r++) {
    let el = byRow.get(r);
    if (!el) {
      if (!writes) continue;
      el = make(doc, ns, 'row');
      el.setAttribute('r', String(r + 1));
      const after = kids(sheetData, 'row', ns).find((x) => Number(attr(x, 'r')) - 1 > r);
      sheetData.insertBefore(el, after || null);
      byRow.set(r, el);
    }
    // (A row's own format: `s` with customFormat, as Excel writes a formatted row; 0 takes it away.)
    if ('style' in change) {
      if (change.style) { el.setAttribute('s', String(change.style)); el.setAttribute('customFormat', '1'); } else { el.removeAttribute('s'); el.removeAttribute('customFormat'); }
    }
    if (change.height === 0 || change.hidden === true) el.setAttribute('hidden', '1');
    else if (change.hidden === false || change.height != null) el.removeAttribute('hidden');
    if (change.height > 0) {
      el.setAttribute('ht', String(change.height));
      el.setAttribute('customHeight', '1');
    } else if (change.fitted > 0) {
      // (AutoFit of a row its wrapped text makes taller: the height found, without customHeight, as Excel writes it.)
      el.setAttribute('ht', String(change.fitted));
      el.removeAttribute('customHeight');
    } else if ('height' in change && change.height === null) {
      el.removeAttribute('ht');
      el.removeAttribute('customHeight');
    }
  }
}

/** Columns' and rows' own formats, `cols` [{ c, s }] and `rows` [{ r, s }]: neighbours with one format written as one
 *  element (Workbook setLineStyles). */
export function setLineStyles(doc, { cols = [], rows = [] }, format = {}) {
  const runs = (list, key) => {
    const out = [];
    for (const x of [...list].sort((a, b) => a[key] - b[key])) {
      const last = out[out.length - 1];
      if (last && last.s === x.s && last.b === x[key] - 1) last.b = x[key];
      else out.push({ a: x[key], b: x[key], s: x.s });
    }
    return out;
  };
  for (const { a, b, s } of runs(cols, 'c')) setColumns(doc, a, b, { style: s }, format);
  for (const { a, b, s } of runs(rows, 'r')) setRows(doc, a, b, { style: s });
}

/** CT_SheetFormatPr's attributes in their order (Excel writes them so; others, x14ac:dyDescent, after). */
const FORMAT_ORDER = ['baseColWidth', 'defaultColWidth', 'defaultRowHeight', 'customHeight', 'zeroHeight', 'thickTop', 'thickBottom', 'outlineLevelRow', 'outlineLevelCol'];

/** The sheet's sheetFormatPr given `attrs` ({ name: value, or null to take it away }), made (after sheetViews) when
 *  the sheet has none. */
export function setSheetFormat(doc, attrs) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  let el = kid(root, 'sheetFormatPr', ns);
  if (!el) {
    el = make(doc, ns, 'sheetFormatPr');
    const after = [...root.childNodes].find((n) => n.nodeType === 1 && !['sheetPr', 'dimension', 'sheetViews'].includes(n.localName));
    root.insertBefore(el, after || null);
  }
  const all = new Map([...el.attributes].map((a) => [a.name, a.value]));
  for (const [k, v] of Object.entries(attrs)) { if (v == null) all.delete(k); else all.set(k, String(v)); }
  for (const a of [...el.attributes]) el.removeAttribute(a.name);
  const rank = (k) => { const i = FORMAT_ORDER.indexOf(k); return i < 0 ? FORMAT_ORDER.length : i; };
  for (const [k, v] of [...all].sort((a, b) => rank(a[0]) - rank(b[0]))) el.setAttribute(k, v);
  if (!el.hasAttribute('defaultRowHeight')) el.setAttribute('defaultRowHeight', '15');
}

/** Format > Default Width...: the width, typed in characters, of every column without one of its own, as Excel
 *  writes it (excel-sizes2.ps1): sheetFormatPr's defaultColWidth in whole pixels as a typed width (12: 12.7109375,
 *  4.5: 5.28515625); columns with their own width keep it. */
export function setDefaultWidth(doc, chars) {
  setSheetFormat(doc, { defaultColWidth: typedWidth(chars) });
}

/** A row height given to the whole sheet (every row selected), as Excel writes it: sheetFormatPr's defaultRowHeight
 *  with customHeight, and every row the file has given that height. */
export function setSheetRowHeight(doc, points) {
  setSheetFormat(doc, { defaultRowHeight: points, customHeight: '1' });
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const rows = kids(kid(root, 'sheetData', ns), 'row', ns);
  for (const r of rows.map((el) => Number(attr(el, 'r')) - 1)) setRows(doc, r, r, { height: points });
}

