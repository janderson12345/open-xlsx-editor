// AutoFilter as Excel does it (measured over COM with Range.AutoFilter, and the XML Excel then saves: docs/FILTER.md).
// Turned on from one cell it takes the data around it. A column filtered by values keeps the rows whose shown text is
// one of them, case aside, and blanks when asked; the rest are hidden. The sheet's autoFilter element and the hidden
// _xlnm._FilterDatabase name are written as Excel writes them; a kind of filter not read here (custom, top 10,
// colour, dates) is kept as it is.
import { kids, kid, attr } from './xml.js';
import { rangeName } from './refs.js';
import { formatCodeOf } from './styles.js';
import { cellText } from './numfmt.js';
import { sortArea, compareText, SORT_MESSAGES } from './sortops.js';
import { AFTER_MERGES } from './sheetedit.js';

const make = (doc, ns, name) => doc.createElementNS(ns, doc.documentElement.prefix ? `${doc.documentElement.prefix}:${name}` : name);
/** The worksheet elements that follow autoFilter (CT_Worksheet's order), for placing it when it is new. */
const AFTER_FILTER = ['sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', ...AFTER_MERGES];

/** What a filter would cover for selection `sel`: the sheet's filter when it has one, else what Sort would take (one
 *  cell: the data around it). { area } or { error } with Excel's words. */
export function filterArea(sheet, sel) {
  if (sheet.autoFilter) return { area: sheet.autoFilter };
  const area = sortArea(sheet, sel);
  return area ? { area } : { error: SORT_MESSAGES.empty };
}

const blank = (cell) => !cell || ((cell.f == null || cell.f === '') && (cell.v == null || cell.v === ''));
/** A cell's text as the sheet shows it, which is what a filter matches: '' for a blank. */
export function shownText(wb, cell) {
  if (blank(cell)) return '';
  if (cell.t === 'e') return String(cell.v);
  if (typeof cell.v === 'boolean') return cell.v ? 'TRUE' : 'FALSE';
  if (typeof cell.v === 'number') return cellText(cell, Infinity, formatCodeOf(wb.styles, cell.s || 0), wb.date1904).trim();
  return String(cell.v);
}

/** Whether row r passes column c's filter (one read here). */
function passes(wb, sheet, r, c, col) {
  const cell = sheet.cells.get(r)?.get(c);
  if (blank(cell)) return col.blank;
  const text = shownText(wb, cell).toLowerCase();
  return col.values.some((v) => v.toLowerCase() === text);
}

/**
 * The rows of `filter` ({ r0, c0, r1, c1, columns }) to hide: those failing a column filtered by values. Where a filter
 * of a kind not read here is set, a row passing the others keeps the hidden state it has.
 */
export function hiddenRows(wb, si, filter) {
  const sheet = wb.sheet(si);
  const plain = [...filter.columns].filter(([, col]) => !col.other);
  const other = plain.length < filter.columns.size;
  const out = new Set();
  for (let r = filter.r0 + 1; r <= filter.r1; r++) {
    const fails = plain.some(([c, col]) => !passes(wb, sheet, r, c, col));
    if (fails || (other && sheet.rows.get(r)?.hidden)) out.add(r);
  }
  return out;
}

/** The order the filter list shows values in: numbers (by value), then text as Sort orders it, then the rest. */
function listOrder(a, b) {
  if (a.number != null || b.number != null) return a.number == null ? 1 : b.number == null ? -1 : a.number - b.number;
  return compareText(a.text, b.text);
}

/**
 * Column c's filter list: [{ text, checked }] of the values in the rows the other columns let through (one entry for
 * texts alike but for case, spelled as first found), in Excel's order, and { blanks: checked or null when there are
 * none }.
 */
export function filterChoices(wb, si, filter, c) {
  const sheet = wb.sheet(si);
  const others = [...filter.columns].filter(([k, col]) => k !== c && !col.other);
  const mine = filter.columns.get(c);
  const seen = new Map();
  let blanks = false;
  for (let r = filter.r0 + 1; r <= filter.r1; r++) {
    if (others.some(([k, col]) => !passes(wb, sheet, r, k, col))) continue;
    const cell = sheet.cells.get(r)?.get(c);
    if (blank(cell)) { blanks = true; continue; }
    const text = shownText(wb, cell);
    const key = text.toLowerCase();
    if (!seen.has(key)) seen.set(key, { text, number: typeof cell.v === 'number' ? cell.v : null });
  }
  const values = [...seen.values()].sort(listOrder).map(({ text }) => ({
    text, checked: !mine || mine.other || mine.values.some((v) => v.toLowerCase() === text.toLowerCase()),
  }));
  return { values, blanks: blanks ? !mine || mine.other || mine.blank : null };
}

/** Write `filter` as the sheet's autoFilter element (null: taken out). A column of a kind not read here keeps its XML;
 *  values are written in order, as Excel writes them. */
export function setAutoFilterXml(doc, filter) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  let el = kid(root, 'autoFilter', ns);
  if (!filter) { if (el) root.removeChild(el); return; }
  if (!el) {
    el = make(doc, ns, 'autoFilter');
    const next = [...root.childNodes].find((n) => n.nodeType === 1 && AFTER_FILTER.includes(n.localName));
    root.insertBefore(el, next || null);
  }
  el.setAttribute('ref', rangeName(filter));
  const old = new Map(kids(el, 'filterColumn', ns).map((fc) => [Number(attr(fc, 'colId')), fc]));
  for (const fc of old.values()) el.removeChild(fc);
  const after = [...el.childNodes].find((n) => n.nodeType === 1) || null;
  for (const [c, col] of [...filter.columns].sort((a, b) => a[0] - b[0])) {
    const id = c - filter.c0;
    if (col.other) { if (old.has(id)) el.insertBefore(old.get(id), after); continue; }
    const fc = make(doc, ns, 'filterColumn');
    fc.setAttribute('colId', String(id));
    const filters = make(doc, ns, 'filters');
    if (col.blank) filters.setAttribute('blank', '1');
    for (const v of [...col.values].sort(compareText)) {
      const f = make(doc, ns, 'filter');
      f.setAttribute('val', v);
      filters.appendChild(f);
    }
    fc.appendChild(filters);
    el.insertBefore(fc, after);
  }
}
