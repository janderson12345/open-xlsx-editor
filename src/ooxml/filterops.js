// AutoFilter as Excel does it (measured over COM with Range.AutoFilter, and the XML Excel then saves: docs/FILTER.md).
// Turned on from one cell it takes the data around it. A column filtered by values keeps the rows whose shown text is
// one of them, case aside, and blanks when asked; the rest are hidden. Text, number, top 10, average, date and colour
// filters are filtermore.js's. The sheet's autoFilter element and the hidden _xlnm._FilterDatabase name are written as
// Excel writes them; a kind of filter not read here (icons) is kept as it is.
import { kids, kid, attr } from './xml.js';
import { rangeName, colName } from './refs.js';
import { setRows } from './sizes.js';
import { readSheet } from './sheet.js';
import { withEl } from './sheetops.js';
import { sheetPrefix } from '../calc/tokens.js';
import { formatCodeOf } from './styles.js';
import { cellText } from './numfmt.js';
import { sortArea, compareText, SORT_MESSAGES } from './sortops.js';
import { AFTER_MERGES } from './sheetedit.js';
import { kid as kidOf, parseXml } from './xml.js';
import { columnTest, filterColumnXml, dxfFillColor, colorDxfXml, todaySerial, inDateGroups, dateFiltersXml } from './filtermore.js';
import { isDateFormat, dateParts } from './format.js';
import { xfOf } from './styles.js';
import { resolveColor } from './theme.js';

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

/** Whether row r passes column c's list of values (and date groups: a date in one of them passes, as Excel's). */
function passes(wb, sheet, r, c, col) {
  const cell = sheet.cells.get(r)?.get(c);
  if (blank(cell)) return col.blank;
  if (col.dates?.length && typeof cell.v === 'number' && isDateFormat(formatCodeOf(wb.styles, cell.s || 0)) && inDateGroups(dateParts(cell.v, wb.date1904), col.dates)) return true;
  const text = shownText(wb, cell).toLowerCase();
  return col.values.some((v) => v.toLowerCase() === text);
}

/** A test of the rows of `filter` for its column c's filter (`today` a serial date, for the date filters). */
export function rowTest(wb, sheet, filter, c, col, today = todaySerial()) {
  if (!col.custom && !col.top10 && !col.dynamic && !col.color) return (r) => passes(wb, sheet, r, c, col);
  const cell = (r) => sheet.cells.get(r)?.get(c);
  const rows = [];
  for (let r = filter.r0 + 1; r <= filter.r1; r++) rows.push(r);
  const color = (x) => resolveColor(x, wb.theme, wb.styles.palette);
  return columnTest(col, {
    value: (r) => (blank(cell(r)) ? null : cell(r).v),
    text: (r) => shownText(wb, cell(r)),
    rows,
    cellLook: (r) => {
      const xf = xfOf(wb.styles, cell(r)?.s || 0);
      const fill = wb.styles.fills[xf.fillId];
      const font = wb.styles.fonts[xf.fontId];
      return { fill: fill?.pattern === 'solid' ? color(fill.fg) : null, font: font?.color && !font.color.auto && !(font.color.theme === 1 && !font.color.tint) ? color(font.color) : null };
    },
    dxfColor: (id) => dxfFillColor(wb, id),
    today,
  });
}

/**
 * The rows of `filter` ({ r0, c0, r1, c1, columns }) to hide: those failing a column filtered by values. Where a filter
 * of a kind not read here is set, a row passing the others keeps the hidden state it has.
 */
export function hiddenRows(wb, si, filter, today = todaySerial()) {
  const sheet = wb.sheet(si);
  const plain = [...filter.columns].filter(([, col]) => !col.other).map(([c, col]) => rowTest(wb, sheet, filter, c, col, today));
  const other = plain.length < filter.columns.size;
  const out = new Set();
  for (let r = filter.r0 + 1; r <= filter.r1; r++) {
    const fails = plain.some((test) => !test(r));
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
  const others = [...filter.columns].filter(([k, col]) => k !== c && !col.other).map(([k, col]) => rowTest(wb, sheet, filter, k, col));
  const mine = filter.columns.get(c);
  const listed = mine && !mine.other && !mine.custom && !mine.top10 && !mine.dynamic && !mine.color;
  const seen = new Map();
  let blanks = false;
  for (let r = filter.r0 + 1; r <= filter.r1; r++) {
    if (others.some((test) => !test(r))) continue;
    const cell = sheet.cells.get(r)?.get(c);
    if (blank(cell)) { blanks = true; continue; }
    const text = shownText(wb, cell);
    const key = text.toLowerCase();
    const date = typeof cell.v === 'number' && isDateFormat(formatCodeOf(wb.styles, cell.s || 0)) ? dateParts(cell.v, wb.date1904) : null;
    if (!seen.has(key)) seen.set(key, { text, number: typeof cell.v === 'number' ? cell.v : null, date: date && { y: date.y, m: date.m, d: date.d } });
  }
  // (A date's entry carries its parts, for the list's years, months and days, ticked by the date groups too.)
  const values = [...seen.values()].sort(listOrder).map(({ text, date }) => ({
    text, ...(date ? { date } : {}),
    checked: !listed || mine.values.some((v) => v.toLowerCase() === text.toLowerCase()) || !!(date && mine.dates && inDateGroups(date, mine.dates)),
  }));
  return { values, blanks: blanks ? !listed || mine.blank : null };
}

/** Write `filter` as the sheet's autoFilter element (null: taken out). A column of a kind not read here keeps its XML;
 *  values are written in order, as Excel writes them; the other filters as filtermore.js writes them (`wb` and `si`
 *  for their values, the average, and a colour filter's dxf). */
export function setAutoFilterXml(doc, filter, wb = null, si = null, today = todaySerial()) {
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
    if (wb && (col.custom || col.top10 || col.dynamic || col.color)) {
      const sheet = wb.sheet(si);
      const values = [];
      for (let r = filter.r0 + 1; r <= filter.r1; r++) { const x = sheet.cells.get(r)?.get(c); if (!blank(x)) values.push(x.v); }
      const nums = values.filter((v) => typeof v === 'number');
      const average = nums.reduce((s, v) => s + v, 0) / (nums.length || 1);
      if (col.color && col.color.rgb && col.color.dxfId == null) col.color.dxfId = addDxf(wb, colorDxfXml(col.color.rgb));
      const inner = filterColumnXml(col, { values, average, today, dxfId: col.color?.dxfId ?? 0 });
      const fc = doc.importNode(parseXml(`<w xmlns="${ns}"><filterColumn colId="${id}">${inner}</filterColumn></w>`).documentElement.firstChild, true);
      el.insertBefore(fc, after);
      continue;
    }
    // (Dates by year, month or day: as Excel 365 writes them, in the filterColumn's extLst: filtermore.js.)
    if (col.dates?.length) {
      const fc = doc.importNode(parseXml(`<w xmlns="${ns}"><filterColumn colId="${id}">${dateFiltersXml(col)}</filterColumn></w>`).documentElement.firstChild, true);
      el.insertBefore(fc, after);
      continue;
    }
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

/** A dxf added to the workbook's styles (its dxfs made when there are none); returns its index. */
export function addDxf(wb, xml) {
  const sroot = wb.stylesDoc.documentElement;
  const sns = sroot.namespaceURI;
  let dxfs = kidOf(sroot, 'dxfs', sns);
  if (!dxfs) {
    dxfs = wb.stylesDoc.createElementNS(sns, sroot.prefix ? `${sroot.prefix}:dxfs` : 'dxfs');
    const after = [...sroot.childNodes].find((n) => n.nodeType === 1 && ['tableStyles', 'colors', 'extLst'].includes(n.localName));
    sroot.insertBefore(dxfs, after || null);
  }
  dxfs.appendChild(wb.stylesDoc.importNode(parseXml(`<w xmlns="${sns}">${xml}</w>`).documentElement.firstChild, true));
  const n = kids(dxfs, 'dxf', sns).length;
  dxfs.setAttribute('count', String(n));
  wb.dirty.add(wb.stylesPath);
  return n - 1;
}

/** Sheet i's AutoFilter on (`filter`) or off (null): the autoFilter element, rows hidden or shown, the name. */
export function applyAutoFilter(wb, i, filter) {
  const part = wb._part(i);
  const was = part.model.autoFilter;
  setAutoFilterXml(part.doc, filter, wb, i, wb.today);
  const area = filter || was;
  if (area) {
    const hide = filter ? hiddenRows(wb, i, filter, wb.today) : new Set();
    // (Runs of rows alike written at once.)
    for (let r = area.r0 + 1; r <= area.r1;) {
      const on = hide.has(r);
      let end = r;
      while (end + 1 <= area.r1 && hide.has(end + 1) === on) end++;
      setRows(part.doc, r, end, { hidden: on });
      r = end + 1;
    }
  }
  if (filter) filterName(wb, i, filter);
  wb._parsed.set(wb.sheets[i], { doc: part.doc, model: wb._withMdw(readSheet(part.doc, wb.sharedStrings)) });
  wb.dirty.add(wb.sheets[i].path);
  wb.edited = true;
  wb.resetCalc();
}

/** The hidden name Excel keeps a sheet's filter range in: _xlnm._FilterDatabase, local to the sheet. */
export function filterName(wb, i, area, name = '_xlnm._FilterDatabase', hidden = true) {
  const abs = (r, c) => `$${colName(c)}$${r + 1}`;
  const text = `${sheetPrefix(wb.sheets[i].name)}${abs(area.r0, area.c0)}:${abs(area.r1, area.c1)}`;
  const found = wb.definedNames.find((d) => d.name === name && d.sheet === i);
  if (found) {
    if (found.text !== text) { found.text = text; found.el.textContent = text; wb.bookDirty = true; }
    return;
  }
  const root = wb.doc.documentElement;
  let names = kid(root, 'definedNames', wb.ns);
  if (!names) {
    names = wb.doc.createElementNS(wb.ns, 'definedNames');
    // (CT_Workbook's order: definedNames comes after sheets, functionGroups and externalReferences.)
    const AFTER = ['calcPr', 'oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'];
    root.insertBefore(names, [...root.childNodes].find((n) => n.nodeType === 1 && AFTER.includes(n.localName)) || null);
  }
  const el = wb.doc.createElementNS(wb.ns, 'definedName');
  el.setAttribute('name', name);
  el.setAttribute('localSheetId', String(i));
  if (hidden) el.setAttribute('hidden', '1');
  el.textContent = text;
  names.appendChild(el);
  wb.definedNames.push(withEl({ name, sheet: i, text }, el));
  wb.bookDirty = true;
}
