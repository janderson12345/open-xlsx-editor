// Tables made and changed as Excel makes and changes them (docs/TABLES.md; measured over COM,
// tools/excel-compare/excel-tables.ps1 on table-cases.json, test/tablemake-excel.json): Insert > Table and Format as
// Table (ListObjects.Add), the style and Table Style Options, the total row and each column's total, the header row,
// the filter buttons, the name and Resize Table (Convert to Range: tableconvert.js). Each takes the workbook and a table of readTables
// (tables.js) and changes the table part, the sheet and the workbook's formulas.
import { serializeXml, kids, kid, NS } from './xml.js';
import { readTables } from './tables.js';
import { rangeName } from './refs.js';
import { validName } from './names.js';
import { mapFormulas } from './sheetops.js';
import { tokenize } from '../calc/tokens.js';

const CT_TABLE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';
const REL_TABLE = `${NS.r}/table`;
const SHEET_AFTER = ['extLst'];

/** Excel's words when it refuses. */
export const TABLE_EDIT_MESSAGES = {
  overlap: 'A table cannot overlap a range that contains a PivotTable report, query results, protected cells or another table.',
  name: "The name that you entered is not valid.\n\nReasons for this can include:\n- The name does not begin with a letter or an underscore\n- The name contains a space or other invalid characters\n- The name conflicts with an Excel built-in name or the name of another object in the workbook",
  resize: 'The headers must remain in the same row and the resulting table range must overlap the original table range.',
  merged: 'This operation requires the merged cells to be identically sized.',
};

/** A column's total: totalsRowFunction and the SUBTOTAL function number Excel writes for it (calc-each, calc-rest:
 *  Count is COUNTA's 103, Count Numbers COUNT's 102, StdDev 107, Var 110). */
export const TOTAL_FUNCTIONS = { average: 101, countNums: 102, count: 103, max: 104, min: 105, stdDev: 107, sum: 109, var: 110 };

const tableOf = (wb, name) => readTables(wb).find((t) => t.name.toLowerCase() === name.toLowerCase());
const overlaps = (a, b) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;

/** A column name in a table reference: ' before [, ], # and ' (tables.js reads them back). */
export const structName = (name) => String(name).replace(/(['[\]#])/g, "'$1");

/** The table at cell (r, c) of sheet si, or null. */
export function tableAt(wb, si, r, c) {
  return readTables(wb).find((t) => t.si === si && r >= t.r0 && r <= t.r1 && c >= t.c0 && c <= t.c1) || null;
}

/** A cell's text as a header takes it: a string as it is, a number or a formula's result as text, a blank as null. */
function headerText(cell) {
  if (!cell || cell.v == null || cell.v === '') return null;
  if (typeof cell.v === 'boolean') return cell.v ? 'TRUE' : 'FALSE';
  return String(cell.v);
}

/** Header names made unique as Excel makes them (add-heads: a blank is Column1, Column2...; a name repeated gets 2, 3,
 *  ...; capitals aside). */
export function uniqueHeaders(texts) {
  const used = new Set();
  const out = [];
  const take = (n) => { used.add(n.toLowerCase()); out.push(n); };
  for (const t of texts) {
    if (t == null) { out.push(null); continue; }
    let n = t;
    for (let k = 2; used.has(n.toLowerCase()); k++) n = `${t}${k}`;
    take(n);
  }
  let k = 1;
  return out.map((n) => {
    if (n != null) return n;
    while (used.has(`column${k}`)) k++;
    used.add(`column${k}`);
    return `Column${k}`;
  });
}

/** The next free "Table<n>" (a name no table or defined name has). */
function freeName(wb) {
  const taken = new Set([...readTables(wb).map((t) => t.name.toLowerCase()), ...wb.definedNames.map((d) => d.name.toLowerCase())]);
  let n = 1;
  while (taken.has(`table${n}`)) n++;
  return `Table${n}`;
}

function tableXml({ id, name, ref, headers, totals, filterRef, columns, style }) {
  const cols = columns.map((c, i) => `<tableColumn id="${i + 1}" name="${esc(c)}"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<table xmlns="${NS.x}" id="${id}" name="${esc(name)}" displayName="${esc(name)}" ref="${ref}"${headers ? '' : ' headerRowCount="0"'}${totals ? ' totalsRowCount="1"' : ' totalsRowShown="0"'}>` +
    (filterRef ? `<autoFilter ref="${filterRef}"/>` : '') + `<tableColumns count="${columns.length}">${cols}</tableColumns>` +
    `<tableStyleInfo${style ? ` name="${esc(style)}"` : ''} showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`;
}
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** The sheet's <tableParts> with a part for relationship `rid` added, made before extLst when there is none. */
function addTablePart(doc, rid) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  let wrap = kid(root, 'tableParts', ns);
  if (!wrap) {
    wrap = doc.createElementNS(ns, root.prefix ? `${root.prefix}:tableParts` : 'tableParts');
    root.insertBefore(wrap, [...root.childNodes].find((n) => n.nodeType === 1 && SHEET_AFTER.includes(n.localName)) || null);
  }
  const el = doc.createElementNS(ns, root.prefix ? `${root.prefix}:tablePart` : 'tablePart');
  el.setAttributeNS(NS.r, 'r:id', rid);
  wrap.appendChild(el);
  wrap.setAttribute('count', String(kids(wrap, 'tablePart', ns).length));
}

/**
 * Insert > Table, Ctrl+T and Format as Table: a table over `area` ({ r0, c0, r1, c1 }) of sheet si, as ListObjects.Add
 * makes it. With `headers`, its first row's texts are the column names (made unique; blanks named ColumnN, the header
 * cells written as text); a single row gets an empty data row below, the cells under it shifted down (add-one-row).
 * Without, a header row of Column1... is put above, the cells there shifted down (add-noheaders). The style is
 * `style`, else the workbook's default table style, else TableStyleMedium2. A sheet's AutoFilter over the cells goes
 * (the table has its own). Returns the table's name, or { error } in Excel's words.
 */
export function createTable(wb, si, area, { headers = true, style } = {}) {
  let { r0, c0, r1, c1 } = area;
  const tables = readTables(wb);
  if (tables.some((t) => t.si === si && overlaps(t, area))) return { error: TABLE_EDIT_MESSAGES.overlap };
  const sheet = wb.sheet(si);
  if ((sheet.merges || []).some((m) => overlaps(m, area))) return { error: TABLE_EDIT_MESSAGES.merged };
  if (sheet.autoFilter && overlaps(sheet.autoFilter, area)) wb.setAutoFilter(si, null);
  let insertRow = false;
  if (!headers) {
    wb.changeLines(si, 'row', r0, 1, [c0, c1]);
    r1 += 1;
  } else if (r0 === r1) {
    // (A header alone: an empty data row below it, the cells there moved down, and the table marked so.)
    wb.changeLines(si, 'row', r0 + 1, 1, [c0, c1]);
    r1 += 1;
    insertRow = true;
  }
  const line = wb.sheet(si).cells.get(r0);
  const texts = [];
  for (let c = c0; c <= c1; c++) texts.push(headers ? headerText(line?.get(c)) : null);
  const columns = uniqueHeaders(texts);
  wb.apply(si, columns.map((name, k) => ({ r: r0, c: c0 + k, value: name })));
  const name = freeName(wb);
  const id = Math.max(0, ...tables.map((t) => t.id)) + 1;
  let n = 1;
  while (wb.pkg.has(`xl/tables/table${n}.xml`)) n++;
  const path = `xl/tables/table${n}.xml`;
  const ref = rangeName({ r0, c0, r1, c1 });
  const chosen = style === undefined ? wb.styles.defaultTableStyle || 'TableStyleMedium2' : style;
  let xml = tableXml({ id, name, ref, headers: true, totals: false, filterRef: ref, columns, style: chosen });
  if (insertRow) xml = xml.replace(` ref="${ref}"`, ` ref="${ref}" insertRow="1" insertRowShift="1"`);
  wb.pkg.setText(path, xml);
  wb.pkg.addOverride(path, CT_TABLE);
  const rid = wb.pkg.addRel(REL_TABLE, `../tables/table${n}.xml`, false, wb.sheets[si].path);
  const part = wb._part(si);
  addTablePart(part.doc, rid);
  wb._reread(si, part);
  wb.resetCalc();
  return name;
}

/** The table element's attributes in the order Excel writes them (headers-off: headerRowCount before totalsRowShown). */
const TABLE_ATTRS = ['id', 'name', 'displayName', 'ref', 'tableType', 'insertRow', 'insertRowShift', 'headerRowCount', 'totalsRowCount', 'totalsRowShown'];
function orderAttrs(root) {
  // (mc:Ignorable first and xr:uid after id, where Excel's own file has them; anything else unknown after the rest.)
  const rank = (a) => (a.name === 'mc:Ignorable' ? -1 : a.localName === 'uid' ? 0.5 : a.namespaceURI ? 99 : TABLE_ATTRS.indexOf(a.name) < 0 ? 99 : TABLE_ATTRS.indexOf(a.name));
  const all = [...root.attributes].filter((a) => !a.name.startsWith('xmlns'));
  const sorted = all.map((a, i) => ({ a, i })).sort((x, y) => rank(x.a) - rank(y.a) || x.i - y.i).map(({ a }) => [a.namespaceURI, a.name, a.value]);
  for (const a of all) root.removeAttributeNode(a);
  for (const [ns, name, value] of sorted) { if (ns) root.setAttributeNS(ns, name, value); else root.setAttribute(name, value); }
}

/** Change table `t`'s part through `fn(root, doc)`, then read again what depends on it. */
function changePart(wb, t, fn) {
  const doc = wb.pkg.xml(t.path);
  const root = doc.documentElement;
  const out = fn(root, doc);
  orderAttrs(root);
  wb.pkg.setText(t.path, serializeXml(doc));
  wb.dirty.add(t.path);
  wb.edited = true;
  wb.resetCalc();
  return out;
}
const child = (doc, root, name) => kid(root, name, root.namespaceURI) || root.appendChild(doc.createElementNS(root.namespaceURI, name));

/** Table Design's style gallery and Format as Table on a table: `name` (null: None, written without a name). */
export function setTableStyle(wb, t, name) {
  changePart(wb, t, (root, doc) => {
    const info = kid(root, 'tableStyleInfo', root.namespaceURI) || child(doc, root, 'tableStyleInfo');
    if (name) info.setAttribute('name', name); else info.removeAttribute('name');
    if (!info.hasAttribute('showFirstColumn')) for (const [k, v] of [['showFirstColumn', 0], ['showLastColumn', 0], ['showRowStripes', 1], ['showColumnStripes', 0]]) info.setAttribute(k, String(v));
  });
}

const OPTION_ATTR = { first: 'showFirstColumn', last: 'showLastColumn', rows: 'showRowStripes', cols: 'showColumnStripes' };

/** A Table Style Option on or off: First Column, Last Column, Banded Rows, Banded Columns (options), and Filter Button
 *  (no-filter-button: each column's filterColumn hiddenButton="1" in the autoFilter, taken off again when on). */
export function setTableOption(wb, t, key, on) {
  if (key === 'filterButton') return setFilterButton(wb, t, on);
  if (key === 'filter') return setTableFilter(wb, t, on);
  changePart(wb, t, (root, doc) => {
    const info = kid(root, 'tableStyleInfo', root.namespaceURI) || child(doc, root, 'tableStyleInfo');
    info.setAttribute(OPTION_ATTR[key], on ? '1' : '0');
  });
}

/** Data > Filter in a table (no-autofilter): its autoFilter taken off, or put back over its header and data rows. */
function setTableFilter(wb, t, on) {
  changePart(wb, t, (root, doc) => {
    const ns = root.namespaceURI;
    const f = kid(root, 'autoFilter', ns);
    if (!on && f) root.removeChild(f);
    if (on && !f && t.headers) {
      const el = doc.createElementNS(ns, 'autoFilter');
      el.setAttribute('ref', rangeName({ r0: t.r0, c0: t.c0, r1: t.r1 - t.totals, c1: t.c1 }));
      root.insertBefore(el, kid(root, 'sortState', ns) || kid(root, 'tableColumns', ns));
    }
  });
}

function setFilterButton(wb, t, on) {
  changePart(wb, t, (root, doc) => {
    const ns = root.namespaceURI;
    let filter = kid(root, 'autoFilter', ns);
    if (!filter && !on) return;
    if (!filter && t.headers) {
      filter = doc.createElementNS(ns, 'autoFilter');
      filter.setAttribute('ref', rangeName({ r0: t.r0, c0: t.c0, r1: t.r1 - t.totals, c1: t.c1 }));
      root.insertBefore(filter, kid(root, 'sortState', ns) || kid(root, 'tableColumns', ns));
    }
    if (!filter) return;
    for (let k = 0; k <= t.c1 - t.c0; k++) {
      let fc = kids(filter, 'filterColumn', ns).find((x) => Number(x.getAttribute('colId')) === k);
      if (on) {
        if (fc) { fc.removeAttribute('hiddenButton'); if (fc.attributes.length <= 1 && !fc.firstElementChild) filter.removeChild(fc); }
        continue;
      }
      if (!fc) {
        fc = doc.createElementNS(ns, 'filterColumn');
        fc.setAttribute('colId', String(k));
        const after = kids(filter, 'filterColumn', ns).find((x) => Number(x.getAttribute('colId')) > k);
        filter.insertBefore(fc, after || kid(filter, 'sortState', ns));
      }
      fc.setAttribute('hiddenButton', '1');
    }
  });
}

/** The total cell's formula for column `k` of `t` and function `fn`. */
const totalFormula = (t, k, fn) => `SUBTOTAL(${TOTAL_FUNCTIONS[fn]},${t.name}[${structName(t.columns[k])}])`;

/** Whether the data in column c of `t` is numbers (Excel's default total there: Sum, else Count; totals,
 *  totals-text-last). */
function numeric(wb, t, c) {
  let any = false;
  for (let r = t.r0 + t.headers; r <= t.r1 - t.totals; r++) {
    const cell = wb.sheet(t.si).cells.get(r)?.get(c);
    if (cell?.v == null || cell.v === '') continue;
    if (typeof cell.v !== 'number') return false;
    any = true;
  }
  return any;
}

/**
 * Total Row on or off. On: a row below the data, the cells there moved down when any is filled (totals-below-filled),
 * each column's label or SUBTOTAL written from the part (kept while it was off, totals-off-on), or, the first time,
 * Excel's own: "Total" in the first column and Sum (numbers) or Count in the last. Off: the row's cells emptied, each
 * column's total kept in the part.
 */
export function setTotalsRow(wb, t, on) {
  if (!!t.totals === on) return;
  if (on) {
    const below = t.r1 + 1;
    const line = wb.sheet(t.si).cells.get(below);
    let filled = false;
    for (let c = t.c0; c <= t.c1; c++) if (line?.get(c)?.v != null) filled = true;
    if (filled) wb.changeLines(t.si, 'row', below, 1, [t.c0, t.c1]);
    const fresh = t.totalsRow.every((x) => !x.fn && x.label == null);
    const list = t.totalsRow.map((x, k) => {
      if (!fresh) return x;
      if (k === 0 && t.columns.length > 1) return { fn: null, label: 'Total' };
      if (k === t.columns.length - 1) return { fn: numeric(wb, t, t.c1) ? 'sum' : 'count', label: null };
      return x;
    });
    changePart(wb, t, (root) => {
      root.setAttribute('ref', rangeName({ r0: t.r0, c0: t.c0, r1: below, c1: t.c1 }));
      root.removeAttribute('totalsRowShown');
      root.setAttribute('totalsRowCount', '1');
      writeTotals(root, list);
    });
    const now = { ...t, r1: below, totals: 1 };
    wb.apply(t.si, list.map((x, k) => totalCell(now, k, x)).filter(Boolean));
    return;
  }
  wb.apply(t.si, t.columns.map((_, k) => ({ r: t.r1, c: t.c0 + k, value: null })));
  changePart(wb, t, (root) => {
    root.setAttribute('ref', rangeName({ r0: t.r0, c0: t.c0, r1: t.r1 - 1, c1: t.c1 }));
    root.removeAttribute('totalsRowCount');
  });
}

/** The change writing column k's total into the total row of `t`, or null for none. */
function totalCell(t, k, x) {
  const at = { r: t.r1, c: t.c0 + k };
  if (x.fn && TOTAL_FUNCTIONS[x.fn]) return { ...at, formula: totalFormula(t, k, x.fn) };
  if (x.fn === 'custom' && x.formula) return { ...at, formula: x.formula };
  if (x.label != null) return { ...at, value: x.label };
  return null;
}

/** Each tableColumn's totalsRowFunction and totalsRowLabel as `list` gives them. */
function writeTotals(root, list) {
  const cols = kids(kid(root, 'tableColumns', root.namespaceURI), 'tableColumn', root.namespaceURI);
  cols.forEach((el, k) => {
    const x = list[k] || {};
    el.removeAttribute('totalsRowFunction');
    el.removeAttribute('totalsRowLabel');
    if (x.label != null) el.setAttribute('totalsRowLabel', x.label);
    if (x.fn) el.setAttribute('totalsRowFunction', x.fn);
    if (x.fn !== 'custom') for (const f of kids(el, 'totalsRowFormula', root.namespaceURI)) el.removeChild(f);
  });
}

/** The total row's drop-down in column k: `fn` one of TOTAL_FUNCTIONS' or 'none' (calc-each, calc-rest: a function
 *  in the first column takes the place of its label). */
export function setTotalFunction(wb, t, k, fn) {
  if (!t.totals) return;
  const list = t.totalsRow.map((x) => ({ ...x }));
  list[k] = fn === 'none' ? { fn: null, label: null } : { fn, label: null };
  changePart(wb, t, (root) => writeTotals(root, list));
  wb.apply(t.si, [totalCell(t, k, list[k]) || { r: t.r1, c: t.c0 + k, value: null }]);
}

/** Header Row on or off (headers-off, headers-off-on): off, the header cells emptied and the table without its header
 *  and autoFilter; on, the column names written back in the row above the data. */
export function setHeaderRow(wb, t, on) {
  if (!!t.headers === on) return;
  if (!on) {
    wb.apply(t.si, t.columns.map((_, k) => ({ r: t.r0, c: t.c0 + k, value: null })));
    changePart(wb, t, (root) => {
      const ns = root.namespaceURI;
      root.setAttribute('ref', rangeName({ r0: t.r0 + 1, c0: t.c0, r1: t.r1, c1: t.c1 }));
      root.setAttribute('headerRowCount', '0');
      const f = kid(root, 'autoFilter', ns);
      if (f) root.removeChild(f);
    });
    return;
  }
  if (t.r0 === 0) return;
  const r0 = t.r0 - 1;
  wb.apply(t.si, t.columns.map((name, k) => ({ r: r0, c: t.c0 + k, value: name })));
  changePart(wb, t, (root, doc) => {
    const ns = root.namespaceURI;
    root.setAttribute('ref', rangeName({ r0, c0: t.c0, r1: t.r1, c1: t.c1 }));
    root.removeAttribute('headerRowCount');
    const f = doc.createElementNS(ns, 'autoFilter');
    f.setAttribute('ref', rangeName({ r0, c0: t.c0, r1: t.r1 - t.totals, c1: t.c1 }));
    root.insertBefore(f, kid(root, 'sortState', ns) || kid(root, 'tableColumns', ns));
  });
}

/** Table Design's Table Name: `name` checked as Excel checks it (a defined name's rules, not one a name or another
 *  table has), and every formula's references to the table renamed (rename: SUM(Table1[Qty]) is SUM(Sales[Qty])).
 *  Returns { error } when refused. */
export function renameTable(wb, t, name) {
  const n = String(name ?? '').trim();
  if (n.toLowerCase() === t.name.toLowerCase() && n === t.name) return null;
  const clash = readTables(wb).some((x) => x !== t && x.name.toLowerCase() === n.toLowerCase() && x.path !== t.path)
    || wb.definedNames.some((d) => d.name.toLowerCase() === n.toLowerCase());
  if (!validName(n) || clash) return { error: TABLE_EDIT_MESSAGES.name };
  changePart(wb, t, (root) => { root.setAttribute('name', n); root.setAttribute('displayName', n); });
  const was = t.name.toLowerCase();
  mapFormulas(wb, (text) => mapStruct(text, (tk) => (tk.table && tk.table.toLowerCase() === was ? `${n}[${tk.spec}]` : null)), { name: t.name });
  wb.resetCalc();
  return null;
}

/** A formula with its table references given by `fn(token)` (null: kept as it is). */
export function mapStruct(text, fn) {
  let toks;
  try { toks = tokenize(text); } catch { return text; }
  let out = '';
  let last = 0;
  for (const tk of toks) {
    if (tk.type !== 'struct') continue;
    const next = fn(tk);
    if (next == null) continue;
    out += text.slice(last, tk.start) + next;
    last = tk.end;
  }
  return out + text.slice(last);
}

/**
 * Resize Table: to `area`, whose header row must be the table's and which must overlap it (Excel's refusal otherwise).
 * Columns kept by place, new ones named from their header cells (blanks ColumnN), columns cut off gone from the part;
 * a total row moves to the new last row, taken out first with the cells under it moving up (resize-totals).
 */
export function resizeTable(wb, t, area) {
  if (area.r0 !== t.r0 || !overlaps(area, t) || area.r1 <= area.r0) return { error: TABLE_EDIT_MESSAGES.resize };
  if (readTables(wb).some((x) => x.si === t.si && x.path !== t.path && overlaps(x, area))) return { error: TABLE_EDIT_MESSAGES.overlap };
  const totals = t.totals;
  if (totals) {
    // (The total row taken out, the cells under it moved up, and made again at the new last row: resize-totals.)
    setTotalsRow(wb, t, false);
    wb.changeLines(t.si, 'row', t.r1, -1, [t.c0, t.c1]);
  }
  const line = wb.sheet(t.si).cells.get(area.r0);
  const kept = [];
  const texts = [];
  for (let c = area.c0; c <= area.c1; c++) {
    const k = c - t.c0;
    const old = k >= 0 && k < t.columns.length ? t.columns[k] : null;
    kept.push(old != null);
    texts.push(old ?? (t.headers ? headerText(line?.get(c)) : null));
  }
  const names = uniqueHeaders(texts);
  const dataEnd = totals ? area.r1 - 1 : area.r1;
  changePart(wb, t, (root, doc) => {
    const ns = root.namespaceURI;
    root.setAttribute('ref', rangeName({ r0: area.r0, c0: area.c0, r1: dataEnd, c1: area.c1 }));
    const f = kid(root, 'autoFilter', ns);
    if (f) f.setAttribute('ref', rangeName({ r0: area.r0, c0: area.c0, r1: dataEnd, c1: area.c1 }));
    const wrap = kid(root, 'tableColumns', ns);
    const old = kids(wrap, 'tableColumn', ns);
    let id = Math.max(0, ...old.map((el) => Number(el.getAttribute('id')) || 0));
    for (const el of old) wrap.removeChild(el);
    names.forEach((nm, j) => {
      const k = area.c0 + j - t.c0;
      const el = kept[j] ? old[k] : doc.createElementNS(ns, 'tableColumn');
      if (!kept[j]) { el.setAttribute('id', String(++id)); el.setAttribute('name', nm); }
      wrap.appendChild(el);
    });
    wrap.setAttribute('count', String(names.length));
  });
  if (t.headers) wb.apply(t.si, names.map((nm, j) => (kept[j] ? null : { r: area.r0, c: area.c0 + j, value: nm })).filter(Boolean));
  if (totals) setTotalsRow(wb, tableOf(wb, t.name), true);
  return null;
}

/**
 * What typing in or by a table does to it, after cell (r, c) of sheet si was entered (rename-header, calc-custom,
 * type-below): a header cell names its column (a blank one ColumnN, a repeat made unique); a total cell typed holds the
 * column's label (text) or custom total (a formula: totalsRowFunction custom and its totalsRowFormula); a cell typed
 * just below a table without a total row, in its columns, makes the table a row longer (Excel's AutoExpansion).
 */
export function tableEntry(wb, si, r, c) {
  const cell = wb.sheet(si).cells.get(r)?.get(c);
  const tables = readTables(wb).filter((t) => t.si === si);
  const t = tables.find((x) => r >= x.r0 && r <= x.r1 && c >= x.c0 && c <= x.c1);
  if (t && t.headers && r === t.r0) {
    const k = c - t.c0;
    const texts = t.columns.map((nm, j) => (j === k ? headerText(cell) : nm));
    const others = texts.filter((_, j) => j !== k).map((x) => x.toLowerCase());
    let name = texts[k];
    if (name == null) { let n = 1; while (others.includes(`column${n}`)) n++; name = `Column${n}`; }
    else { const base = name; for (let n = 2; others.includes(name.toLowerCase()); n++) name = `${base}${n}`; }
    changePart(wb, t, (root) => kids(kid(root, 'tableColumns', root.namespaceURI), 'tableColumn', root.namespaceURI)[k].setAttribute('name', name));
    if (headerText(cell) !== name) wb.apply(si, [{ r, c, value: name }]);
    return;
  }
  if (t && t.totals && r === t.r1) {
    const k = c - t.c0;
    const list = t.totalsRow.map((x) => ({ ...x }));
    const formula = cell?.f;
    list[k] = formula != null ? { fn: 'custom', label: null, formula } : cell?.v != null && cell.v !== '' ? { fn: null, label: String(cell.v) } : { fn: null, label: null };
    changePart(wb, t, (root, doc) => {
      writeTotals(root, list);
      if (formula != null) {
        const el = kids(kid(root, 'tableColumns', root.namespaceURI), 'tableColumn', root.namespaceURI)[k];
        const f = doc.createElementNS(root.namespaceURI, 'totalsRowFormula');
        f.textContent = formula;
        el.appendChild(f);
      }
    });
    return;
  }
  const below = tables.find((x) => !x.totals && r === x.r1 + 1 && c >= x.c0 && c <= x.c1);
  if (below && cell?.v != null) {
    changePart(wb, below, (root) => {
      const ns = root.namespaceURI;
      root.setAttribute('ref', rangeName({ r0: below.r0, c0: below.c0, r1: r, c1: below.c1 }));
      const f = kid(root, 'autoFilter', ns);
      if (f) f.setAttribute('ref', rangeName({ r0: below.r0, c0: below.c0, r1: r, c1: below.c1 }));
    });
  }
}
