// A new PivotTable (Insert > PivotTable), written as Excel writes one it has just made (docs/PIVOT.md, measured with
// tools/excel-compare/excel-pivot.ps1: test/pivot-excel.json's "before" parts): a cache over the source (a table by
// its name, otherwise the sheet's range), its records, and the table: its fields on rows and columns and its values,
// in the compact form and PivotStyleLight16. The parts are made bare, then pivot.js's refresh fills the cache, the
// records, the items (sorted, as a new pivot's are) and the cells, so a new pivot and a refreshed one are laid out by
// the same code.
import { NS } from './xml.js';
import { rangeName } from './refs.js';
import { readTables } from './tables.js';
import { readPivots, refreshPivot, PIVOT_MESSAGES } from './pivot.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml';
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const XR = 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision';
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** The value functions offered, as Excel names a value of each ("Sum of Units"; measured). */
export const PIVOT_FUNCTIONS = { sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min' };

/** Why a layout can't be made here (the refresh lays out only these; docs/PIVOT.md), or null. */
export function layoutRefusal({ rows, cols, values }) {
  if (!rows.length && !cols.length) return 'Choose a field for Rows or Columns.';
  if (!values.length) return 'Choose a field for Values.';
  if (cols.length > 1) return 'More than one column field can\'t be laid out here yet.';
  if (cols.length && values.length > 1) return 'Several values with a column field can\'t be laid out here yet.';
  if (new Set([...rows, ...cols]).size < rows.length + cols.length) return 'A field can be in Rows or Columns, not both.';
  return null;
}

/** A GUID in braces, as Excel's xr:uid. */
function guid() {
  const h = () => Math.floor(Math.random() * 16).toString(16).toUpperCase();
  const run = (n) => Array.from({ length: n }, h).join('');
  return `{${run(8)}-${run(4)}-4${run(3)}-${'89AB'[Math.floor(Math.random() * 4)]}${run(3)}-${run(12)}}`;
}

/** The next free name of the form PivotTable1, PivotTable2... among the workbook's pivots. */
function nextName(wb) {
  const used = new Set(readPivots(wb).map((p) => String(p.name).toLowerCase()));
  let n = 1;
  while (used.has(`pivottable${n}`)) n++;
  return `PivotTable${n}`;
}

/** The next free part number k for `${dir}${stem}k.xml`. */
function nextPart(pkg, dir, stem) {
  let k = 1;
  while (pkg.has(`${dir}${stem}${k}.xml`)) k++;
  return k;
}

/** The source's headings: { names } or { error } (an empty one is Excel's refusal). */
function headings(wb, src) {
  const sheet = wb.sheet(src.si);
  const names = [];
  for (let c = src.c0; c <= src.c1; c++) {
    const v = sheet.cells.get(src.r0)?.get(c)?.v;
    if (v == null || v === '') return { error: PIVOT_MESSAGES.fieldName };
    names.push(String(v));
  }
  return { names };
}

/** The fields of a source range { si, r0, c0, r1, c1 }: { names } or { error } (Excel's words), for the box. */
export function pivotSourceFields(wb, src) {
  if (!src || src.r1 <= src.r0) return { error: PIVOT_MESSAGES.source };
  return headings(wb, src);
}

/**
 * Make a PivotTable of `src` ({ si, r0, c0, r1, c1 }: the headings in its first row) at cell `at` ({ si, r, c }),
 * with `layout` = { rows, cols: field indexes in order, values: [{ field, fn (PIVOT_FUNCTIONS) }] }. Writes the parts,
 * then refreshes it: returns refreshPivot's { changes } (the cells, for the caller's history), and { name, path }; or
 * { error } with nothing changed. `now` (a Date) is the cache's refreshed date.
 */
export function createPivot(wb, src, at, layout, { now = new Date(), name = null } = {}) {
  const fields = pivotSourceFields(wb, src);
  if (fields.error) return fields;
  const why = layoutRefusal(layout);
  if (why) return { error: why };
  const { names } = fields;
  const pkg = wb.pkg;
  const pivotName = name || nextName(wb);
  const k = nextPart(pkg, 'xl/pivotTables/', 'pivotTable');
  const c = nextPart(pkg, 'xl/pivotCache/', 'pivotCacheDefinition');
  const tablePath = `xl/pivotTables/pivotTable${k}.xml`;
  const cachePath = `xl/pivotCache/pivotCacheDefinition${c}.xml`;
  const recordsPath = `xl/pivotCache/pivotCacheRecords${c}.xml`;
  const root = wb.doc.documentElement;
  const cacheIds = [...wb.doc.getElementsByTagNameNS(root.namespaceURI, 'pivotCache')].map((e) => Number(e.getAttribute('cacheId')) || 0);
  const cacheId = Math.max(0, ...cacheIds) + 1;
  // (A source that is a table's header and data rows is named by the table, as Excel names it.)
  const table = readTables(wb).find((t) => t.si === src.si && t.r0 === src.r0 && t.c0 === src.c0 && t.r1 - t.totals === src.r1 && t.c1 === src.c1);
  const source = table ? `<worksheetSource name="${esc(table.name)}"/>` : `<worksheetSource ref="${rangeName(src)}" sheet="${esc(wb.sheets[src.si].name)}"/>`;
  // (Excel's refreshed date: days since 1900 in local time, with the time of day.)
  const serial = (now.getTime() - now.getTimezoneOffset() * 60000) / 86400000 + 25569;
  pkg.setText(recordsPath, `${DECL}<pivotCacheRecords xmlns="${MAIN}" xmlns:r="${REL}" xmlns:mc="${MC}" mc:Ignorable="xr" xmlns:xr="${XR}" count="0"></pivotCacheRecords>`);
  const recId = pkg.addRel(`${REL}/pivotCacheRecords`, `pivotCacheRecords${c}.xml`, false, cachePath);
  pkg.setText(cachePath, `${DECL}<pivotCacheDefinition xmlns="${MAIN}" xmlns:r="${REL}" xmlns:mc="${MC}" mc:Ignorable="xr" xmlns:xr="${XR}" r:id="${recId}" refreshedDate="${serial}" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="0" xr:uid="${guid()}">` +
    `<cacheSource type="worksheet">${source}</cacheSource><cacheFields count="${names.length}">${names.map((n) => `<cacheField name="${esc(n)}" numFmtId="0"><sharedItems/></cacheField>`).join('')}</cacheFields>` +
    '<extLst><ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:pivotCacheDefinition/></ext></extLst></pivotCacheDefinition>');
  // The table: each field on its axis or a value; Excel adds the values' own column field (-2) when there are several.
  const { rows, cols, values } = layout;
  const axis = (i) => (rows.includes(i) ? ' axis="axisRow"' : cols.includes(i) ? ' axis="axisCol"' : '');
  const pivotFields = names.map((_, i) => `<pivotField${axis(i)}${values.some((v) => v.field === i) ? ' dataField="1"' : ''} showAll="0"/>`).join('');
  const list = (tag, xs) => (xs.length ? `<${tag} count="${xs.length}">${xs.map((x) => `<field x="${x}"/>`).join('')}</${tag}>` : '');
  const colList = values.length > 1 ? [...cols, -2] : cols;
  const dataFields = values.map((v) => `<dataField name="${esc(`${PIVOT_FUNCTIONS[v.fn]} of ${names[v.field]}`)}" fld="${v.field}"${v.fn === 'sum' ? '' : ` subtotal="${v.fn}"`} baseField="0" baseItem="0"/>`).join('');
  pkg.setText(tablePath, `${DECL}<pivotTableDefinition xmlns="${MAIN}" xmlns:mc="${MC}" mc:Ignorable="xr" xmlns:xr="${XR}" xr:uid="${guid()}" name="${esc(pivotName)}" cacheId="${cacheId}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">` +
    `<location ref="${rangeName({ r0: at.r, c0: at.c, r1: at.r, c1: at.c })}" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="${names.length}">${pivotFields}</pivotFields>` +
    `${list('rowFields', rows)}${list('colFields', colList)}<dataFields count="${values.length}">${dataFields}</dataFields>` +
    '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>' +
    '<extLst><ext uri="{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:pivotTableDefinition hideValuesRow="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"/></ext>' +
    '<ext uri="{747A6164-185A-40DC-8AA5-F01512510D54}" xmlns:xpdl="http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout"><xpdl:pivotTableDefinition16/></ext></extLst></pivotTableDefinition>');
  // Wired in as Excel wires them: the sheet to its table, the table to its cache, the workbook to the cache by its id.
  pkg.addRel(`${REL}/pivotCacheDefinition`, `../pivotCache/pivotCacheDefinition${c}.xml`, false, tablePath);
  pkg.addRel(`${REL}/pivotTable`, `../pivotTables/pivotTable${k}.xml`, false, wb.sheets[at.si].path);
  const cacheRel = pkg.addRel(`${REL}/pivotCacheDefinition`, `pivotCache/pivotCacheDefinition${c}.xml`);
  pkg.addOverride(tablePath, `${CT}.pivotTable+xml`);
  pkg.addOverride(cachePath, `${CT}.pivotCacheDefinition+xml`);
  pkg.addOverride(recordsPath, `${CT}.pivotCacheRecords+xml`);
  addCacheEntry(wb, cacheId, cacheRel);
  const done = refreshPivot(wb, { si: at.si, path: tablePath, name: pivotName, cachePath, recordsPath }, { fresh: true });
  if (done.error) return done;
  return { ...done, name: pivotName, path: tablePath };
}

/** The workbook's <pivotCache cacheId r:id/> (in <pivotCaches>, after calcPr, as the schema orders it). */
function addCacheEntry(wb, cacheId, rId) {
  const root = wb.doc.documentElement;
  const ns = root.namespaceURI;
  let caches = [...root.childNodes].find((n) => n.localName === 'pivotCaches');
  if (!caches) {
    caches = wb.doc.createElementNS(ns, 'pivotCaches');
    const ORDER = ['customWorkbookViews', 'oleSize', 'calcPr', 'definedNames', 'externalReferences', 'functionGroups', 'sheets'];
    const after = ORDER.map((name) => [...root.childNodes].find((n) => n.localName === name)).find(Boolean);
    root.insertBefore(caches, after ? after.nextSibling : null);
  }
  const one = wb.doc.createElementNS(ns, 'pivotCache');
  one.setAttribute('cacheId', String(cacheId));
  one.setAttributeNS(NS.r, 'r:id', rId);
  caches.appendChild(one);
  wb.bookDirty = true;
}
