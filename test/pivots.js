// For the pivot tests: a workbook with one of test/pivot-excel.json's cases, its data on a sheet "Src" from A1 and the
// pivot table Excel made of it (its three parts as Excel wrote them) on a sheet "Pivot", its cells as Excel showed them.
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseRange } from '../src/ooxml/refs.js';
import { NS } from '../src/ooxml/xml.js';

export const PIVOT = JSON.parse(fs.readFileSync('test/pivot-excel.json', 'utf8'));
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml';

export function pivotBook(k, which = 'before') {
  const c = PIVOT.cases[k];
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const src = wb.addSheet(wb.sheets.length, 'Src');
  const piv = wb.addSheet(wb.sheets.length, 'Pivot');
  const data = which === 'after' ? applyChange(PIVOT.data, PIVOT.change) : PIVOT.data;
  wb.apply(src, data.flatMap((row, r) => row.map((value, col) => ({ r, c: col, value }))));
  const parts = c.parts[which];
  wb.apply(piv, c[which].cells.map(([r, col, value]) => ({ r, c: col, value })));
  const pkg = wb.pkg;
  pkg.setText('xl/pivotCache/pivotCacheRecords1.xml', parts.records);
  pkg.setText('xl/pivotCache/pivotCacheDefinition1.xml', parts.cache.replace(/ sheet="Data"/, ' sheet="Src"'));
  pkg.setText('xl/pivotTables/pivotTable1.xml', parts.table);
  // (The cache definition names its records by r:id; the ids here are new ones, so that attribute follows.)
  const recId = pkg.addRel(`${REL}/pivotCacheRecords`, 'pivotCacheRecords1.xml', false, 'xl/pivotCache/pivotCacheDefinition1.xml');
  pkg.setText('xl/pivotCache/pivotCacheDefinition1.xml', pkg.text('xl/pivotCache/pivotCacheDefinition1.xml').replace(/ r:id="[^"]*"/, ` r:id="${recId}"`));
  pkg.addRel(`${REL}/pivotCacheDefinition`, '../pivotCache/pivotCacheDefinition1.xml', false, 'xl/pivotTables/pivotTable1.xml');
  pkg.addRel(`${REL}/pivotTable`, '../pivotTables/pivotTable1.xml', false, wb.sheets[piv].path);
  const cacheRel = pkg.addRel(`${REL}/pivotCacheDefinition`, 'pivotCache/pivotCacheDefinition1.xml');
  pkg.addOverride('xl/pivotTables/pivotTable1.xml', `${CT}.pivotTable+xml`);
  pkg.addOverride('xl/pivotCache/pivotCacheDefinition1.xml', `${CT}.pivotCacheDefinition+xml`);
  pkg.addOverride('xl/pivotCache/pivotCacheRecords1.xml', `${CT}.pivotCacheRecords+xml`);
  // The workbook lists its caches by the id the table names (after calcPr, as the schema orders them).
  const cacheId = /cacheId="(\d+)"/.exec(parts.table)[1];
  const root = wb.doc.documentElement;
  const caches = wb.doc.createElementNS(root.namespaceURI, 'pivotCaches');
  const one = wb.doc.createElementNS(root.namespaceURI, 'pivotCache');
  one.setAttribute('cacheId', cacheId);
  one.setAttributeNS(NS.r, 'r:id', cacheRel);
  caches.appendChild(one);
  const after = [...root.childNodes].find((n) => n.localName === 'calcPr') || [...root.childNodes].find((n) => n.localName === 'definedNames') || [...root.childNodes].find((n) => n.localName === 'sheets');
  root.insertBefore(caches, after.nextSibling);
  wb.bookDirty = true;
  return { wb, src, piv, area: parseRange(c[which].ref) };
}

/** The data with the case's change made. */
export function applyChange(data, change) {
  const out = data.map((row) => [...row]);
  for (const ch of change) out[ch.r][ch.c] = ch.value;
  return out;
}
