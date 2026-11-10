// From excel-pivot.ps1's output and the workbooks it saved (outside the repository): test/pivot-excel.json, each case's
// cells before and after Excel's refresh and the three pivot parts Excel wrote each time (the pivot table, its cache
// definition and its records). The cache's refreshedBy (the Excel user's name) is left out, and the script stops if
// the workbook's author appears anywhere in what it would write.
// usage: node tools/excel-compare/pivot-parts.mjs <pivot-cases.json> <excel-pivot output.json> <books folder> test/pivot-excel.json
import fs from 'node:fs';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';

const [casesFile, rawFile, books, out] = process.argv.slice(2);
const spec = JSON.parse(fs.readFileSync(casesFile, 'utf8'));
const raw = JSON.parse(fs.readFileSync(rawFile, 'utf8').replace(/^\uFEFF/, ''));
const PARTS = { table: 'xl/pivotTables/pivotTable1.xml', cache: 'xl/pivotCache/pivotCacheDefinition1.xml', records: 'xl/pivotCache/pivotCacheRecords1.xml' };
const people = new Set();
const partsOf = (file) => {
  const z = unzipSync(fs.readFileSync(path.join(books, file)));
  const core = strFromU8(z['docProps/core.xml'] || new Uint8Array());
  for (const m of core.matchAll(/<(?:dc:creator|cp:lastModifiedBy)>([^<]+)</g)) people.add(m[1]);
  return Object.fromEntries(Object.entries(PARTS).map(([k, p]) => [k, strFromU8(z[p]).replace(/ refreshedBy="[^"]*"/, '')]));
};
const cells = (s) => ({ ref: s.ref, cells: s.cells.map(([r, c, v]) => [r, c, v]) });
const cases = spec.cases.map((c, k) => ({
  ...c,
  before: cells(raw[k].before),
  after: cells(raw[k].after),
  parts: { before: partsOf(raw[k].file), after: partsOf(raw[k].file.replace(/\.xlsx$/, '-after.xlsx')) },
}));
const text = JSON.stringify({ data: spec.data, change: spec.change, cases });
for (const p of people) if (p.trim() && text.includes(p)) throw new Error('a workbook author’s name is in the parts: not written');
fs.writeFileSync(out, text);
console.log(`${cases.length} cases, ${(text.length / 1024).toFixed(0)} KB`);
