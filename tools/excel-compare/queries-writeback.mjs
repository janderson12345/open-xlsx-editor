// Writes the editor's changes of the fixture's M (queries.js setQueryFormula) to files for Excel to open and refresh
// (excel-queries-check.ps1): each case is the fixture with one or more queries given a new formula, saved by the
// editor. The cases are in queries-writeback-cases.json ([{ name, set: { Query: "M" } }]).
// usage: node tools/excel-compare/queries-writeback.mjs <out folder>
import fs from 'node:fs';
import path from 'node:path';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { setXmlImpl } from '../../src/ooxml/xml.js';
import { Workbook } from '../../src/ooxml/workbook.js';
import { setQueryFormula } from '../../src/ooxml/queries.js';

setXmlImpl({ DOMParser, XMLSerializer });
const out = process.argv[2];
if (!out) throw new Error('usage: node tools/excel-compare/queries-writeback.mjs <out folder>');
fs.mkdirSync(out, { recursive: true });
const cases = JSON.parse(fs.readFileSync(new URL('./queries-writeback-cases.json', import.meta.url), 'utf8'));
for (const c of cases) {
  const wb = Workbook.fromBytes(fs.readFileSync('test/queries-excel.xlsx'));
  for (const [name, m] of Object.entries(c.set)) {
    const err = setQueryFormula(wb, name, m);
    if (err) throw new Error(`${c.name}: ${err.error}`);
  }
  fs.writeFileSync(path.join(out, `${c.name}.xlsx`), wb.toBytes());
}
console.log(`${cases.length} files in ${out}`);
