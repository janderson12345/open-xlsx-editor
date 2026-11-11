// Each case of test/pivot-excel.json made here as a new PivotTable (pivotnew.js), from the made-up sales on a sheet
// "Src" to Pivot!A3, and saved for Excel to open and refresh (excel-pivot-check.ps1): the two the same means Excel takes
// the new pivot's parts and cells as its own.
// usage: node tools/excel-compare/pivotnew-ours.mjs <folder>   then   pwsh tools/excel-compare/excel-pivot-check.ps1 <folder>
import fs from 'node:fs';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { setXmlImpl } from '../../src/ooxml/xml.js';
import { Workbook } from '../../src/ooxml/workbook.js';
import { createPivot } from '../../src/ooxml/pivotnew.js';
import { PIVOT } from '../../test/pivots.js';

setXmlImpl({ DOMParser, XMLSerializer });
const dir = process.argv[2];
fs.mkdirSync(dir, { recursive: true });
const F = PIVOT.data[0];
PIVOT.cases.forEach((c, k) => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const src = wb.addSheet(wb.sheets.length, 'Src');
  const piv = wb.addSheet(wb.sheets.length, 'Pivot');
  wb.apply(src, PIVOT.data.flatMap((row, r) => row.map((value, col) => ({ r, c: col, value }))));
  const layout = { rows: c.rows.map((f) => F.indexOf(f)), cols: c.cols.map((f) => F.indexOf(f)), values: c.values.map((v) => ({ field: F.indexOf(v.field), fn: v.fn })) };
  const out = createPivot(wb, { si: src, r0: 0, c0: 0, r1: PIVOT.data.length - 1, c1: F.length - 1 }, { si: piv, r: 2, c: 0 }, layout);
  if (out.error) throw new Error(`${c.name}: ${out.error}`);
  for (const ch of out.changes) wb.apply(ch.si, ch.contents);
  fs.writeFileSync(`${dir}/ours-${k + 1}.xlsx`, wb.toBytes());
});
console.log(`${PIVOT.cases.length} saved`);
