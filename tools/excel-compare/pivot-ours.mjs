// The editor's refresh of each case of test/pivot-excel.json, saved for Excel to open (excel-pivot-check.ps1): Excel's
// pivot table of the made-up sales, the source changed as Excel changed it, refreshed here, saved.
// usage: node tools/excel-compare/pivot-ours.mjs <folder>
import fs from 'node:fs';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { setXmlImpl } from '../../src/ooxml/xml.js';
import { readPivots, refreshPivot } from '../../src/ooxml/pivot.js';
import { PIVOT, pivotBook } from '../../test/pivots.js';

setXmlImpl({ DOMParser, XMLSerializer });
const dir = process.argv[2];
fs.mkdirSync(dir, { recursive: true });
PIVOT.cases.forEach((c, k) => {
  const { wb, src } = pivotBook(k, 'before');
  for (const ch of PIVOT.change) wb.apply(src, [{ r: ch.r, c: ch.c, value: ch.value }]);
  const out = refreshPivot(wb, readPivots(wb)[0]);
  if (out.error) throw new Error(`${c.name}: ${out.error}`);
  for (const ch of out.changes) wb.apply(ch.si, ch.contents);
  fs.writeFileSync(`${dir}/ours-${k + 1}.xlsx`, wb.toBytes());
});
console.log(`${PIVOT.cases.length} saved`);
