// Each move of test/pivot-move-excel.json made here (Excel's pivot of one layout, its fields moved to another's by
// pivotfields.js) and saved for Excel to open and refresh (excel-pivot-check.ps1).
// usage: node tools/excel-compare/pivotmove-ours.mjs <folder>   then   pwsh tools/excel-compare/excel-pivot-check.ps1 <folder>
import fs from 'node:fs';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { setXmlImpl } from '../../src/ooxml/xml.js';
import { readPivots } from '../../src/ooxml/pivot.js';
import { movePivotFields } from '../../src/ooxml/pivotfields.js';
import { PIVOT, pivotBook } from '../../test/pivots.js';

setXmlImpl({ DOMParser, XMLSerializer });
const dir = process.argv[2];
fs.mkdirSync(dir, { recursive: true });
const F = PIVOT.data[0];
const moves = JSON.parse(fs.readFileSync('test/pivot-move-excel.json', 'utf8'));
moves.forEach((m, k) => {
  const { wb } = pivotBook(m.from, 'before');
  const c = PIVOT.cases[m.to];
  const out = movePivotFields(wb, readPivots(wb)[0], { rows: c.rows.map((f) => F.indexOf(f)), cols: c.cols.map((f) => F.indexOf(f)), values: c.values.map((v) => ({ field: F.indexOf(v.field), fn: v.fn })) });
  if (out.error) throw new Error(out.error);
  for (const ch of out.changes) if (ch.contents.length) wb.apply(ch.si, ch.contents);
  fs.writeFileSync(`${dir}/ours-${k + 1}.xlsx`, wb.toBytes());
});
console.log(`${moves.length} saved`);
