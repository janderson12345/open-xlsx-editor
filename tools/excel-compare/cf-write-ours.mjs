// Each case of cf-write-cases.json added by the editor (cfedit.js) to the same cells Excel had, saved for Excel to open
// (excel-cf-look.ps1, which compares how Excel draws them with its own: cf-write-<k>.xlsx beside them).
// usage: node tools/excel-compare/cf-write-ours.mjs <folder>
import fs from 'node:fs';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { setXmlImpl } from '../../src/ooxml/xml.js';
import { Workbook } from '../../src/ooxml/workbook.js';
import { parseRange } from '../../src/ooxml/refs.js';

setXmlImpl({ DOMParser, XMLSerializer });
const dir = process.argv[2];
fs.mkdirSync(dir, { recursive: true });
const cases = JSON.parse(fs.readFileSync('tools/excel-compare/cf-write-cases.json', 'utf8')).cases;
const words = ['banana', 'apple', 'cherry', 'band', 'apple', 'kiwi', 'mango', 'banana', 'pear', 'plum'];
cases.forEach((c, k) => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/widths.xlsx')));
  const si = wb.addSheet(0, 'Rules');
  wb.apply(si, words.flatMap((w, i) => [{ r: i, c: 0, value: i + 1 }, { r: i, c: 1, value: w }]));
  for (const r of c.rules) wb.setConditional(si, [parseRange(r.range)], r);
  fs.writeFileSync(`${dir}/ours-${k + 1}.xlsx`, wb.toBytes());
});
console.log(`${cases.length} saved`);
