// Each case of dynamic-cases.mjs typed here, as the editor takes an entry, on the probe's data (A1:A6 numbers, B1:B6
// words), and saved for Excel to open (excel-dynamic-check.ps1): Excel reading the same spills, values and formulas
// from the editor's files as it made itself (test/dynamic-excel.json) means the files are written as Excel writes them.
// usage: node tools/excel-compare/dynamic-ours.mjs <folder>   then   pwsh tools/excel-compare/excel-dynamic-check.ps1 <folder> <out.json>
import fs from 'node:fs';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { setXmlImpl } from '../../src/ooxml/xml.js';
import { Workbook } from '../../src/ooxml/workbook.js';
import { entryChanges } from '../../src/ooxml/entry.js';
import { parseRef } from '../../src/ooxml/refs.js';

setXmlImpl({ DOMParser, XMLSerializer });
const dir = process.argv[2];
fs.mkdirSync(dir, { recursive: true });
const { cases } = JSON.parse(fs.readFileSync('tools/excel-compare/dynamic-cases.json', 'utf8'));
const NUMS = [5, 3, 8, 3, 1, 8];
const WORDS = ['pear', 'apple', 'fig', 'apple', 'kiwi', 'plum'];
cases.forEach((k, n) => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  // (The probe's sheet first, so Excel's Worksheets.Item(1) is it.)
  const si = wb.addSheet(0, 'Dyn');
  wb.apply(si, NUMS.flatMap((v, r) => [{ r, c: 0, value: v }, { r, c: 1, value: WORDS[r] }]));
  const type = (ref, text) => { const { r, c } = parseRef(ref); wb.apply(si, [entryChanges(wb, si, r, c, text).content]); };
  if (k.block) type(k.block.ref, k.block.value);
  type('D1', k.formula);
  if (k.then) type(k.then.ref, k.then.formula);
  fs.writeFileSync(`${dir}/ours-${n + 1}.xlsx`, wb.toBytes());
});
console.log(`${cases.length} saved`);
