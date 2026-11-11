// Data validation rules written by the editor (dvedit.js, as its Data Validation box writes them), saved for Excel to
// open (excel-dv-check.ps1): rules.xlsx, five rules on widths.xlsx's first sheet, one of them split by another;
// expected.json: what Excel should read for each cell.
// usage: node tools/excel-compare/dv-ours.mjs <folder>
import fs from 'node:fs';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { setXmlImpl } from '../../src/ooxml/xml.js';
import { Workbook } from '../../src/ooxml/workbook.js';
import { parseRange } from '../../src/ooxml/refs.js';

setXmlImpl({ DOMParser, XMLSerializer });
const dir = process.argv[2];
fs.mkdirSync(dir, { recursive: true });
const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/widths.xlsx')));
const set = (ref, rule) => wb.setValidation(0, [parseRange(ref)], rule);
set('C1:C10', { type: 'whole', op: 'greater', f1: '0' });
set('C4:C5', { type: 'list', f1: 'Yes,No', dropDown: false });
set('D1:D3', { type: 'date', op: 'greaterEqual', f1: '1/1/2024' });
set('E1:E3', { type: 'time', op: 'between', f1: '9:00', f2: '17:00' });
set('F1:F3', { type: 'decimal', op: 'equal', f1: '=$B$1', prompt: { show: true, title: 'Rate', text: 'A rate from B1' }, error: { show: true, style: 'warning', title: 'Odd rate', text: 'Not B1' } });
fs.writeFileSync(`${dir}/rules.xlsx`, wb.toBytes());
// Excel's codes: Type 1 whole, 2 decimal, 3 list, 4 date, 5 time; Operator 1 between, 3 equal, 5 greater, 7 greaterEqual;
// AlertStyle 1 stop, 2 warning. Formula1 as Excel gives it: a date or time as it shows one. C11 has no rule: nothing.
fs.writeFileSync(`${dir}/expected.json`, JSON.stringify({
  C1: [1, 5, '0', 1, ''], C3: [1, 5, '0', 1, ''], C4: [3, null, 'Yes,No', 1, ''], C5: [3, null, 'Yes,No', 1, ''], C6: [1, 5, '0', 1, ''], C10: [1, 5, '0', 1, ''],
  D2: [4, 7, '1/1/2024', 1, ''], E2: [5, 1, '9:00:00 AM', 1, ''], F2: [2, 3, '=$B$1', 2, 'Rate|A rate from B1|Odd rate|Not B1'], C11: [null, null, null, null, ''],
}));
console.log('saved');
