// The editor's result of each case of anchor-cases-plain.json on test/fixtures/anchors.xlsx, saved for Excel to open
// (excel-anchors-check.ps1): case-1.xlsx, case-2.xlsx...
// usage: node tools/excel-compare/anchors-ours.mjs <folder>
import fs from 'node:fs';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { setXmlImpl } from '../../src/ooxml/xml.js';
import { Workbook } from '../../src/ooxml/workbook.js';
import { shiftCells } from '../../src/ooxml/rowcol.js';
import { parseRange, colIndex } from '../../src/ooxml/refs.js';

setXmlImpl({ DOMParser, XMLSerializer });
const dir = process.argv[2];
fs.mkdirSync(dir, { recursive: true });
const cases = JSON.parse(fs.readFileSync('tools/excel-compare/anchor-cases-plain.json', 'utf8')).cases;
cases.forEach((c, k) => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/anchors.xlsx')));
  const s = c.step;
  const lines = /^(\d+):(\d+)$/.exec(s.range) || /^([A-Z]+):([A-Z]+)$/.exec(s.range);
  if (!s.shift && lines) {
    const rows = /^\d/.test(s.range);
    const a = rows ? Number(lines[1]) - 1 : colIndex(lines[1]);
    const b = rows ? Number(lines[2]) - 1 : colIndex(lines[2]);
    wb.changeLines(0, rows ? 'row' : 'col', a, (s.do === 'insert' ? 1 : -1) * (b - a + 1));
  } else shiftCells(wb, 0, parseRange(s.range), s.do === 'insert' ? 1 : -1, s.shift === 'down' || s.shift === 'up' ? 'row' : 'col');
  fs.writeFileSync(`${dir}/case-${k + 1}.xlsx`, wb.toBytes());
});
console.log(`${cases.length} saved`);
