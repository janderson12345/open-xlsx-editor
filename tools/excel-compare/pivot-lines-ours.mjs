// The editor's result of each step of test/pivot-lines-excel.json (a pivot table at Pivot!C5 of Src!A1:E19), saved for
// Excel to open (excel-pivot-lines-check.ps1): step-1.xlsx...; a step the editor refuses is left out.
// usage: node tools/excel-compare/pivot-lines-ours.mjs <folder>
import fs from 'node:fs';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { setXmlImpl } from '../../src/ooxml/xml.js';
import { readPivots } from '../../src/ooxml/pivot.js';
import { shiftCells } from '../../src/ooxml/rowcol.js';
import { parseRange, colIndex } from '../../src/ooxml/refs.js';
import { pivotBook } from '../../test/pivots.js';

setXmlImpl({ DOMParser, XMLSerializer });
const dir = process.argv[2];
fs.mkdirSync(dir, { recursive: true });
const steps = JSON.parse(fs.readFileSync('test/pivot-lines-excel.json', 'utf8'));
let saved = 0;
steps.forEach((e, k) => {
  const { wb, src, piv } = pivotBook(0);
  const p = readPivots(wb)[0];
  const cells = [];
  for (const [r, line] of wb.sheet(piv).cells) for (const [c, cell] of line) if (cell.v != null) cells.push({ r, c, value: cell.v });
  wb.apply(piv, [...cells.map(({ r, c }) => ({ r, c, value: null })), ...cells.map(({ r, c, value }) => ({ r: r + 2, c: c + 2, value })), { r: 0, c: 0, value: 'beside' }]);
  wb.pkg.setText(p.path, wb.pkg.text(p.path).replace(/(<location\b[^>]*\sref=")[^"]*"/, '$1C5:D9"'));
  const si = e.sheet === 'Src' ? src : piv;
  try {
    const lines = /^(\d+):(\d+)$/.exec(e.range) || /^([A-Z]+):([A-Z]+)$/.exec(e.range);
    if (!e.shift && lines) {
      const rows = /^\d/.test(e.range);
      const a = rows ? Number(lines[1]) - 1 : colIndex(lines[1]);
      const b = rows ? Number(lines[2]) - 1 : colIndex(lines[2]);
      wb.changeLines(si, rows ? 'row' : 'col', a, (e.do === 'insert' ? 1 : -1) * (b - a + 1));
    } else shiftCells(wb, si, parseRange(e.range), e.do === 'insert' ? 1 : -1, e.shift === 'down' || e.shift === 'up' ? 'row' : 'col');
  } catch { return; }
  fs.writeFileSync(`${dir}/step-${k + 1}.xlsx`, wb.toBytes());
  saved++;
});
console.log(`${saved} saved`);
