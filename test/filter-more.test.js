import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { Workbook } from '../src/ooxml/workbook.js';
import { serializeXml, kid } from '../src/ooxml/xml.js';
import { fromCriteria } from '../src/ooxml/filtermore.js';

// AutoFilter's other filters as Excel does them: test/filter-more-excel.json holds 31 of them, each on a new workbook
// of the same made-up table, its dates days either side of =TODAY() (tools/excel-compare/excel-filter-more.ps1; the day
// Excel ran them recorded). The same table here, its dates from that day, each filter set as the case sets it; the
// rows hidden, the autoFilter element and a colour filter's dxf compared with Excel's.
const excel = JSON.parse(fs.readFileSync('test/filter-more-excel.json', 'utf8'));
const DATES = [-1, 0, 1, -40, 40, -400, -3];
const ROWS = [['apple', 1, 'red'], ['pear', 2, 'blue'], ['Apple', 3, 'red'], ['fig', 1, ''], ['plum', 2.5, 'green'], ['kiwi', 10, 'blue'], ['banana', 7, 'red']];
const DYNAMIC = { 1: 'today', 2: 'yesterday', 3: 'tomorrow', 4: 'thisWeek', 5: 'lastWeek', 7: 'thisMonth', 8: 'lastMonth', 9: 'nextMonth', 10: 'thisQuarter',
  13: 'thisYear', 14: 'lastYear', 16: 'yearToDate', 33: 'aboveAverage', 34: 'belowAverage' };
const argb = (bgr) => `FF${[bgr & 255, (bgr >> 8) & 255, (bgr >> 16) & 255].map((x) => x.toString(16).toUpperCase().padStart(2, '0')).join('')}`;
/** A case's filter as the model takes it, from Excel's AutoFilter arguments. */
function columnOf(how) {
  const [, a, b, c] = how;
  if (how.length === 2) return fromCriteria(a);
  if (how.length === 4) return fromCriteria(a, b === 1 ? 'and' : 'or', c);
  if (b >= 3 && b <= 6) return { top10: { top: b === 3 || b === 5, percent: b >= 5, val: Number(a) }, other: false };
  if (b === 11) return { dynamic: { type: DYNAMIC[a] }, other: false };
  return { color: { rgb: argb(a), font: b === 9 }, other: false };
}
function table() {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const sheet = wb.sheet(0);
  for (const r of [...sheet.cells.keys()]) for (const c of [...sheet.cells.get(r).keys()]) wb.setCell(0, { r, c }, null);
  ['Name', 'Qty', 'Team', 'Date'].forEach((v, c) => wb.setCell(0, { r: 0, c }, v));
  ROWS.forEach((row, k) => { row.forEach((v, c) => { if (v !== '') wb.setCell(0, { r: k + 1, c }, v); }); wb.setCell(0, { r: k + 1, c: 3 }, excel.today + DATES[k]); });
  const date = wb.styleWith(0, { numFmt: 'm/d/yyyy' });
  const cells = ROWS.map((_, k) => ({ r: k + 1, c: 3, s: date }));
  cells.push({ r: 2, c: 0, s: wb.styleWith(0, { fill: { rgb: 'FFFFFF00' } }) }, { r: 4, c: 0, s: wb.styleWith(0, { fill: { rgb: 'FFFFFF00' } }) },
    { r: 6, c: 0, s: wb.styleWith(0, { fill: { rgb: 'FF92D050' } }) }, { r: 3, c: 0, s: wb.styleWith(0, { font: { color: { rgb: 'FFFF0000' } } }) });
  wb.setStyles(0, cells);
  wb.today = excel.today;
  return wb;
}
const tidy = (xml) => xml && xml.replace(/ xmlns(:\w+)?="[^"]*"/g, '').replace(/ xr:uid="[^"]*"/g, '');

test("each of Excel's other filters: the rows hidden and the autoFilter written as Excel's (31 filters)", () => {
  const wrong = [];
  for (const c of excel.cases) {
    const wb = table();
    wb.setAutoFilter(0, { r0: 0, c0: 0, r1: 7, c1: 3, columns: new Map([[c.how[0] - 1, columnOf(c.how)]]) });
    const hidden = [...wb.sheet(0).rows].filter(([, row]) => row.hidden).map(([r]) => r + 1).sort((a, b) => a - b);
    if (JSON.stringify(hidden) !== JSON.stringify(c.hidden)) wrong.push(`${c.name}: Excel hides ${c.hidden}, here ${hidden}`);
    // (A colour filter's dxf is the workbook's next: basic.xlsx has one already, Excel's new workbook none.)
    const el = tidy(serializeXml(kid(wb._part(0).doc.documentElement, 'autoFilter'))).replace(/dxfId="\d+"/, 'dxfId="N"');
    if (el !== tidy(c.autoFilter).replace(/dxfId="\d+"/, 'dxfId="N"')) wrong.push(`${c.name}: Excel ${tidy(c.autoFilter)}\n   here ${el}`);
    if (c.how[2] === 8 || c.how[2] === 9) {
      const dxfs = /<dxfs\b[\s\S]*?<\/dxfs>/.exec(strFromU8(unzipSync(wb.toBytes())['xl/styles.xml']))?.[0] || '';
      const last = dxfs.match(/<dxf>[\s\S]*?<\/dxf>/g)?.at(-1);
      if (last !== /<dxf>[\s\S]*?<\/dxf>/.exec(c.dxfs)[0]) wrong.push(`${c.name}: Excel's dxf ${c.dxfs}, here ${last}`);
    }
  }
  assert.equal(excel.cases.length, 31);
  assert.deepEqual(wrong, []);
});

test('read back: each kind of filter column as written, and filtering again gives the same rows', () => {
  const wb = table();
  const columns = new Map([[1, columnOf([2, '>=2', 1, '<=7'])], [3, columnOf([4, 13, 11])]]);
  wb.setAutoFilter(0, { r0: 0, c0: 0, r1: 7, c1: 3, columns });
  const again = Workbook.fromBytes(wb.toBytes());
  again.today = excel.today;
  const f = again.sheet(0).autoFilter;
  assert.deepEqual([f.columns.get(1).custom, f.columns.get(3).dynamic], [{ and: true, items: [{ op: 'greaterThanOrEqual', val: '2' }, { op: 'lessThanOrEqual', val: '7' }] }, { type: 'thisYear' }]);
  const before = [...again.sheet(0).rows].filter(([, r]) => r.hidden).map(([r]) => r).join();
  again.setAutoFilter(0, f);
  assert.equal([...again.sheet(0).rows].filter(([, r]) => r.hidden).map(([r]) => r).join(), before);
});
