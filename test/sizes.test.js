import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { Workbook } from '../src/ooxml/workbook.js';
import { typedWidth, pxWidth } from '../src/ooxml/sizes.js';

// What Excel saved after the same steps over COM (ColumnWidth, Hidden, RowHeight, AutoFit; docs/SIZES.md).
const fresh = () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/formulas.xlsx')));
  return { wb, si: wb.addSheet(wb.sheets.length) };
};
const sheetXml = (wb, si) => {
  const files = unzipSync(wb.toBytes());
  return strFromU8(files[wb.sheets[si].path]);
};

test('widths as Excel stores them: whole pixels over the digit width, cut to 1/256', () => {
  assert.deepEqual([20, 12.5, 5, 9, 255].map(typedWidth), [20.7109375, 13.28515625, 5.7109375, 9.7109375, 255.7109375]);
  // (AutoFit's pixels: "Hello world" 77, "W" 19, "A much longer piece of text here" 204.)
  assert.deepEqual([77, 19, 204].map(pxWidth), [11, 2.7109375, 29.140625]);
});

test('columns: set, split where a range cuts them, hidden as width 0', () => {
  const { wb, si } = fresh();
  wb.setLines(si, 'col', 1, 1, { width: typedWidth(20) });
  wb.setLines(si, 'col', 3, 5, { width: typedWidth(5) });
  wb.setLines(si, 'col', 4, 4, { width: typedWidth(9) });
  wb.setLines(si, 'col', 7, 7, { hidden: true });
  wb.setLines(si, 'col', 11, 11, { width: pxWidth(77), bestFit: true });
  const cols = sheetXml(wb, si).match(/<cols>.*<\/cols>/)[0].match(/<col [^>]*\/>/g);
  assert.deepEqual(cols, [
    '<col min="2" max="2" width="20.7109375" customWidth="1"/>',
    '<col min="4" max="4" width="5.7109375" customWidth="1"/>',
    '<col min="5" max="5" width="9.7109375" customWidth="1"/>',
    '<col min="6" max="6" width="5.7109375" customWidth="1"/>',
    '<col min="8" max="8" width="0" hidden="1" customWidth="1"/>',
    '<col min="12" max="12" width="11" customWidth="1" bestFit="1"/>',
  ]);
  const model = wb.sheet(si);
  assert.equal(model.cols.find((c) => c.min === 7).hidden, true);
  wb.setLines(si, 'col', 7, 7, { hidden: false });
  assert.equal(wb.sheet(si).cols.find((c) => c.min === 7).hidden, false);
});

test('rows: a height in points, hidden rows, AutoFit taking the fixed height away', () => {
  const { wb, si } = fresh();
  wb.setLines(si, 'row', 1, 1, { height: 30 });
  wb.setLines(si, 'row', 3, 3, { hidden: true });
  wb.setLines(si, 'row', 5, 5, { height: 0 });
  wb.setLines(si, 'row', 9, 9, { height: 15 });
  let rows = sheetXml(wb, si).match(/<row [^>]*>/g);
  assert.deepEqual(rows, ['<row r="2" ht="30" customHeight="1"/>', '<row r="4" hidden="1"/>', '<row r="6" hidden="1"/>', '<row r="10" ht="15" customHeight="1"/>']);
  assert.deepEqual([wb.sheet(si).rows.get(1).ht, wb.sheet(si).rows.get(3).hidden], [30, true]);
  wb.setLines(si, 'row', 1, 1, { height: null });
  rows = sheetXml(wb, si).match(/<row [^>]*>/g);
  assert.equal(rows[0], '<row r="2"/>');
});
