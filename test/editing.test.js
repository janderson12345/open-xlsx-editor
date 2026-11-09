import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { parseInput, editText } from '../src/ooxml/input.js';
import { step, jump, lastCell } from '../src/editor/navigate.js';
import { Workbook } from '../src/ooxml/workbook.js';
import { textOn } from '../src/editor/tabs.js';
import { insertable, rangeText, cycleRef } from '../src/editor/pointing.js';

test('typed entries read as Excel reads them', () => {
  const v = (s) => parseInput(s).value;
  assert.deepEqual(['12', '-3.5', '1,234', '1,234.5', '50%', '1e3', '.5', 'TRUE', 'false', "'007", '', 'abc', '1,23', '12 apples'].map(v),
    [12, -3.5, 1234, 1234.5, 0.5, 1000, 0.5, true, false, '007', null, 'abc', '1,23', '12 apples']);
  assert.equal(parseInput('=a1+sum(b1:b2').formula, 'A1+SUM(B1:B2)');
  assert.match(parseInput('=1+').error, /problem with this formula/);
  assert.equal(editText({ t: 'n', v: 0.1 + 0.2 }), String(0.1 + 0.2));
  assert.equal(editText({ t: 'b', v: true }), 'TRUE');
  assert.equal(editText({ t: 'n', v: 4, f: 'B2*2' }), '=B2*2');
});

test('moving: steps over merged ranges, Ctrl+arrow to the edges of blocks, Ctrl+End to the last used cell', () => {
  const sheet = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx'))).sheet(0);
  const limits = { rows: 200, cols: 30 };
  // Into the merged range A12:C13 from above: its top-left; out of it downwards: below its last row.
  assert.deepEqual(step(sheet, 10, 1, 1, 0, limits), { r: 11, c: 0 });
  assert.deepEqual(step(sheet, 11, 0, 1, 0, limits), { r: 13, c: 0 });
  assert.deepEqual(step(sheet, 0, 0, -1, 0, limits), { r: 0, c: 0 });
  // A1 down: the block A1:A12 ends at A12; from there, the next filled cell (A15); from A15, the edge.
  assert.deepEqual(jump(sheet, 0, 0, 1, 0, limits), { r: 11, c: 0 });
  assert.deepEqual(jump(sheet, 11, 0, 1, 0, limits), { r: 14, c: 0 });
  assert.deepEqual(jump(sheet, 39, 0, 1, 0, limits), { r: 200, c: 0 });
  assert.deepEqual(lastCell(sheet), { r: 39, c: 7 });
});

test('a cell edit updates the model in place and leaves the other cells as read', () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/large.xlsx')));
  const before = wb.sheet(0).cells.get(10).get(3);
  wb.setCell(0, 'D11', 'changed');
  assert.deepEqual(wb.sheet(0).cells.get(10).get(3), { t: 's', v: 'changed', s: 0, f: null });
  assert.notDeepEqual(before, wb.sheet(0).cells.get(10).get(3));
  wb.setCell(0, 'Z6000', 1);
  assert.equal(wb.sheet(0).maxRow, 5999);
  assert.equal(wb.sheet(0).cells.get(11).get(3).v, 1203);
});

test('sheet tabs: the sheets that can be shown, tab colours resolved as cell colours are, readable text on them', () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  assert.deepEqual(wb.sheets.map((_, i) => wb.canShow(i)), [true, true, false]); // (Hidden is state="hidden")
  // (theme 5 is accent 2, E97132 in the default theme, lightened 40%: Excel's "Orange, Accent 2, Lighter 40%")
  assert.deepEqual(wb.sheets.map((_, i) => wb.tabColor(i)), [null, '#F1A983', null]);
  assert.deepEqual(['#F1A983', '#C00000', '#FFFF00', '#002060'].map(textOn), ['#000000', '#ffffff', '#000000', '#ffffff']);
});

test('pointing: where a reference can go, a range as text, F4 cycling $ as Excel does', () => {
  const ok = (text, at = text.length) => insertable(text, at);
  assert.deepEqual(['=', '=A1+', '=SUM(', '=SUM(A1,', '=A1:', '=1 * ', '=A1&'].map((t) => ok(t)), [true, true, true, true, true, true, true]);
  assert.deepEqual(['=A1', '=SUM(A1', '="a+', 'A1+', '=A1%', '=2'].map((t) => ok(t)), [false, false, false, false, false, false]);
  assert.equal(ok('=+A1', 2), false); // (a reference just after the caret)
  assert.equal(ok('=SUM(,A1)', 5), true);
  assert.equal(rangeText({ r: 4, c: 2 }, { r: 1, c: 1 }), 'B2:C5');
  assert.equal(rangeText({ r: 0, c: 0 }, { r: 0, c: 0 }), 'A1');
  let text = '=SUM(B2:C5)+H1';
  const seen = [];
  for (let k = 0; k < 4; k++) { text = cycleRef(text, text.length).text; seen.push(text.slice(text.indexOf(')+') + 2)); }
  assert.deepEqual(seen, ['$H$1', 'H$1', '$H1', 'H1']);
  assert.equal(cycleRef('=SUM(B2:C5)', 7).text, '=SUM($B$2:$C$5)');
  assert.equal(cycleRef("='My Data'!A1", 13).text, "='My Data'!$A$1");
  assert.equal(cycleRef('=A:A', 4).text, '=$A:$A');
  assert.equal(cycleRef('=1+2', 4), null);
});
