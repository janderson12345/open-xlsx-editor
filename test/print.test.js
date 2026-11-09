import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Axis } from '../src/editor/axis.js';
import { breakAxis, paginate } from '../src/editor/print.js';
import { Workbook } from '../src/ooxml/workbook.js';

const sheetOf = (f) => Workbook.fromBytes(new Uint8Array(fs.readFileSync(`test/fixtures/${f}`))).sheet(0);

test('an axis breaks into runs that fit, each taking at least one index, hidden ones costing nothing', () => {
  const a = new Axis(10, 20, new Map([[2, 0], [5, 100]]));
  assert.deepEqual(breakAxis(a, 9, 60), [[0, 3], [4, 4], [5, 5], [6, 8], [9, 9]]);
  assert.deepEqual(breakAxis(a, 0, 60), [[0, 0]]);
});

test('page setup read from the sheet, Normal margins when it has none', () => {
  const basic = sheetOf('basic.xlsx').page;
  assert.deepEqual(basic.margins, { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 });
  assert.equal(basic.paperSize, 1);
  assert.equal(sheetOf('prefixed.xlsx').page.margins.left, 0.7);
});

test('pages down then over: large.xlsx splits its 5,000 rows down and its 20 columns across', () => {
  const sheet = sheetOf('large.xlsx');
  const { pages, paper, area } = paginate(sheet);
  assert.deepEqual(paper, { w: 816, h: 1056 });
  assert.equal(area.x, 0.7 * 96);
  // 9.5in of 20px rows is 45 rows a page; 7.1in of 64px columns is 10 columns.
  assert.deepEqual(pages[0], { r0: 0, r1: 44, c0: 0, c1: 9 });
  assert.deepEqual(pages[1], { r0: 45, r1: 89, c0: 0, c1: 9 });
  assert.equal(pages.length, Math.ceil(5000 / 45) * 2);
  assert.equal(pages[pages.length / 2].c0, 10);
});
