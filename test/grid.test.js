import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Axis } from '../src/editor/axis.js';
import { sheetGeometry } from '../src/editor/grid.js';
import { Workbook } from '../src/ooxml/workbook.js';

test('an axis finds the index under a position, passing over hidden ones', () => {
  const a = new Axis(10, 20, new Map([[2, 40], [3, 0]]));
  assert.equal(a.end, 20 * 8 + 40);
  assert.equal(a.start(4), 80);
  assert.equal(a.size(3), 0);
  assert.deepEqual([0, 19, 20, 39, 40, 79, 80, 99, 100, 10000].map((p) => a.indexAt(p)), [0, 0, 1, 1, 2, 2, 4, 4, 5, 9]);
});

test('sheet geometry from basic.xlsx: default 64px columns and 20px rows, its own widths and heights, hidden ones 0', () => {
  const sheet = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx'))).sheet(0);
  const { rows, cols } = sheetGeometry(sheet);
  assert.deepEqual([0, 1, 2, 3, 4, 5].map((c) => cols.size(c)), [64, 145, 64, 0, 64, 215]);
  assert.deepEqual([0, 3, 5].map((r) => rows.size(r)), [20, 40, 0]);
  assert.ok(rows.count > sheet.maxRow && cols.count > sheet.maxCol);
});
