import { test } from 'node:test';
import assert from 'node:assert/strict';
import { colName, colIndex, parseRef, refName, parseRange, rangeName } from '../src/ooxml/refs.js';
import { formatGeneral, cellText } from '../src/ooxml/numfmt.js';
import { colChars, charsToColWidth, colWidthToPx } from '../src/ooxml/units.js';
import { GENERAL_VALUES, GENERAL_CHARS } from './fixtures/make-fixtures.mjs';
import fs from 'node:fs';

test('column letters both ways, to XFD', () => {
  assert.deepEqual([0, 25, 26, 701, 702, 16383].map(colName), ['A', 'Z', 'AA', 'ZZ', 'AAA', 'XFD']);
  for (let c = 0; c < 16384; c++) assert.equal(colIndex(colName(c)), c);
});

test('references and ranges', () => {
  assert.deepEqual(parseRef('B12'), { r: 11, c: 1 });
  assert.deepEqual(parseRef('$b$12'), { r: 11, c: 1 });
  assert.equal(parseRef('XFE1'), null);
  assert.equal(parseRef('A1048577'), null);
  assert.equal(parseRef('A0'), null);
  assert.equal(refName(0, 0), 'A1');
  assert.deepEqual(parseRange('C3:A1'), { r0: 0, c0: 0, r1: 2, c1: 2 });
  assert.equal(rangeName(parseRange('B2')), 'B2');
  assert.equal(rangeName(parseRange('A1:C3')), 'A1:C3');
});

test('General: up to 11 characters, scientific beyond', () => {
  const cases = [[0, '0'], [1, '1'], [-3.5, '-3.5'], [1 / 3, '0.333333333'], [2 / 3, '0.666666667'], [1234567.891, '1234567.891'],
    [12345678901, '12345678901'], [123456789012, '1.23457E+11'], [0.0001, '0.0001'], [0.00001, '0.00001'], [0.1 + 0.2, '0.3'], [1e20, '1E+20']];
  for (const [n, s] of cases) assert.equal(formatGeneral(n), s, String(n));
  assert.equal(cellText({ t: 'b', v: false }), 'FALSE');
  assert.equal(cellText({ t: 'n', v: null }), '');
  assert.equal(cellText(null), '');
});

// What Excel showed for each value of test/fixtures/general.xlsx in each of its five column widths (its PDF, read
// with tools/excel-compare/cells.cjs), and what formatGeneral gives for the same values in the same widths.
test('General: all 215 values of the probe as Excel shows them, in columns of 30 down to 2.5 characters', () => {
  const excel = JSON.parse(fs.readFileSync('test/general-excel.json', 'utf8')).shown;
  const widths = GENERAL_CHARS.map((ch) => colChars(colWidthToPx(charsToColWidth(ch, 7), 7), 7));
  assert.deepEqual(widths, [30, 8, 6, 4, 2]);
  const wrong = [];
  GENERAL_VALUES.forEach((v, r) => widths.forEach((w, c) => {
    const ours = formatGeneral(v, w);
    if (ours !== excel[r][c]) wrong.push(`${v} in ${w}: Excel ${excel[r][c]}, ours ${ours}`);
  }));
  assert.deepEqual(wrong, []);
});
