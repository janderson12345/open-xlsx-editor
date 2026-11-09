import { test } from 'node:test';
import assert from 'node:assert/strict';
import { colName, colIndex, parseRef, refName, parseRange, rangeName } from '../src/ooxml/refs.js';
import { formatGeneral, cellText } from '../src/ooxml/numfmt.js';

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
    [12345678901, '12345678901'], [123456789012, '1.23457E+11'], [0.0001, '0.0001'], [0.00001, '1E-05'], [0.1 + 0.2, '0.3'], [1e20, '1E+20']];
  for (const [n, s] of cases) assert.equal(formatGeneral(n), s, String(n));
  assert.equal(cellText({ t: 'b', v: false }), 'FALSE');
  assert.equal(cellText({ t: 'n', v: null }), '');
  assert.equal(cellText(null), '');
});
