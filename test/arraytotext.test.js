import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';

// (Excel 365 over COM: an error in an array is written as its text, an error alone stays the error.)
test('ARRAYTOTEXT writes an error in an array as its text, as Excel does', () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const at = (f) => wb.calc.formulaAt(0, 900, 20, f);
  assert.equal(at('ARRAYTOTEXT({1,#N/A})'), '1, #N/A');
  assert.equal(at('ARRAYTOTEXT(REGEXEXTRACT("ab","(a)(x)?",2))'), 'a, #N/A');
  assert.equal(at('ARRAYTOTEXT(NA())').code, '#N/A');
  assert.equal(at('ARRAYTOTEXT(1/0)').code, '#DIV/0!');
});
