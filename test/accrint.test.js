import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { XlError } from '../src/calc/values.js';

// Excel 365's answers for 4251 ACCRINT formulas (test/accrint-excel.json, from tools/excel-compare/accrint-cases.mjs
// and excel-formulas.ps1): dates near 1900-01-01 and 9999-12-31 over every frequency, basis and calc_method, 1500 of
// them held out, sweeps by the day near 9999-12-31, and 300 with issue on a coupon date (docs/CALC.md).
const data = JSON.parse(fs.readFileSync('test/accrint-excel.json', 'utf8'));
const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
const results = data.formulas.map((f, i) => {
  const v = wb.calc.formulaAt(0, 900 + (i % 100), 20, f.slice(1));
  return { f, excel: data.excel[i], ours: v instanceof XlError ? { e: v.code } : v };
});

test('ACCRINT is #NUM! exactly where Excel gives it near the ends of its dates', () => {
  const wrong = results.filter(({ excel, ours }) => (typeof excel === 'object') !== (typeof ours === 'object') || (typeof excel === 'object' && excel.e !== ours.e));
  assert.equal(results.filter((r) => typeof r.excel === 'object').length, 1911);
  assert.deepEqual(wrong.map((r) => r.f), []);
});

test('ACCRINT\'s values are Excel\'s but for the 38 cases of a rule not found yet', () => {
  // (Each has a date on the 29th to the 31st or in February, or issue on a coupon date with settlement before
  // first_interest and basis 1 to 3: docs/CALC.md, "Not yet".)
  const wrong = results.filter(({ excel, ours }) => typeof excel === 'number' && !(Math.abs(ours - excel) <= 1e-12 * Math.max(1, Math.abs(excel))));
  assert.ok(wrong.length <= 38, wrong.map((r) => `${r.f}: Excel ${r.excel}, ours ${r.ours}`).join('\n'));
  assert.ok(results.length - wrong.length >= 4213);
});

test('ACCRINT keeps the coupon walk within Excel\'s dates: ACCRINT(1,731,731,...) is #NUM!, stepping back past 1900', () => {
  assert.equal(wb.calc.formulaAt(0, 900, 20, 'ACCRINT(1,731,731,0.1,1000,4,1,FALSE)').code, '#NUM!');
  assert.equal(wb.calc.formulaAt(0, 900, 20, 'ACCRINT(2958000,2958440,2958450,0.1,1000,4)').code, '#NUM!');
  // (The period after first_interest ends on 9999-12-31 itself: worked out, as in Excel's sweep.)
  assert.equal(typeof wb.calc.formulaAt(0, 900, 20, 'ACCRINT(2958065,2958281,2958464,0.1,1000,2,0,TRUE)'), 'number');
});
