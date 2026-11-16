import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { XlError } from '../src/calc/values.js';

// Excel 365's answers for 513 date formulas around its February 29, 1900, serial 60 (test/date1900-excel.json, from
// tools/excel-compare/date1900-cases.mjs and excel-formulas.ps1; docs/CALC.md).
const data = JSON.parse(fs.readFileSync('test/date1900-excel.json', 'utf8'));
// (The coupon functions with a coupon date before 1900 aren't Excel's yet: docs/CALC.md, "Not yet".)
const NOT_YET = new Set(['=COUPDAYS(100,244,2,1)', '=COUPDAYS(50,244,2,1)', '=COUPDAYS(50,243,2,1)', '=COUPPCD(61,1000,4,0)',
  '=COUPDAYBS(50,881,4,0)']);

test('DATE, DAYS360 and the other date functions take February 1900 as Excel does', () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const wrong = [];
  data.formulas.forEach((f, i) => {
    if (NOT_YET.has(f)) return;
    let v = wb.calc.formulaAt(0, 900, 20, f.slice(1));
    if (v instanceof XlError) v = { e: v.code };
    const x = data.excel[i];
    const ok = typeof x === 'object' ? v?.e === x.e : typeof x === 'number' ? Math.abs(v - x) < 1e-12 : v === x;
    if (!ok) wrong.push(`${f}: Excel ${JSON.stringify(x)}, ours ${JSON.stringify(v)}`);
  });
  assert.deepEqual(wrong, []);
});
