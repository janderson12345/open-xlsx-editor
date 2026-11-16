import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { XlError, Arr } from '../src/calc/values.js';

// Excel 365's answers for 924 REGEXTEST, REGEXEXTRACT and REGEXREPLACE formulas over the PCRE2 features JavaScript
// hasn't got (test/regex-excel.json, from tools/excel-compare/regex-cases.mjs and excel-formulas.ps1; the last 600
// random patterns, held out). docs/CALC.md has the rules.
const data = JSON.parse(fs.readFileSync('test/regex-excel.json', 'utf8'));
const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
const plain = (v) => {
  if (v instanceof Arr) v = v.rows[0][0];
  if (v instanceof XlError) return { e: v.code };
  return v;
};
const same = (ours, excel) => (excel && typeof excel === 'object' ? ours?.e === excel.e : ours === excel);
// (Not Excel's yet, docs/CALC.md "Not yet": \G in an alternation after an empty match, and three random patterns
// mixing conditionals, \K and \G in repeated groups.)
const NOT_YET = 8;

test('regular expressions give Excel\'s answers: \\K, recursion, conditionals, \\G, \\X, verbs, options partway, errors', () => {
  const wrong = [];
  data.formulas.forEach((f, i) => {
    const ours = plain(wb.calc.formulaAt(0, 900, 20, f.slice(1)));
    if (!same(ours, data.excel[i])) wrong.push(`${f}: Excel ${JSON.stringify(data.excel[i])}, ours ${JSON.stringify(ours)}`);
  });
  assert.ok(wrong.length <= NOT_YET, wrong.join('\n'));
});

test('a pattern that backtracks without end is #VALUE!, as at PCRE2\'s match limit, and soon', () => {
  for (const f of ['REGEXTEST(REPT("a",30)&"b","^(a+)+$")', 'REGEXTEST(REPT("ab",20)&"c","^(ab|a|b)*$")',
    'REGEXREPLACE(REPT("a",3000)&"!","(a|aa)+$","#")']) {
    const t = Date.now();
    assert.equal(wb.calc.formulaAt(0, 900, 20, f).code, '#VALUE!', f);
    // (About 0.7 s alone; up to 8 times that while the other test files run beside it. Without the step limit these
    // take minutes, so 15 s still tells the two apart.)
    assert.ok(Date.now() - t < 15000, `${f} took ${Date.now() - t} ms`);
  }
  // (One that needn't backtrack is worked out, however long.)
  assert.equal(wb.calc.formulaAt(0, 900, 20, 'REGEXTEST(REPT("a",28),"^(a+)+$")'), true);
  assert.equal(wb.calc.formulaAt(0, 900, 20, 'LEN(REGEXEXTRACT(REPT("ab",10000),"(?:ab)+"))'), 20000);
});
