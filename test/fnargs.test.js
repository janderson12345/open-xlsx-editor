import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { catalog, functionInfo } from '../src/calc/catalog.js';
import { argSlots, argBoxes, minArgs, buildCall, splitArgs, callForCaret } from '../src/calc/fnargs.js';
import { describeFunction, describeArg } from '../src/calc/describe.js';
import { parseInput } from '../src/ooxml/input.js';
import { argumentValue, resultText } from '../src/editor/fnargsbox.js';
import { build } from './steps.js';

const json = (p) => JSON.parse(fs.readFileSync(p, 'utf8').replace(/^\uFEFF/, ''));
// (What Excel took over COM: tools/excel-compare/excel-fnargs.ps1 and excel-fnmin.ps1.)
const TAKEN = json('test/fnargs-excel.json');
const FEWEST = json('test/fnmin-excel.json');

test('Function Arguments: the fewest arguments of every listed function are the ones Excel takes', () => {
  assert.equal(FEWEST.length, catalog().length);
  for (const x of FEWEST) assert.equal(minArgs(functionInfo(x.name)), x.min, x.name);
});

test('Function Arguments: the formula written is one Excel takes and keeps as written', () => {
  const taken = new Map(TAKEN.map((x) => [x.typed, x]));
  // (Each box's values, the call written, and Excel's verdict on that formula.)
  const cases = [
    ['SUM', ['A1:A3'], '=SUM(A1:A3)'], ['SUM', ['A1:A3', '', '5'], '=SUM(A1:A3,,5)'], ['SUM', ['', '5'], '=SUM(,5)'],
    ['IF', ['A1>1'], '=IF(A1>1,)'], ['IF', ['A1>1', '', '"no"'], '=IF(A1>1,,"no")'], ['IF', ['A1>1', '"yes"', ''], '=IF(A1>1,"yes")'],
    ['IF', ['', '"yes"'], '=IF(,"yes")'], ['VLOOKUP', ['2', 'A1:A3', '1', ''], '=VLOOKUP(2,A1:A3,1)'], ['VLOOKUP', [], '=VLOOKUP(,,)'],
    ['VLOOKUP', ['', 'A1:A3', '1'], '=VLOOKUP(,A1:A3,1)'], ['VLOOKUP', ['2', '', '1'], '=VLOOKUP(2,,1)'],
    ['ROUND', ['1.25'], '=ROUND(1.25,)'], ['ROUND', ['', '1'], '=ROUND(,1)'], ['LEFT', ['B1'], '=LEFT(B1)'],
    ['PMT', ['0.01', '12', '1000', '', ''], '=PMT(0.01,12,1000)'], ['PMT', ['0.01', '12', '1000', '', '1'], '=PMT(0.01,12,1000,,1)'],
    ['TODAY', [], '=TODAY()'], ['PI', [], '=PI()'], ['COUNTIF', ['A1:A3', '">1"'], '=COUNTIF(A1:A3,">1")'],
    ['TEXTJOIN', ['","', 'TRUE', 'A1:A3', ''], '=TEXTJOIN(",",TRUE,A1:A3)'], ['CONCAT', ['B1', '', 'B1'], '=CONCAT(B1,,B1)'],
    ['MID', ['B1', '2'], '=MID(B1,2,)'], ['AVERAGE', ['', '', ''], null],
  ];
  for (const [name, values, formula] of cases) {
    const got = buildCall(name, values);
    if (formula == null) { assert.equal(got.refused, true, name); continue; }
    assert.equal(`=${got.text}`, formula);
    assert.equal(got.refused, false, formula);
    assert.equal(taken.get(formula)?.ok, true, `Excel takes ${formula}`);
    assert.equal(taken.get(formula).formula, formula, `and keeps it as written: ${formula}`);
  }
  // (Nothing filled in a function that needs one argument: Excel refuses =SUM(), =LEFT(), =AVERAGE().)
  for (const name of ['SUM', 'LEFT', 'AVERAGE', 'COUNT']) {
    const got = buildCall(name, []);
    assert.deepEqual(got, { text: `${name}()`, refused: true });
    assert.equal(taken.get(`=${name}()`).ok, false);
  }
  // (And the ones Excel refuses for too few aren't written.)
  for (const f of ['=IF(A1>1)', '=VLOOKUP(2,A1:A3)', '=ROUND(1.25)', '=MID(B1,2)', '=PMT(0.01,12)', '=TEXTJOIN(",",TRUE)']) assert.equal(taken.get(f).ok, false, f);
});

test('Function Arguments: the boxes, required and optional, a repeated one adding a box as the last is filled', () => {
  const names = (n, v) => argBoxes(functionInfo(n), v).map((b) => b.name + (b.optional ? '?' : '')).join(' ');
  assert.equal(names('SUM', []), 'number1 number2?');
  assert.equal(names('SUM', ['1']), 'number1 number2?');
  assert.equal(names('SUM', ['1', '2']), 'number1 number2? number3?');
  assert.equal(names('SUM', ['1', '', '3']), 'number1 number2? number3? number4?');
  assert.equal(names('VLOOKUP', []), 'lookup_value table_array col_index_num range_lookup?');
  assert.equal(names('SUMIFS', ['A:A', 'B:B']), 'sum_range criteria_range1 criteria1 criteria_range2? criteria2?');
  assert.equal(names('IFS', ['x']), 'logical_test1 value_if_true1 logical_test2? value_if_true2?');
  assert.equal(names('PI', []), '');
  assert.equal(argSlots(functionInfo('REDUCE')).fixed.map((a) => a.optional).join(), 'true,false,false');
  // (At most 255 arguments, as Excel takes.)
  assert.equal(argBoxes(functionInfo('SUM'), Array(300).fill('1')).length, 255);
});

test('Function Arguments: every listed function and argument described', () => {
  for (const f of catalog()) {
    assert.match(describeFunction(f.name), /^[A-Z].+\.$/, f.name);
    for (const b of argBoxes(f, Array(6).fill('x'))) assert.match(describeArg(f.name, b.name), /^[A-Z0-9"].+\.$/, `${f.name} ${b.name}`);
  }
  assert.equal(describeArg('PMT', 'type'), '0 or left out for payments at the end of each period; 1 for the start.');
  assert.match(describeArg('T.TEST', 'type'), /paired/);
  assert.match(describeArg('SUM', 'number7'), /add up/);
  assert.match(describeArg('INDIRECT', 'a1'), /R1C1/);
});

test('Function Arguments: the call at the caret, its arguments read back', () => {
  assert.deepEqual(callForCaret('=SUM', 4), { name: 'SUM', start: 1, end: 4, args: [] });
  assert.deepEqual(callForCaret('=SUM(', 5), { name: 'SUM', start: 1, end: 5, args: [] });
  assert.deepEqual(callForCaret('=1+sum(A1,"a,b",MAX(1,2))*2', 7), { name: 'SUM', start: 3, end: 25, args: ['A1', '"a,b"', 'MAX(1,2)'] });
  assert.deepEqual(callForCaret('=IF(A1>1,SUM(A1:A3),"x")', 14), { name: 'SUM', start: 9, end: 19, args: ['A1:A3'] });
  assert.equal(callForCaret('=IF(A1>1,SUM(A1:A3),"x")', 14, { inside: false }), null);
  assert.equal(callForCaret('=SUM(A1', 7, { inside: false }), null);
  assert.equal(callForCaret('=NOPE', 5), null);
  assert.equal(callForCaret('SUM', 3), null);
  assert.deepEqual(splitArgs(' A1 , {1,2;3,4}, \'My, sheet\'!B2'), [' A1 ', ' {1,2;3,4}', " 'My, sheet'!B2"]);
  // (What the box writes is a formula the editor takes.)
  assert.ok(parseInput(`=${buildCall('VLOOKUP', []).text}`).formula);
});

test('Function Arguments: each argument\'s value and the formula result, at the cell', () => {
  const { wb, si } = build({ cells: { A1: '1', A2: '2', A3: '3', B1: 'abc', C1: '=A1*10' } });
  wb.calculate();
  const v = (t) => argumentValue(wb.calc, si, 4, 3, t);
  assert.equal(v('A1:A3'), '{1;2;3}');
  assert.equal(v('A1:B1'), '{1,"abc"}');
  assert.equal(v('B1'), '"abc"');
  assert.equal(v('C1'), '10');
  assert.equal(v('A1>1'), 'FALSE');
  assert.equal(v('1/0'), '#DIV/0!');
  assert.equal(v(''), '');
  assert.equal(v('(('), '');
  assert.equal(resultText(wb.calc, si, 4, 3, 'SUM(A1:A3,,5)'), '11');
  assert.equal(resultText(wb.calc, si, 4, 3, 'LEFT(B1,2)'), 'ab');
  assert.equal(resultText(wb.calc, si, 4, 3, 'VLOOKUP(,,)'), '#N/A');
  // (Every formula Excel took, worked out here as Excel did: its cell's text, but where Excel formatted the number.)
  for (const x of TAKEN.filter((t) => t.ok && !/^=(PMT|TODAY|PI)\(/.test(t.typed))) assert.equal(resultText(wb.calc, si, 0, 3, x.typed.slice(1)), x.text, x.typed);
});
