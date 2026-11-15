import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { catalog } from '../src/calc/catalog.js';
import { argBoxes, buildCall, needsQuotes } from '../src/calc/fnargs.js';
import { argKind } from '../src/calc/argkinds.js';
import { parseInput } from '../src/ooxml/input.js';

// (What Excel made of each text typed after "=", over COM: tools/excel-compare/excel-fnquote.ps1. A defined name Rate
// and a sheet "My Data" were in that workbook.)
const READ = JSON.parse(fs.readFileSync('test/fnquote-excel.json', 'utf8').replace(/^\uFEFF/, ''));
const isName = (n) => n.toLowerCase() === 'rate';

test('Function Arguments quoting: a text reads as a formula here exactly when Excel takes it as one', () => {
  assert.equal(READ.length, 58);
  for (const x of READ) assert.equal(!parseInput(`=${x.typed}`).error, x.ok, x.typed);
});

test('Function Arguments quoting: text quoted when Excel refuses it as a formula, or reads it only as names that are not defined', () => {
  for (const x of READ) {
    if (!x.typed.trim()) continue;
    // (Excel's verdict: refused, or #NAME? from names alone; a misspelt function's #NAME? isn't text.)
    const namesOnly = x.ok && x.text === '#NAME?' && !/\(/.test(x.typed);
    const want = !x.typed.startsWith('"') && (!x.ok || namesOnly);
    // (Two differences, both quoted here: R1C1, which Excel reads as a reference in A1 style too ($A$1), and N/A, names
    // alone that Excel works out to #VALUE! rather than #NAME?.)
    if (x.typed === 'R1C1' || x.typed === 'N/A') { assert.equal(needsQuotes(x.typed, isName), true); continue; }
    assert.equal(needsQuotes(x.typed, isName), want, x.typed);
  }
  assert.equal(needsQuotes('Rate', isName), false);
  assert.equal(needsQuotes('rate*2', isName), false);
  assert.equal(needsQuotes('', isName), false);
});

test('Function Arguments quoting: only in text and any boxes, quotes inside doubled; a box read back as Excel would', () => {
  const call = (n, v) => buildCall(n, v, { isName }).text;
  assert.equal(call('COUNTIF', ['A1:A3', '>1']), 'COUNTIF(A1:A3,">1")');
  assert.equal(call('IF', ['A1>1', 'yes', 'no']), 'IF(A1>1,"yes","no")');
  assert.equal(call('LEFT', ['apple pie', '2']), 'LEFT("apple pie",2)');
  assert.equal(call('SUBSTITUTE', ['x"y', 'x', 'z']), 'SUBSTITUTE("x""y","x","z")');
  assert.equal(call('IF', ['A1>1', 'Rate', '"x"']), 'IF(A1>1,Rate,"x")');
  // (A number, logical, reference or array box keeps what is typed, as does a LAMBDA's.)
  assert.equal(call('SUM', ['abc']), 'SUM(abc)');
  assert.equal(call('VLOOKUP', ['abc', 'A1:B3', '2', 'abc']), 'VLOOKUP("abc",A1:B3,2,abc)');
  assert.equal(call('LAMBDA', ['x', 'x+1']), 'LAMBDA(x,x+1)');
  assert.equal(call('TEXTJOIN', [',', 'TRUE', 'a', 'b', 'c']), 'TEXTJOIN(",",TRUE,"a","b","c")');
  for (const f of ['COUNTIF(A1:A3,">1")', 'IF(A1>1,"yes","no")', 'SUBSTITUTE("x""y","x","z")']) assert.ok(parseInput(`=${f}`).formula, f);
});

test('Function Arguments hints: every argument has a kind', () => {
  const kinds = new Set(['number', 'text', 'logical', 'reference', 'array', 'any']);
  for (const f of catalog()) for (const b of argBoxes(f, Array(4).fill('x'))) assert.ok(kinds.has(argKind(f.name, b.name)), `${f.name} ${b.name}`);
  const k = (f, a) => argKind(f, a);
  assert.deepEqual([k('SUM', 'number3'), k('LEFT', 'text'), k('IF', 'logical_test'), k('IF', 'value_if_true'), k('SUMIF', 'range'), k('VLOOKUP', 'table_array')],
    ['number', 'text', 'logical', 'any', 'reference', 'array']);
  assert.deepEqual([k('DSUM', 'criteria'), k('COUNTIF', 'criteria'), k('INDIRECT', 'a1'), k('HEX2DEC', 'number')], ['reference', 'any', 'logical', 'text']);
});
