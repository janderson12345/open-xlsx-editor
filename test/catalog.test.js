import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import '../src/calc/evaluate.js';
import { FUNCTIONS } from '../src/calc/functions.js';
import { catalog, CATEGORIES, functionInfo, syntaxOf, searchFunctions, DEFAULT_RECENT } from '../src/calc/catalog.js';
import { callAt, argIndex } from '../src/editor/fntip.js';

test('Insert Function lists every function the engine works out, in Excel\'s categories, and only those', () => {
  // (Excel leaves DATEDIF out of its lists; ANCHORARRAY and SINGLE are the file's forms of D1# and @.)
  const hidden = new Set(['DATEDIF', 'ANCHORARRAY', 'SINGLE']);
  const listed = new Set(catalog().map((f) => f.name));
  assert.deepEqual(Object.keys(FUNCTIONS).filter((n) => !hidden.has(n) && !listed.has(n)), []);
  assert.deepEqual([...listed].filter((n) => !FUNCTIONS[n] || hidden.has(n)), []);
  for (const f of catalog()) assert.ok(CATEGORIES.includes(f.category), f.name);
  // (A few as Microsoft's "Excel functions (by category)" has them.)
  const cat = (n) => functionInfo(n).category;
  assert.deepEqual(['PMT', 'XIRR', 'EDATE', 'AGGREGATE', 'T.TEST', 'XLOOKUP', 'TEXTJOIN', 'LAMBDA', 'ISOMITTED', 'BITAND', 'TDIST', 'STDEV'].map(cat),
    ['Financial', 'Financial', 'Date & Time', 'Math & Trig', 'Statistical', 'Lookup & Reference', 'Text', 'Logical', 'Information', 'Engineering', 'Compatibility', 'Compatibility']);
  assert.ok(DEFAULT_RECENT.every((n) => functionInfo(n)));
});

test('syntax with Excel\'s argument names; search by name', () => {
  assert.equal(syntaxOf('sum'), 'SUM(number1, [number2], ...)');
  assert.equal(syntaxOf('XLOOKUP'), 'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])');
  assert.equal(syntaxOf('PI'), 'PI()');
  assert.equal(syntaxOf('NOPE'), '');
  assert.deepEqual(searchFunctions('sum').slice(0, 3).map((f) => f.name), ['SUM', 'SUMIF', 'SUMIFS']);
  assert.ok(searchFunctions('dist').some((f) => f.name === 'T.DIST'));
  assert.deepEqual(searchFunctions('  '), []);
});

test('the screen tip: the call the caret is in, and which argument', () => {
  assert.deepEqual(callAt('=SUM(', 5), { name: 'SUM', arg: 0 });
  assert.deepEqual(callAt('=SUM(A1,', 8), { name: 'SUM', arg: 1 });
  assert.deepEqual(callAt('=IF(A1>0,ROUND(B1,', 18), { name: 'ROUND', arg: 1 });
  assert.deepEqual(callAt('=IF(A1>0,ROUND(B1,2),', 21), { name: 'IF', arg: 2 });
  assert.deepEqual(callAt('=CONCAT("a,(b",', 15), { name: 'CONCAT', arg: 1 });
  assert.deepEqual(callAt("=SUM('My (Data)'!A1,", 20), { name: 'SUM', arg: 1 });
  assert.deepEqual(callAt('=_xlfn.XLOOKUP(', 15), { name: 'XLOOKUP', arg: 0 });
  assert.equal(callAt('=SUM({1,2', 9), null);
  assert.equal(callAt('=SUM("a,', 8), null);
  assert.equal(callAt('SUM(', 4), null);
  assert.equal(callAt('=A1+', 4), null);
  const sum = functionInfo('SUM').args;
  assert.deepEqual([0, 1, 2, 5].map((k) => argIndex(sum, k)), [0, 1, 1, 1]);
  const sumifs = functionInfo('SUMIFS').args;
  assert.deepEqual([0, 1, 2, 3, 4, 5].map((k) => argIndex(sumifs, k)), [0, 1, 2, 1, 2, 1]);
  assert.equal(argIndex(functionInfo('ROUND').args, 2), -1);
});
