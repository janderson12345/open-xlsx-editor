import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { justifyLines, justifySteps, JUSTIFY_REFUSED } from '../src/ooxml/justifyops.js';

// (Widths here by length: a line fits at `n` characters or fewer. Excel's own widths are tested in the browser,
// test/editing-links.mjs, on its lines in test/justify-excel.json.)
const upTo = (n) => (line) => line.length <= n;

test('paragraphs joined and broken at spaces, as Excel\'s Range.Justify', () => {
  // (Excel: "one", "two", "three", "four" in A1:A4 gave "one two", "three four".)
  assert.deepEqual(justifyLines(['one', 'two', 'three', 'four'], upTo(10)), ['one two', 'three four']);
  // (A word wider than the line has one of its own; Supercalifragilisticexpialidocious, then "and more".)
  assert.deepEqual(justifyLines(['Supercalifragilistic and more'], upTo(10)), ['Supercalifragilistic', 'and more']);
  // (Spaces at the start dropped; those beyond the one a break takes stay at the line's end: "  lead and   inner  ",
  // "next" gave "lead and  ", "inner  ", "next".)
  assert.deepEqual(justifyLines(['  lead and   inner  ', 'next'], upTo(10)), ['lead and  ', 'inner  ', 'next']);
});

test('blank rows kept after each paragraph, leading ones too', () => {
  // (Excel: alpha beta, gamma, blank, delta epsilon, zeta: three lines, the blank, three lines, the second paragraph a
  // row lower; "one", two blanks, "two", "three": the blanks kept; two blanks first, then "one", "two": still first.)
  assert.deepEqual(justifyLines(['alpha beta', 'gamma', null, 'delta epsilon', 'zeta'], upTo(7)), ['alpha', 'beta', 'gamma', null, 'delta', 'epsilon', 'zeta']);
  assert.deepEqual(justifyLines(['one', null, null, 'two', 'three'], upTo(10)), ['one', null, null, 'two three']);
  assert.deepEqual(justifyLines([null, null, 'one', 'two', null], upTo(10)), [null, null, 'one two']);
  assert.deepEqual(justifyLines([null, null], upTo(10)), []);
});

const book = (cells) => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'Justify');
  wb.apply(si, cells);
  return { wb, si };
};
const columnOf = (wb, si, c, n) => Array.from({ length: n }, (_, r) => wb.sheet(si).cells.get(r)?.get(c)?.v ?? null);

test('the first column rewritten, the rows it no longer needs cleared, the others left', () => {
  const { wb, si } = book([{ r: 0, c: 0, value: 'one' }, { r: 1, c: 0, value: 'two' }, { r: 0, c: 1, value: 'bee' }, { r: 1, c: 2, value: 5 }]);
  const got = justifySteps(wb, si, { r0: 0, c0: 0, r1: 2, c1: 2 }, upTo(10));
  assert.equal(got.extends, false);
  wb.apply(si, got.steps[0].contents);
  // (Excel, A1:C3: A1 "one two", A2 cleared, B1 "bee" and C2 5 as they were.)
  assert.deepEqual(columnOf(wb, si, 0, 3), ['one two', null, null]);
  assert.equal(wb.sheet(si).cells.get(0).get(1).v, 'bee');
  assert.equal(wb.sheet(si).cells.get(1).get(2).v, 5);
});

test('text going on below the selection, over what is there', () => {
  const { wb, si } = book([{ r: 0, c: 0, value: 'one two three four five' }, { r: 3, c: 0, value: 'below' }, { r: 4, c: 0, value: 7 }]);
  const got = justifySteps(wb, si, { r0: 0, c0: 0, r1: 1, c1: 0 }, upTo(4));
  assert.equal(got.extends, true);
  wb.apply(si, got.steps[0].contents);
  // (Excel wrote its lines over A4 "below" and A5's 7.)
  assert.deepEqual(columnOf(wb, si, 0, 5), ['one', 'two', 'three', 'four', 'five']);
});

test('a number, TRUE, an error or a formula in the first column refused, in Excel\'s words', () => {
  assert.equal(JUSTIFY_REFUSED, 'Cannot justify cells containing numbers or formulas.');
  for (const odd of [{ value: 42 }, { value: true }, { formula: '1+1' }, { formula: 'NA()' }, { formula: '"text"' }]) {
    const { wb, si } = book([{ r: 0, c: 0, value: 'alpha' }, { r: 1, c: 0, ...odd }]);
    assert.equal(justifySteps(wb, si, { r0: 0, c0: 0, r1: 1, c1: 0 }, upTo(10)).error, JUSTIFY_REFUSED, JSON.stringify(odd));
  }
  // (A number typed as text, '123, is text: "one 123".)
  const { wb, si } = book([{ r: 0, c: 0, value: 'one' }, { r: 1, c: 0, value: '123' }]);
  assert.equal(justifySteps(wb, si, { r0: 0, c0: 0, r1: 1, c1: 0 }, upTo(10)).steps[0].contents[0].value, 'one 123');
});
