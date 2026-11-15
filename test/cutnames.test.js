import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { moveRefs } from '../src/calc/tokens.js';
import { defineName, deleteName, moveNamesForCut } from '../src/ooxml/names.js';
import { build } from './steps.js';

// Excel's names and formulas after B2:C3 of Sheet1 was cut (Range.Cut in a hidden Excel; docs/CLIPBOARD.md), to E5,
// to another sheet's E5, and to B3 (down one row, over itself).
const CUT = { sheet: 'Sheet1', r0: 1, c0: 1, r1: 2, c1: 2 };
const TO = { 'to E5': { sheet: 'Sheet1', dr: 3, dc: 3 }, 'to Data!E5': { sheet: 'Data', dr: 3, dc: 3 }, 'to B3': { sheet: 'Sheet1', dr: 1, dc: 0 } };
const NAMES = [['Qty', '$B$2'], ['Block', '$B$2:$C$3'], ['Part', '$B$2:$B$5'], ['Far', '$H$9'], ['Over', '$E$5'], ['Rel', 'B2'], ['Wide', '$A$1:$F$8']];
const EXCEL = {
  'to E5': {
    names: { Qty: 'Sheet1!$E$5', Block: 'Sheet1!$E$5:$F$6', Part: 'Sheet1!$B$2:$B$5', Far: 'Sheet1!$H$9', Over: 'Sheet1!#REF!', Rel: 'Sheet1!B2', Wide: 'Sheet1!$A$1:$F$8', Cnst: '7' },
    formulas: ['SUM(B2:B5)', '#REF!', 'SUM(E5:E7)', 'SUM(D4:F6)'],
  },
  'to Data!E5': {
    names: { Qty: 'Data!$E$5', Block: 'Data!$E$5:$F$6', Part: 'Sheet1!$B$4:$B$5', Far: 'Sheet1!$H$9', Over: 'Sheet1!$E$5', Rel: 'Sheet1!B2', Wide: 'Sheet1!$A$1:$F$8', Cnst: '7' },
    formulas: ['SUM(B4:B5)', 'E5', 'SUM(E5:E7)', 'SUM(D4:F6)'],
  },
  'to B3': {
    names: { Qty: 'Sheet1!$B$3', Block: 'Sheet1!$B$3:$C$4', Part: 'Sheet1!$B$3:$B$5', Far: 'Sheet1!$H$9', Over: 'Sheet1!$E$5', Rel: 'Sheet1!B2', Wide: 'Sheet1!$A$1:$F$8', Cnst: '7' },
    formulas: ['SUM(B3:B5)', 'E5', 'SUM(E5:E7)', 'SUM(D4:F6)'],
  },
};

test('formulas after a cut, as Excel\'s: a range partly cut, the cells left and moved when they make a rectangle', () => {
  for (const [where, to] of Object.entries(TO)) {
    const got = ['SUM(B2:B5)', 'E5', 'SUM(E5:E7)', 'SUM(D4:F6)'].map((f) => moveRefs(f, { here: 'Sheet1', from: CUT, to }));
    assert.deepEqual(got, EXCEL[where].formulas, where);
  }
  // (The case already in the docs: A1:A3 with A1 moved onto A3, =SUM(A1:A3) is =SUM(A2:A3).)
  assert.equal(moveRefs('SUM(A1:A3)', { here: 'S', from: { sheet: 'S', r0: 0, c0: 0, r1: 0, c1: 0 }, to: { sheet: 'S', dr: 2, dc: 0 } }), 'SUM(A2:A3)');
});

test('defined names after a cut, as Excel\'s: moved, shrunk, #REF! or left', () => {
  for (const [where, to] of Object.entries(TO)) {
    const { wb } = build({ cells: { A1: '1' } });
    for (const d of [...wb.definedNames]) deleteName(wb, { name: d.name, sheet: d.sheet ?? null });
    for (const [n, ref] of NAMES) defineName(wb, { name: n, refersTo: `=Sheet1!${ref}` });
    defineName(wb, { name: 'Cnst', refersTo: '=7' });
    moveNamesForCut(wb, CUT, to);
    assert.deepEqual(Object.fromEntries(wb.definedNames.map((d) => [d.name, d.text])), EXCEL[where].names, where);
  }
});
