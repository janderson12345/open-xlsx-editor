import { test } from 'node:test';
import assert from 'node:assert/strict';
import { marksFromRuns, shiftMarks, toggleMark, runsOf } from '../src/editor/richmarks.js';

test('part of a cell\'s text formatted while edited: marks toggled, kept in place as the text changes, turned into runs', () => {
  let marks = toggleMark([], 6, 10, 'b');
  assert.deepEqual(runsOf('Plain bold', marks), [{ text: 'Plain ', font: null }, { text: 'bold', font: { b: true } }]);
  // (Typed before the mark: it moves on; typed after: it stays.)
  marks = shiftMarks(marks, 'Plain bold', 'Very plain bold');
  assert.deepEqual(runsOf('Very plain bold', marks), [{ text: 'Very plain ', font: null }, { text: 'bold', font: { b: true } }]);
  marks = shiftMarks(marks, 'Very plain bold', 'Very plain bold!');
  assert.deepEqual(runsOf('Very plain bold!', marks).map((r) => r.text), ['Very plain ', 'bold', '!']);
  // (Toggled again over part of it: that part plain.)
  marks = toggleMark(marks, 11, 13, 'b');
  assert.deepEqual(runsOf('Very plain bold!', marks).map((r) => [r.text, !!r.font?.b]), [['Very plain bo', false], ['ld', true], ['!', false]]);
  // (A rich cell's runs as marks, and back.)
  const runs = [{ text: 'H', font: null }, { text: '2', font: { i: true } }, { text: 'O', font: { u: 'single' } }];
  assert.deepEqual(runsOf('H2O', marksFromRuns(runs)), [{ text: 'H', font: null }, { text: '2', font: { i: true } }, { text: 'O', font: { u: true } }]);
  assert.equal(runsOf('plain', []), null);
});
