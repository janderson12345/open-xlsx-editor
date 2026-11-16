import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { fontsFromRuns, editAt, applyEdit, toggleFont, setFont, sizeAt, runsOf } from '../src/editor/richmarks.js';

/** `fonts` after `before` was edited into `after` (the selection before it `sel`, the caret after it `caret`). */
const typed = (fonts, before, after, sel, caret) => applyEdit(fonts, editAt(before, after, sel, caret));

test('part of a cell\'s text formatted while edited: toggled, kept in place as the text changes, turned into runs', () => {
  let fonts = toggleFont(Array(10).fill(null), 6, 10, 'b');
  assert.deepEqual(runsOf('Plain bold', fonts), [{ text: 'Plain ', font: null }, { text: 'bold', font: { b: true } }]);
  // (Typed at the start: the first character's font, plain here.)
  fonts = typed(fonts, 'Plain bold', 'Very Plain bold', [0, 0], 5);
  assert.deepEqual(runsOf('Very Plain bold', fonts), [{ text: 'Very Plain ', font: null }, { text: 'bold', font: { b: true } }]);
  // (Typed after the bold run: bold, the font before the caret.)
  fonts = typed(fonts, 'Very Plain bold', 'Very Plain bold!', [15, 15], 16);
  assert.deepEqual(runsOf('Very Plain bold!', fonts).map((r) => r.text), ['Very Plain ', 'bold!']);
  // (Toggled again over part of it: that part plain.)
  fonts = toggleFont(fonts, 11, 13, 'b');
  assert.deepEqual(runsOf('Very Plain bold!', fonts).map((r) => [r.text, !!r.font?.b]), [['Very Plain bo', false], ['ld!', true]]);
  // (A rich cell's runs as fonts, and back; Ctrl+U a single underline.)
  const runs = [{ text: 'H', font: null }, { text: '2', font: { i: true } }, { text: 'O', font: { u: 'single' } }];
  assert.deepEqual(runsOf('H2O', fontsFromRuns(runs)), runs);
  assert.deepEqual(runsOf('ab', toggleFont([null, null], 1, 2, 'u')), [{ text: 'a', font: null }, { text: 'b', font: { u: 'single' } }]);
  assert.equal(runsOf('plain', Array(5).fill(null)), null);
  // (Turned off where it was all on: the run's own font, nothing set, is the cell's.)
  assert.equal(runsOf('ab', toggleFont(toggleFont([null, null], 0, 2, 'i'), 0, 2, 'i')), null);
});

test('a run\'s size, colour, font and subscript follow the text as it is edited, not only bold and italic', () => {
  const red = { size: 16, color: 'FFFF0000', name: 'Courier New' };
  const sub = { vertAlign: 'subscript' };
  let fonts = fontsFromRuns([{ text: 'H', font: null }, { text: '2', font: sub }, { text: 'O ', font: null }, { text: 'RED', font: red }]);
  fonts = toggleFont(fonts, 2, 3, 'b');
  // (Typed inside the red run, then its first two letters typed over: red 16pt Courier New both times; typed after
  // the subscript 2: subscript.)
  fonts = typed(fonts, 'H2O RED', 'H2O RExD', [6, 6], 7);
  fonts = typed(fonts, 'H2O RExD', 'H2O zzxD', [4, 6], 6);
  fonts = typed(fonts, 'H2O zzxD', 'H2yO zzxD', [2, 2], 3);
  assert.deepEqual(runsOf('H2yO zzxD', fonts), [{ text: 'H', font: null }, { text: '2y', font: sub }, { text: 'O', font: { b: true } },
    { text: ' ', font: null }, { text: 'zzxD', font: red }]);
});

test('where the edit was: the caret and the selection before it place it where the text alone can\'t', () => {
  // ("b" typed between a and b of "ab": the new b is the second character, not the third.)
  assert.deepEqual(editAt('ab', 'abb', [1, 1], 2), { a: 1, b: 1, n: 1 });
  assert.deepEqual(editAt('ab', 'abb', null, null), { a: 2, b: 2, n: 1 });
  assert.deepEqual(editAt('abc', 'ac', [2, 2], 1), { a: 1, b: 2, n: 0 }); // (Backspace)
  assert.deepEqual(editAt('abc', 'ac', [1, 1], 1), { a: 1, b: 2, n: 0 }); // (Delete)
  assert.deepEqual(editAt('abcd', 'aXd', [1, 3], 2), { a: 1, b: 3, n: 1 }); // (a selection typed over)
  // (A composition updated: the selection known no better than the caret, the texts tell.)
  assert.deepEqual(editAt('abk', 'abか', [3, 3], 3), { a: 2, b: 3, n: 1 });
});

test('the font typed text takes is the one Excel\'s own insert gives it (test/richtype-excel.json, measured over COM)', () => {
  const cases = JSON.parse(fs.readFileSync(new URL('./richtype-excel.json', import.meta.url), 'utf8'));
  let checked = 0;
  for (const c of cases) {
    // (Text added after the last character, which the object model gives the cell's own font: not followed, see
    // richmarks.js applyEdit.)
    if (c.at[0] > c.text.length) continue;
    const font = (ch) => (ch.b ? { b: true } : ch.color === 255 ? { color: 'FFFF0000', size: 16 } : null);
    const marked = c.runs[0][2] === 'b' ? { b: true } : { color: 'FFFF0000', size: 16 };
    const fonts = [...c.text].map((_, i) => (i >= c.runs[0][0] - 1 && i < c.runs[0][0] - 1 + c.runs[0][1] ? marked : null));
    const a = c.at[0] - 1;
    const after = applyEdit(fonts, { a, b: a + c.at[1], n: c.add.length });
    assert.deepEqual(after.map((f) => (f ? JSON.stringify(f) : null)), c.chars.map((ch) => (font(ch) ? JSON.stringify(font(ch)) : null)), c.name);
    checked++;
  }
  assert.equal(checked, cases.length - 1);
});

test('a size, colour or font set on part of the text: those characters only, runs kept apart, a new font losing the old scheme', () => {
  const fonts = fontsFromRuns([{ text: 'ab', font: null }, { text: 'cd', font: { b: true, name: 'Aptos Narrow', scheme: 'minor', family: 2 } }]);
  let next = setFont(fonts, 1, 3, { size: 16 });
  assert.deepEqual(runsOf('abcd', next), [{ text: 'a', font: null }, { text: 'b', font: { size: 16 } },
    { text: 'c', font: { b: true, name: 'Aptos Narrow', scheme: 'minor', family: 2, size: 16 } }, { text: 'd', font: { b: true, name: 'Aptos Narrow', scheme: 'minor', family: 2 } }]);
  next = setFont(next, 0, 4, { color: { rgb: 'FFFF0000' } });
  assert.ok(next.every((f) => f.color.rgb === 'FFFF0000'));
  next = setFont(next, 2, 4, { name: 'Courier New' });
  assert.deepEqual([next[2].name, next[2].scheme, next[2].family, next[2].b], ['Courier New', null, null, true]);
  assert.equal(next[0].name, undefined);
  // (Grow Font steps from the first selected character's size, the cell's where it has none.)
  assert.equal(sizeAt(next, 1, 11), 16);
  assert.equal(sizeAt(next, 0, 11), 11);
  // (Nothing selected: nothing changes.)
  assert.equal(setFont(fonts, 2, 2, { size: 9 }), fonts);
});
