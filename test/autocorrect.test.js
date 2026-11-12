import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { correction } from '../src/editor/autocorrect.js';
import { REPLACEMENTS } from '../src/editor/autocorrect-list.js';

// (The word before the end, corrected as AutoCorrect does when it ends; the text after it applied.)
const fixed = (text) => { const f = correction(text); return f ? text.slice(0, f.from) + f.text + text.slice(f.to) : text; };

test("AutoCorrect's replacement list: entries as Excel's own list has them (read over COM), the capital kept", () => {
  assert.equal(fixed('teh'), 'the');
  assert.equal(fixed('Teh'), 'The');
  assert.equal(fixed('we recieve'), 'we receive');
  assert.equal(fixed('i'), 'I');
  assert.equal(fixed('dont'), 'don’t');
  assert.equal(fixed('(c)'), '©');
  assert.equal(fixed('Total(tm)'), 'Total™');
  assert.equal(fixed('wait...'), 'wait…');
  assert.equal(fixed('teh.'), 'the.');
  // (An entry is the object's own: "constructor" isn't one.)
  assert.equal(fixed('constructor'), 'constructor');
  assert.ok(Object.values(REPLACEMENTS).every((v) => typeof v === 'string'));
});

test('AutoCorrect: days, TWo INitial CApitals, a sentence after a full stop, not the cell’s first word, not a formula', () => {
  assert.equal(fixed('monday'), 'Monday');
  assert.equal(fixed('THe'), 'The');
  assert.equal(fixed('TWO'), 'TWO');
  assert.equal(fixed('hello'), 'hello');
  assert.equal(fixed('Done. next'), 'Done. Next');
  assert.equal(fixed('Done. teh'), 'Done. The');
  assert.equal(fixed('=teh'), '=teh');
  assert.equal(fixed('=A1&" teh'), '=A1&" teh');
  assert.equal(fixed('-teh'), '-teh');
  assert.equal(fixed('-5 teh'), '-5 the');
  assert.equal(fixed('123'), '123');
});
