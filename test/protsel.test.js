import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { build } from './steps.js';
import { protectedStep, canSelect, protectSheet } from '../src/ooxml/protection.js';
import { refName, parseRef } from '../src/ooxml/refs.js';

// A protected sheet's keys and clicks, as Excel's (its window, 15 November 2026): B2, C2, B3, C3 and E5 unlocked, the
// rest locked; the active cell read over COM after each key.
const setup = (selectLocked, selectUnlocked) => {
  const { wb, si } = build({ cells: { A1: 'locked', B2: 'B2', C2: 'C2', B3: 'B3', C3: 'C3', E5: 'E5' } });
  const open = wb.styleWith(0, { protection: { locked: false } });
  wb.setStyles(si, ['B2', 'C2', 'B3', 'C3', 'E5'].map((a) => ({ ...parseRef(a), s: open })));
  protectSheet(wb, si, { selectLocked, selectUnlocked });
  return { wb, si };
};
const walk = (wb, si, start, keys) => {
  let at = parseRef(start);
  return keys.map(([dr, dc, how]) => { const p = protectedStep(wb, si, at, dr, dc, how); if (p) at = p; return p === undefined ? 'usual' : refName(at.r, at.c); });
};
const R = [0, 1, 'arrow']; const L = [0, -1, 'arrow']; const D = [1, 0, 'arrow']; const U = [-1, 0, 'arrow'];
const TAB = [0, 1, 'tab']; const ENTER = [1, 0, 'enter'];

test('Select locked cells off: arrows go through the unlocked cells, rows or columns, wrapping; Tab and Enter too', () => {
  const { wb, si } = setup(false, true);
  assert.deepEqual(walk(wb, si, 'B2', [R, R, R, R, R]), ['C2', 'B3', 'C3', 'E5', 'B2'], 'Right');
  assert.deepEqual(walk(wb, si, 'B2', [D, D]), ['B3', 'C2'], 'Down');
  assert.deepEqual(walk(wb, si, 'C2', [L, L]), ['B2', 'E5'], 'Left');
  assert.deepEqual(walk(wb, si, 'E5', [U]), ['C3'], 'Up');
  assert.deepEqual(walk(wb, si, 'C3', [TAB, ENTER]), ['E5', 'B2'], 'Tab, then Enter');
  assert.equal(canSelect(wb, si, 0, 0), false, 'a click on locked A1: nothing');
  assert.equal(canSelect(wb, si, 2, 2), true, 'on unlocked C3: selected');
});

test('nothing selectable: arrows and clicks do nothing; Tab in rows, Enter down columns', () => {
  const { wb, si } = setup(false, false);
  assert.deepEqual(walk(wb, si, 'B2', [R]), ['B2'], 'Right stays');
  assert.equal(canSelect(wb, si, 4, 4), false, 'a click on E5: nothing');
  assert.deepEqual(walk(wb, si, 'B2', [TAB, ENTER]), ['C2', 'C3'], 'Tab, then Enter');
});

test('everything selectable: arrows and Enter as usual; Tab through the unlocked cells', () => {
  const { wb, si } = setup(true, true);
  assert.deepEqual(walk(wb, si, 'E5', [TAB]), ['B2'], 'Tab from E5 to B2');
  assert.deepEqual(walk(wb, si, 'A1', [R, ENTER]), ['usual', 'usual'], 'Right and Enter move as usual');
});
