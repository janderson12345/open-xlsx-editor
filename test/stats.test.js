import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { selectionStats } from '../src/ooxml/stats.js';
import { area, build } from './steps.js';

// The status bar's Average, Count and Sum (stats.js), by the rules of Excel's status bar.
test('status bar: Average, Count and Sum of the numbers; Count of everything filled; one cell or one filled says nothing', () => {
  const { wb, si } = build({ cells: { A1: '1', A2: '2.5', A3: 'text', A4: 'TRUE', B1: '$1,000.50', B2: '5', C1: '=1/0' } });
  const s = selectionStats(wb, si, area('A1:A5'));
  assert.deepEqual([s.count, s.numbers, s.sum, s.average, s.text.sum, s.text.average], [4, 2, 3.5, 1.75, '3.5', '1.75']);
  assert.equal(selectionStats(wb, si, area('A1')), null);
  assert.equal(selectionStats(wb, si, area('A5:A9')), null);
  // (The number format of the first cell with a number: B1's currency.)
  assert.equal(selectionStats(wb, si, area('B1:B2')).text.sum, '$1,005.50');
  // (An error: Count alone.)
  assert.deepEqual(selectionStats(wb, si, area('A1:C1')), { count: 3, numbers: 2 });
  // (Hidden rows and columns left out.)
  wb.setLines(si, 'row', 1, 1, { hidden: true });
  assert.equal(selectionStats(wb, si, area('A1:A3')).sum, 1);
});
