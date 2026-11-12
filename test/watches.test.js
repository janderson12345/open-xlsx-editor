import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { Workbook } from '../src/ooxml/workbook.js';
import { readWatches, addWatches, deleteWatches } from '../src/ooxml/watches.js';

// Excel's file for the same watches (Application.Watches.Add of B2, then A1, then another sheet's D5): each sheet's
// cellWatches in the order added, after pageMargins.
test('watched cells saved as Excel saves them, read back, deleted', () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/formulas.xlsx')));
  const a = wb.addSheet(wb.sheets.length);
  const b = wb.addSheet(wb.sheets.length);
  addWatches(wb, a, [{ r: 1, c: 1 }, { r: 0, c: 0 }]);
  assert.equal(addWatches(wb, a, [{ r: 1, c: 1 }]), 0);
  addWatches(wb, b, [{ r: 4, c: 3 }]);
  const files = unzipSync(wb.toBytes());
  assert.match(strFromU8(files[wb.sheets[a].path]), /<pageMargins[^>]*\/><cellWatches><cellWatch r="B2"\/><cellWatch r="A1"\/><\/cellWatches>/);
  assert.match(strFromU8(files[wb.sheets[b].path]), /<cellWatches><cellWatch r="D5"\/><\/cellWatches>/);
  const again = Workbook.fromBytes(wb.toBytes());
  assert.deepEqual(readWatches(again), [{ si: a, r: 1, c: 1 }, { si: a, r: 0, c: 0 }, { si: b, r: 4, c: 3 }]);
  deleteWatches(again, [{ si: b, r: 4, c: 3 }]);
  assert.doesNotMatch(strFromU8(unzipSync(again.toBytes())[again.sheets[b].path]), /cellWatches/);
});
