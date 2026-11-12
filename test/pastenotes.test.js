import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { snapshot, pasteNotes } from '../src/ooxml/clipops.js';
import { setNote } from '../src/ooxml/noteedit.js';
import { readNotes } from '../src/ooxml/notes.js';

// (Excel 365 over COM: B2:C3 with a note on B2 copied, PasteSpecial(xlPasteComments) onto E5 where E5 and F6 had notes
// of their own: E5 the copy's note, F6's taken away; Paste and PasteSpecial(xlPasteAll) bring the note too; then E5:F6
// copied onto empty T5 brings only E5's.)
const book = () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'Notes');
  for (const [ref, v] of [['B2', 1], ['C2', 2], ['B3', 3], ['C3', 4], ['E5', 'old'], ['F6', 'f6']]) wb.setCell(si, ref, v);
  setNote(wb, si, 1, 1, 'note on b2', 'Avery Chen');
  setNote(wb, si, 4, 4, 'old note e5', 'Avery Chen');
  setNote(wb, si, 5, 5, 'old note f6', 'Avery Chen');
  return { wb, si };
};
const notesOf = (wb, si) => Object.fromEntries([...readNotes(wb, si)].map(([k, n]) => [k, n.runs.map((x) => x.text).join('')]));

test("a paste's notes as Excel's: each pasted cell ends with the copy's note there or none", () => {
  const { wb, si } = book();
  const copy = snapshot(wb, si, { ar: 1, ac: 1, r0: 1, c0: 1, r1: 2, c1: 2 });
  assert.equal(copy.notes.length, 1);
  assert.equal(pasteNotes(wb, si, { r0: 4, c0: 4, r1: 5, c1: 5 }, copy), true);
  const now = notesOf(wb, si);
  assert.match(now['4,4'], /note on b2$/);
  assert.equal(now['5,5'], undefined);
  assert.match(now['1,1'], /note on b2$/);
  assert.equal(readNotes(wb, si).get('4,4').author, 'Avery Chen');
  // (Pasted again over the same cells: nothing changes.)
  assert.equal(pasteNotes(wb, si, { r0: 4, c0: 4, r1: 5, c1: 5 }, copy), false);
  // (E5:F6 copied onto T5:U6: only T5 gets a note.)
  const again = snapshot(wb, si, { ar: 4, ac: 4, r0: 4, c0: 4, r1: 5, c1: 5 });
  pasteNotes(wb, si, { r0: 4, c0: 19, r1: 5, c1: 20 }, again);
  const t = notesOf(wb, si);
  assert.deepEqual(['4,19', '4,20', '5,19', '5,20'].map((k) => !!t[k]), [true, false, false, false]);
  // (Repeated over a selection twice the copy's size.)
  pasteNotes(wb, si, { r0: 10, c0: 0, r1: 13, c1: 1 }, copy);
  assert.deepEqual(['10,0', '12,0', '10,1', '11,0'].map((k) => !!notesOf(wb, si)[k]), [true, true, false, false]);
});
