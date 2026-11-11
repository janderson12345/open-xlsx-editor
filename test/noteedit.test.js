import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { readNotes } from '../src/ooxml/notes.js';
import { setNote, deleteNote } from '../src/ooxml/noteedit.js';

const open = (f) => Workbook.fromBytes(new Uint8Array(fs.readFileSync(`test/fixtures/${f}`)));
const again = (wb) => Workbook.fromBytes(wb.toBytes());

test('a new note on a sheet with none: the comments part, the VML box, the relationships and legacyDrawing made', () => {
  const wb = open('widths.xlsx');
  setNote(wb, 0, 4, 2, 'Check this width', 'Avery Chen');
  const back = again(wb);
  const n = readNotes(back, 0).get('4,2');
  assert.equal(n.kind, 'note');
  assert.equal(n.author, 'Avery Chen');
  assert.deepEqual(n.runs.map((r) => [r.text, r.bold]), [['Avery Chen:', true], ['\nCheck this width', false]], 'as Excel\'s New Note writes it');
  const sheet = back.pkg.text(back.sheets[0].path);
  assert.ok(/<legacyDrawing\b[^>]*\sr:id="[^"]+"\/><\/worksheet>/.test(sheet), 'legacyDrawing at the end of the sheet');
  assert.ok(/<Default Extension="vml"/.test(back.pkg.text('[Content_Types].xml')));
  assert.ok(/comments\+xml/.test(back.pkg.text('[Content_Types].xml')));
});

test("an existing note's text replaced, its author kept; another added; one deleted; a threaded comment deleted with its thread", () => {
  const wb = open('notes.xlsx');
  setNote(wb, 0, 1, 1, 'Totals checked', 'Someone Else');
  setNote(wb, 0, 5, 0, 'A new one', 'Sam Rivera');
  assert.ok(deleteNote(wb, 0, 3, 3));
  assert.ok(deleteNote(wb, 0, 2, 2), 'the threaded comment at C3');
  assert.equal(deleteNote(wb, 0, 9, 9), false, 'nothing there');
  const notes = readNotes(again(wb), 0);
  assert.deepEqual([...notes.keys()].sort(), ['1,1', '5,0']);
  assert.deepEqual(notes.get('1,1').runs.map((r) => r.text), ['Avery Chen:', '\nTotals checked']);
  assert.equal(notes.get('5,0').author, 'Sam Rivera');
  assert.ok(!/ref="C3"/.test(again(wb).pkg.text('xl/threadedComments/threadedComment1.xml')), 'its replies gone too');
});
