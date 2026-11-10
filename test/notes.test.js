import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { readNotes } from '../src/ooxml/notes.js';

// notes.xlsx is written as Excel 365 writes notes and threaded comments (test/fixtures/make-fixtures.mjs); Excel opens
// it and reads notes on B2 and D4 by Avery Chen and a thread on C3 with Sam Rivera's reply (docs/NOTES.md).
test('notes and threaded comments read as Excel shows them', () => {
  const bytes = new Uint8Array(fs.readFileSync('test/fixtures/notes.xlsx'));
  const wb = Workbook.fromBytes(bytes);
  const notes = readNotes(wb, 0);
  assert.deepEqual([...notes.keys()].sort(), ['1,1', '2,2', '3,3']);
  const b2 = notes.get('1,1');
  assert.equal(b2.kind, 'note');
  assert.equal(b2.author, 'Avery Chen');
  assert.equal(b2.text, 'Avery Chen:\nCheck this total', 'the text as Excel gives it');
  assert.deepEqual(b2.runs.map((r) => r.bold), [true, false], 'the author in bold');
  assert.equal(notes.get('3,3').text, 'Avery Chen:\nFirst line\nSecond line');
  const c3 = notes.get('2,2');
  assert.equal(c3.kind, 'thread', 'the threaded comment, not its placeholder note');
  assert.deepEqual(c3.items.map((x) => [x.author, x.text]), [['Avery Chen', 'Is this the right total?'], ['Sam Rivera', 'Yes, checked against the invoice.']]);
  assert.equal(c3.items[1].date, '2026-01-02T14:05:00.00');
  // (Read only: saved unchanged, every part byte for byte.)
  wb.setCell(0, 4, 0, 'edited');
  const again = Workbook.fromBytes(wb.toBytes());
  assert.deepEqual([...readNotes(again, 0).keys()].sort(), ['1,1', '2,2', '3,3'], 'kept on save');
});
