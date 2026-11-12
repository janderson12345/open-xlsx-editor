import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { spawnSync } from 'node:child_process';
import { Workbook } from '../src/ooxml/workbook.js';
import { addThread, replyThread, editThread, resolveThread } from '../src/ooxml/threadedit.js';
import { deleteNote } from '../src/ooxml/noteedit.js';
import { readNotes } from '../src/ooxml/notes.js';

// Threaded comments written as Excel 365 writes them: test/threads-excel.json holds Excel's parts after B2 was given a
// thread and two replies, D4 a thread then deleted, and F6 a thread (tools/excel-compare/excel-threads.ps1; the
// person Excel signed them with replaced by a made-up one). The same steps here, each part compared with Excel's,
// the made-up ids and times aside.
const excel = JSON.parse(fs.readFileSync('test/threads-excel.json', 'utf8')).parts;
const open = () => Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
const AUTHOR = 'Avery Chen';

/** Ids, times and shape numbers out of the way: each id by the order it first appears. */
function bare(xml) {
  const seen = new Map();
  return xml.replace(/\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}/g, (g) => { if (!seen.has(g)) seen.set(g, `{id${seen.size}}`); return seen.get(g); })
    .replace(/dT="[^"]+"/g, 'dT').replace(/_x0000_s\d+/g, '_x0000_s').replace(/>\s+</g, '><').replace(/\s+/g, ' ');
}
const partOf = (files, start) => strFromU8(files[Object.keys(files).find((k) => k.startsWith(start))]);

test('a thread, its replies, one deleted and another: each part as Excel writes it (test/threads-excel.json)', () => {
  const wb = open();
  addThread(wb, 0, 1, 1, 'Check these totals', AUTHOR);
  replyThread(wb, 0, 1, 1, 'Done & checked <twice>', AUTHOR);
  replyThread(wb, 0, 1, 1, 'Thanks\nlooks right', AUTHOR);
  addThread(wb, 0, 3, 3, 'Delete me', AUTHOR);
  deleteNote(wb, 0, 3, 3);
  addThread(wb, 0, 5, 5, 'A second thread', AUTHOR);
  const files = unzipSync(wb.toBytes());
  const sheetRels = strFromU8(files['xl/worksheets/_rels/sheet1.xml.rels']);
  const comments = /Target="\.\.\/(comments\d*\.xml)"/.exec(sheetRels)[1];
  const vml = /Target="\.\.\/drawings\/(vmlDrawing\d*\.vml)"/.exec(sheetRels)[1];
  assert.equal(bare(partOf(files, 'xl/threadedComments/')), bare(excel['xl/threadedComments/threadedComment1.xml']));
  // (Excel's person is someone signed in, ours someone who isn't: their user and provider ids differ.)
  assert.equal(bare(strFromU8(files['xl/persons/person.xml'])).replace(/userId="[^"]*" providerId="[^"]*"/, ''), bare(excel['xl/persons/person.xml']).replace(/userId="[^"]*" providerId="[^"]*"/, ''));
  assert.equal(bare(strFromU8(files[`xl/${comments}`])), bare(excel['xl/comments1.xml']));
  // (Excel breaks the VML's lines where it likes: spaces after ; and > aside.)
  const shapes = (x) => bare(x).replace(/;\s+/g, ';').replace(/>\s+/g, '>').match(/<v:shape .*?<\/v:shape>/g);
  assert.deepEqual(shapes(strFromU8(files[`xl/drawings/${vml}`])), shapes(excel['xl/drawings/vmlDrawing1.vml']));
  // The relationships and content types Excel has for them.
  for (const type of ['threadedComment', 'relationships/comments', 'relationships/vmlDrawing']) assert.match(sheetRels, new RegExp(type));
  assert.match(strFromU8(files['xl/_rels/workbook.xml.rels']), /relationships\/person" Target="persons\/person.xml"/);
  const ct = strFromU8(files['[Content_Types].xml']);
  for (const t of ['vnd.ms-excel.threadedcomments+xml', 'vnd.ms-excel.person+xml', 'spreadsheetml.comments+xml']) assert.ok(ct.includes(t), t);
  fs.mkdirSync('test/out/threads', { recursive: true });
  fs.writeFileSync('test/out/threads/threads.xlsx', wb.toBytes());
  const out = spawnSync(process.execPath, ['tools/check-ooxml.mjs', 'test/out/threads/threads.xlsx'], { encoding: 'utf8' });
  assert.equal(out.status, 0, out.stdout + out.stderr);
});

test('read back: the thread with its replies; a reply edited and deleted; a second author; a cell with a note refused', () => {
  const wb = open();
  addThread(wb, 0, 1, 1, 'First', 'Avery Chen');
  replyThread(wb, 0, 1, 1, 'Second', 'Sam Lee');
  replyThread(wb, 0, 1, 1, 'Third', 'Avery Chen');
  editThread(wb, 0, 1, 1, 1, 'Second, edited');
  editThread(wb, 0, 1, 1, 2, null);
  const again = Workbook.fromBytes(wb.toBytes());
  const t = readNotes(again, 0).get('1,1');
  assert.deepEqual(t.items.map((x) => `${x.author}: ${x.text}`), ['Avery Chen: First', 'Sam Lee: Second, edited']);
  assert.equal((strFromU8(unzipSync(wb.toBytes())['xl/persons/person.xml']).match(/<person /g) || []).length, 2);
  assert.throws(() => addThread(wb, 0, 1, 1, 'Again'), /reply to it/);
  assert.throws(() => replyThread(wb, 0, 7, 7, 'Nothing here'), /No comment/);
});

test("Excel's reading of the threads written here: cells, authors, texts and replies in order (test/threads-readback-excel.json)", () => {
  // (excel-threads-check.ps1 opened test/out/threads/threads.xlsx, made by the first test, in Excel.)
  const back = JSON.parse(fs.readFileSync('test/threads-readback-excel.json', 'utf8'));
  assert.deepEqual(back.threads.map((t) => [t.cell, t.author, t.text, t.replies.map((r) => `${r.author}: ${r.text}`)]), [
    ['B2', AUTHOR, 'Check these totals', [`${AUTHOR}: Done & checked <twice>`, `${AUTHOR}: Thanks\nlooks right`]],
    ['F6', AUTHOR, 'A second thread', []],
  ]);
  // (Excel 365 shows the threads, not their placeholder notes.)
  assert.equal(back.legacyComments, 0);
});

test('a thread resolved as Excel writes it: done="1" on its first comment, after its id; no reply until reopened (measured over COM)', () => {
  const wb = open();
  const when = new Date(Date.UTC(2026, 10, 12, 21, 14, 35));
  addThread(wb, 0, 1, 1, 'first', AUTHOR, { when });
  replyThread(wb, 0, 1, 1, 'a reply', AUTHOR, { when });
  resolveThread(wb, 0, 1, 1, true);
  const part = () => { const p = Object.keys(wb.pkg.files).find((k) => k.startsWith('xl/threadedComments/')); return wb.pkg.text(p); };
  assert.match(part(), /<threadedComment ref="B2" dT="[^"]+" personId="[^"]+" id="\{[0-9A-F-]+\}" done="1"><text>first<\/text>/);
  assert.doesNotMatch(part().replace(/^[\s\S]*?done="1"/, ''), /done=/);
  assert.equal(readNotes(wb, 0).get('1,1').resolved, true);
  assert.equal(replyThread(wb, 0, 1, 1, 'reply on resolved', AUTHOR, { when }), null);
  assert.doesNotMatch(part(), /reply on resolved/);
  resolveThread(wb, 0, 1, 1, false);
  assert.doesNotMatch(part(), /done=/);
  assert.equal(readNotes(wb, 0).get('1,1').resolved, false);
  assert.ok(replyThread(wb, 0, 1, 1, 'after reopening', AUTHOR, { when }));
});
