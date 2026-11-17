import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { readNotes } from '../src/ooxml/notes.js';
import { setNoteVisible, setAllNotesVisible } from '../src/ooxml/noteedit.js';
import { convertNotes, noteCommentText, NO_PERSON } from '../src/ooxml/noteconvert.js';
import { convertInput } from '../tools/excel-compare/notes-convert-ours.mjs';

// Show/Hide Note, Show All Notes and Convert to Comments as Excel 365 writes them: test/notes-visible-excel.json holds
// what Excel saved and read over COM (tools/excel-compare/excel-notes-visible.ps1: Comment.Visible,
// Application.DisplayCommentIndicator, Workbook.ConvertComments), test/notes-readback-excel.json what Excel read in the
// workbooks written here (notes-visible-ours.mjs, excel-notes-readback.ps1).
const json = (f) => JSON.parse(fs.readFileSync(f, 'utf8').replace(/^\uFEFF/, ''));
/** (PowerShell's ConvertTo-Json writes a list of one as the one.) */
const list = (x) => (Array.isArray(x) ? x : [x]);
const excel = json('test/notes-visible-excel.json');
const readback = json('test/notes-readback-excel.json');
const open = (f) => Workbook.fromBytes(new Uint8Array(fs.readFileSync(`test/fixtures/${f}`)));
const again = (wb) => Workbook.fromBytes(wb.toBytes());
const VML = 'xl/drawings/vmlDrawing1.vml';
const shapes = (vml) => [...vml.matchAll(/<v:shape\b[\s\S]*?<\/v:shape>/g)].map((m) => m[0]);
const shapeAt = (vml, r, c) => shapes(vml).find((s) => new RegExp(`<x:Row>${r}</x:Row>`).test(s) && new RegExp(`<x:Column>${c}</x:Column>`).test(s));

test("Show/Hide Note on Excel's own shapes: shown as Excel's Comment.Visible writes it, hidden back to the bytes it had", () => {
  // (Excel's: B2 shown, D4 hidden, F6 shown and hidden again; and the same three with Show All Notes on, all shown.)
  const some = excel.visible.parts[VML];
  const all = excel.showall.parts[VML];
  const wb = open('notes.xlsx');
  wb.pkg.setText(VML, some);
  assert.ok(setNoteVisible(wb, 0, 3, 3, true));
  assert.ok(setNoteVisible(wb, 0, 5, 5, true));
  assert.equal(wb.pkg.text(VML), all, 'D4 and F6 shown: Excel\'s drawing with all three shown, byte for byte');
  assert.equal(setNoteVisible(wb, 0, 1, 1, true), false, 'B2 shown already');
  assert.ok(setNoteVisible(wb, 0, 3, 3, false));
  assert.ok(setNoteVisible(wb, 0, 5, 5, false));
  assert.equal(wb.pkg.text(VML), some, 'hidden again: Excel\'s, byte for byte');
  assert.equal(setNoteVisible(wb, 0, 9, 9, true), false, 'no note there');
});

test('a note shown here opens shown (x:Visible read back), the rest of its shape as it was; Excel reads it shown', () => {
  const wb = open('notes.xlsx');
  const before = wb.pkg.text(VML);
  assert.ok(setNoteVisible(wb, 0, 1, 1, true));
  const back = again(wb);
  assert.equal(readNotes(back, 0).get('1,1').visible, true);
  assert.equal(readNotes(back, 0).get('3,3').visible, false);
  const b2 = shapeAt(back.pkg.text(VML), 1, 1);
  assert.ok(/visibility:visible/.test(b2) && /<x:Column>1<\/x:Column><x:Visible\/><\/x:ClientData>/.test(b2), b2);
  assert.equal(shapeAt(back.pkg.text(VML), 3, 3), shapeAt(before, 3, 3), "D4's shape untouched");
  assert.ok(setNoteVisible(wb, 0, 1, 1, false));
  assert.equal(wb.pkg.text(VML), before, 'hidden again: the drawing as it was');
  const seen = list(readback['shown.xlsx'])[0];
  assert.deepEqual(seen.notes.map((n) => [n.cell, n.visible]), [['B2', true], ['D4', false]], 'Excel: B2 shown, D4 (shown and hidden here) not');
});

test("Show All Notes: every note shown, saved so; off, every note hidden; the thread's placeholder left", () => {
  const wb = open('notes.xlsx');
  const before = wb.pkg.text(VML);
  setNoteVisible(wb, 0, 1, 1, true);
  assert.equal(setAllNotesVisible(wb, true), 1, 'D4 (B2 shown already)');
  const notes = readNotes(again(wb), 0);
  assert.ok(notes.get('1,1').visible && notes.get('3,3').visible);
  assert.equal(shapeAt(wb.pkg.text(VML), 2, 2), shapeAt(before, 2, 2), "C3's thread placeholder untouched");
  assert.equal(setAllNotesVisible(wb, false), 2, 'off: both hidden, B2 shown before too, as Excel hides them');
  assert.equal(wb.pkg.text(VML), before);
  assert.deepEqual(list(readback['showall.xlsx'])[0].notes.map((n) => [n.cell, n.visible]), [['B2', true], ['D4', true]], 'Excel reads every note shown');
  // (Excel's own: on, every note Visible; off, every one hidden, F6 shown before it too.)
  assert.ok(list(excel.showall.during)[0].notes.every((n) => n.visible));
  assert.ok(list(excel.showall.afterPutBack)[0].notes.every((n) => !n.visible));
});

test("a converted note's text as Excel's: the first line dropped only when it is the note's author's name and a colon", () => {
  for (const book of [excel.convert, excel.ours]) {
    book.before.forEach((sheet, k) => {
      for (const n of sheet.notes) {
        const t = book.after[k].threads.find((x) => x.cell === n.cell);
        assert.equal(noteCommentText({ author: n.author, text: n.text }), t.items[0].text.replace(/\r\n/g, '\n'), `${sheet.name}!${n.cell}`);
      }
    });
  }
  assert.equal(noteCommentText({ author: 'Avery Chen', text: 'Avery Chen: inline' }), 'Avery Chen: inline', 'not a line of its own: kept (not measured)');
});

/** A part's threaded comments as [ref, dated, person (name, userId, provider), text, reply]. */
function threadsOf(wb, part, persons) {
  const people = new Map([...persons.matchAll(/<person displayName="([^"]*)" id="([^"]*)" userId="([^"]*)" providerId="([^"]*)"\/>/g)].map((m) => [m[2], `${m[1]}|${m[3]}|${m[4]}`]));
  people.set(NO_PERSON, 'none');
  return [...part.matchAll(/<threadedComment\b([^>]*)>([\s\S]*?)<\/threadedComment>/g)].map((m) => [
    /\sref="([^"]+)"/.exec(m[1])[1], /\sdT="/.test(m[1]), people.get(/\spersonId="([^"]+)"/.exec(m[1])[1]), /<text>([\s\S]*?)<\/text>/.exec(m[2])[1], /\sparentId="/.test(m[1]),
  ]);
}
/** A comments part's comments as [ref, its author, text], and whether each author is tc= the thread's id it has. */
function commentsOf(part) {
  const authors = [...part.matchAll(/<author>([\s\S]*?)<\/author>/g)].map((m) => m[1]);
  const list = [...part.matchAll(/<comment\b([^>]*)>([\s\S]*?)<\/comment>/g)].map((m) => ({ ref: /\sref="([^"]+)"/.exec(m[1])[1], author: authors[Number(/\sauthorId="(\d+)"/.exec(m[1])[1])], uid: /\sxr:uid="([^"]+)"/.exec(m[1])?.[1], text: /<t>([\s\S]*?)<\/t>/.exec(m[2])[1] }));
  return { list: list.map((c) => [c.ref, c.text]), authors, inOrder: authors.length === list.length && list.every((c, i) => authors[i] === c.author) };
}
/** A VML drawing's shapes as [id, z-index, row, column, anchor, style without z-index]. */
const boxesOf = (vml) => shapes(vml).map((s) => [/\sid="([^"]+)"/.exec(s)[1], /z-index:(\d+)/.exec(s)[1], /<x:Row>(\d+)/.exec(s)[1], /<x:Column>(\d+)/.exec(s)[1], /<x:Anchor>\s*([^<]*)</.exec(s)[1].trim(),
  /style='([^']*)'/.exec(s)[1].replace(/\s+/g, '').replace(/z-index:\d+;/, '')]);

test("Convert to Comments on every sheet, each part as Excel's ConvertComments writes it (ids aside)", () => {
  const wb = convertInput();
  const c3 = shapeAt(wb.pkg.text(VML), 2, 2);
  assert.equal(convertNotes(wb), 9, 'eight notes on the first sheet, one on the second');
  const parts = excel.ours.parts;
  const persons = wb.pkg.text('xl/persons/person.xml');
  for (const n of [1, 2]) {
    const ours = threadsOf(wb, wb.pkg.text(`xl/threadedComments/threadedComment${n}.xml`), persons);
    const theirs = threadsOf(null, parts[`xl/threadedComments/threadedComment${n}.xml`], parts['xl/persons/person.xml']);
    assert.deepEqual(ours, theirs, `threadedComment${n}.xml: by cell, no date, the note's author a person of their own, CR LF`);
    const oc = commentsOf(wb.pkg.text(`xl/comments${n}.xml`));
    const tc = commentsOf(parts[`xl/comments${n}.xml`]);
    // (C3's placeholder, there before, keeps its line breaks here; Excel saves every part again, with CR LF.)
    const lf = (l) => l.map(([ref, t]) => [ref, t.replace(/\r\n/g, '\n')]);
    assert.deepEqual(lf(oc.list), lf(tc.list), `comments${n}.xml: the placeholders by cell`);
    assert.ok(oc.list.filter(([ref]) => ref !== 'C3').every(([, t]) => t.includes('\r\nComment:\r\n    ')), 'the new ones with CR LF');
    assert.ok(oc.inOrder && tc.inOrder, `comments${n}.xml: one tc= author for each, in order`);
    assert.ok(/mc:Ignorable="xr"/.test(wb.pkg.text(`xl/comments${n}.xml`)));
    // (C3's thread box, there before, kept as it was but its z-index: Excel, saving, sized it to its placeholder's text.)
    const size = (b) => b.map((x) => (x[2] === '2' && x[3] === '2' && n === 1 ? x.slice(0, 4) : x));
    assert.deepEqual(size(boxesOf(wb.pkg.text(`xl/drawings/vmlDrawing${n}.vml`))), size(boxesOf(parts[`xl/drawings/vmlDrawing${n}.vml`])),
      `vmlDrawing${n}.vml: the thread's box kept, the notes' gone, new boxes after it at a new thread's place, ids on, z-index again`);
  }
  assert.equal(shapeAt(wb.pkg.text(VML), 2, 2), c3.replace(/z-index:\d+/, 'z-index:1'));
  // (The persons part: the old people kept, one new for each author's name; a note with no author none.)
  const people = (p) => [...p.matchAll(/<person displayName="([^"]*)" id="[^"]*" userId="([^"]*)" providerId="([^"]*)"\/>/g)].map((m) => m.slice(1).join('|')).sort();
  assert.deepEqual(people(persons), people(parts['xl/persons/person.xml']));
  assert.ok(persons.indexOf('userId="Avery Chen"') < persons.indexOf('userId="avery@example.com"'), 'the new people first, as Excel put them');
  assert.ok(!readNotes(wb, 0).size || [...readNotes(wb, 0).values()].every((n) => n.kind === 'thread'), 'no notes left');
  // (Excel reading the workbook converted here: the same threads, authors, texts and no dates as it read its own.)
  assert.deepEqual(readback['converted.xlsx'], excel.ours.after);
});

test("Convert to Comments on Excel's own workbook: a shown note's box comes hidden; nothing to convert, nothing changed", () => {
  const vml = excel.convert.parts[VML];
  assert.ok(shapes(vml).every((s) => /visibility:hidden/.test(s) && !/<x:Visible/.test(s)), "D4 shown before: Excel's box hidden");
  assert.ok(!/\sdT="/.test(excel.convert.parts['xl/threadedComments/threadedComment1.xml'].replace(/<threadedComment ref="F6"[^>]*>/, '')), 'no dates');
  const wb = open('widths.xlsx');
  const bytes = wb.toBytes();
  assert.equal(convertNotes(wb), 0);
  assert.ok(!wb.edited);
  assert.deepEqual(wb.toBytes(), bytes);
});

test("Convert to Comments leaves a sheet it's told to skip (a protected one: Excel's left its notes) as it was", () => {
  const wb = convertInput();
  const first = wb.pkg.text('xl/comments1.xml');
  assert.equal(convertNotes(wb, { skip: (k) => k === 0 }), 1);
  assert.equal(wb.pkg.text('xl/comments1.xml'), first);
  assert.equal(readNotes(wb, 1).get('2,1').kind, 'thread');
});
