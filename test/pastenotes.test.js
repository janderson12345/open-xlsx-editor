import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { snapshot, pasteNotes, pasteValidation } from '../src/ooxml/clipops.js';
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

test("a paste's validation as Excel's: the copy's rules over the pasted cells, joining the rule itself on its own sheet (Excel's file), moved as formulas are elsewhere", async () => {
  const { serializeXml } = await import('../src/ooxml/xml.js');
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'Rules');
  for (const [ref, v] of [['B2', 1], ['C2', 2], ['B3', 3], ['C3', 4]]) wb.setCell(si, ref, v);
  wb.setValidation(si, [{ r0: 1, c0: 2, r1: 2, c1: 2 }], { type: 'custom', f1: '=C2>B2' });
  wb.setValidation(si, [{ r0: 1, c0: 1, r1: 2, c1: 1 }], { type: 'list', f1: '=$A$1:$A$3' });
  wb.setValidation(si, [{ r0: 6, c0: 8, r1: 6, c1: 8 }], { type: 'whole', op: 'between', f1: '0', f2: '9' });
  const copy = snapshot(wb, si, { ar: 1, ac: 1, r0: 1, c0: 1, r1: 2, c1: 2 });
  pasteValidation(wb, si, { r0: 4, c0: 7, r1: 5, c1: 8 }, copy); // H5
  pasteValidation(wb, si, { r0: 4, c0: 11, r1: 5, c1: 12 }, copy); // L5
  pasteValidation(wb, si, { r0: 4, c0: 14, r1: 7, c1: 17 }, copy); // O5:R8
  const sqrefs = [...serializeXml(wb._part(si).doc).matchAll(/<dataValidation [^>]*sqref="([^"]+)"[^>]*>(?:<formula1>([^<]*))?/g)].map((m) => `${m[1]} | ${m[2]}`);
  // (Excel's own, after Paste on H5 and PasteSpecial(xlPasteValidation) on L5 and O5:R8; I7's rule, not pasted over, kept.)
  assert.deepEqual(sqrefs, ['C2:C3 I5:I6 M5:M6 P5:P8 R5:R8 | C2&gt;B2', 'B2:B3 H5:H6 L5:L6 O5:O8 Q5:Q8 | $A$1:$A$3', 'I7 | 0']);
  // (On another sheet: a rule of its own, its formula moved as a copied formula is.)
  const other = wb.addSheet(wb.sheets.length, 'Other');
  pasteValidation(wb, other, { r0: 4, c0: 7, r1: 5, c1: 8 }, copy);
  const there = [...serializeXml(wb._part(other).doc).matchAll(/sqref="([^"]+)"[^>]*><formula1>([^<]*)/g)].map((m) => `${m[1]} | ${m[2]}`);
  assert.deepEqual(there, ['I5:I6 | I5&gt;H5', 'H5:H6 | $A$1:$A$3']);
});

test('validation pasted with Transpose, as Excel\'s: the turned cells join the rule on its own sheet, elsewhere a rule moved by the corner', async () => {
  const { serializeXml } = await import('../src/ooxml/xml.js');
  // (Excel: A1:A3 or A1:B2 copied, PasteSpecial(xlPasteAll, Transpose) at C1 or D1; the dataValidations it saved.)
  const rules = (wb, k) => [...serializeXml(wb._part(k).doc).matchAll(/<dataValidation [^>]*sqref="([^"]+)"[^>]*>(?:<formula1>([^<]*))?/g)].map((m) => `${m[1]} | ${m[2]}`);
  const sheet = (set) => {
    const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
    const si = wb.addSheet(wb.sheets.length, 'Turned');
    for (const [range, rule] of set) wb.setValidation(si, [range], rule);
    return { wb, si };
  };
  const A13 = { r0: 0, c0: 0, r1: 2, c1: 0 };
  const C1E1 = { r0: 0, c0: 2, r1: 0, c1: 4 };
  const whole = { type: 'whole', op: 'between', f1: '1', f2: '10' };
  for (const [rule, want] of [[whole, 'A1:A3 C1:E1 | 1'], [{ type: 'custom', f1: '=A1>B1' }, 'A1:A3 C1:E1 | A1&gt;B1'], [{ type: 'list', f1: '=$H$1:$H$3' }, 'A1:A3 C1:E1 | $H$1:$H$3']]) {
    const { wb, si } = sheet([[A13, rule]]);
    pasteValidation(wb, si, C1E1, snapshot(wb, si, A13), { transpose: true });
    assert.deepEqual(rules(wb, si), [want]);
  }
  // (A block: A's rule to row 1, B's to row 2; over C1:F1's list, F1 keeps it.)
  const block = sheet([[{ r0: 0, c0: 0, r1: 1, c1: 0 }, whole], [{ r0: 0, c0: 1, r1: 1, c1: 1 }, { type: 'list', f1: 'a,b' }]]);
  pasteValidation(block.wb, block.si, { r0: 0, c0: 3, r1: 1, c1: 4 }, snapshot(block.wb, block.si, { r0: 0, c0: 0, r1: 1, c1: 1 }), { transpose: true });
  assert.deepEqual(rules(block.wb, block.si), ['A1:A2 D1:E1 | 1', 'B1:B2 D2:E2 | "a,b"']);
  const over = sheet([[A13, whole], [{ r0: 0, c0: 2, r1: 0, c1: 5 }, { type: 'list', f1: 'p,q' }]]);
  pasteValidation(over.wb, over.si, C1E1, snapshot(over.wb, over.si, A13), { transpose: true });
  assert.deepEqual(rules(over.wb, over.si), ['A1:A3 C1:E1 | 1', 'F1 | "p,q"']);
  // (Another sheet: a rule of its own, =C1>D1 over C1:E1.)
  const custom = sheet([[A13, { type: 'custom', f1: '=A1>B1' }]]);
  const other = custom.wb.addSheet(custom.wb.sheets.length, 'Other');
  pasteValidation(custom.wb, other, C1E1, snapshot(custom.wb, custom.si, A13), { transpose: true });
  assert.deepEqual(rules(custom.wb, other), ['C1:E1 | C1&gt;D1']);
});

test('validation copied or cut to another sheet, as Excel\'s: the same formulas there, the cut cells leaving their rules, the pasted-over cells theirs', async () => {
  const { serializeXml } = await import('../src/ooxml/xml.js');
  const { moveCutParts } = await import('../src/ooxml/clipops.js');
  const { parseRange } = await import('../src/ooxml/refs.js');
  // (Excel: B2:B3 a list =$A$1:$A$3 and C2:C3 =C2>B2 (or =$E$1>C2), B2:C3 copied or cut to another sheet's H5; there,
  // in one case, whole numbers over H5:K8. Its rules compared by the cells they cover and their first formula.)
  const cellsOf = (sqref) => sqref.split(' ').flatMap((t) => { const a = parseRange(t); const out = []; for (let r = a.r0; r <= a.r1; r++) for (let c = a.c0; c <= a.c1; c++) out.push(`${r},${c}`); return out; }).sort().join(' ');
  const rules = (wb, k) => [...serializeXml(wb._part(k).doc).matchAll(/<dataValidation [^>]*sqref="([^"]+)"[^>]*>(?:<formula1>([^<]*))?/g)].map((m) => `${cellsOf(m[1])} | ${m[2]}`).sort();
  const want = (list) => list.map(([sq, f]) => `${cellsOf(sq)} | ${f}`).sort();
  for (const [name, abs, cut, over, src, dest] of [
    ['copy', false, false, false, [['B2:B3', '$A$1:$A$3'], ['C2:C3', 'C2&gt;B2']], [['I5:I6', 'I5&gt;H5'], ['H5:H6', '$A$1:$A$3']]],
    ['cut', false, true, false, [], [['I5:I6', 'I5&gt;H5'], ['H5:H6', '$A$1:$A$3']]],
    ['copy with $E$1', true, false, false, [['B2:B3', '$A$1:$A$3'], ['C2:C3', '$E$1&gt;C2']], [['I5:I6', '$E$1&gt;I5'], ['H5:H6', '$A$1:$A$3']]],
    ['cut with $E$1', true, true, false, [], [['I5:I6', '$E$1&gt;I5'], ['H5:H6', '$A$1:$A$3']]],
    ['copy over a rule', false, false, true, [['B2:B3', '$A$1:$A$3'], ['C2:C3', 'C2&gt;B2']], [['J5:K8 H7:I8', '1'], ['I5:I6', 'I5&gt;H5'], ['H5:H6', '$A$1:$A$3']]],
  ]) {
    const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
    const si = wb.addSheet(wb.sheets.length, 'Src');
    const d = wb.addSheet(wb.sheets.length, 'Dest');
    wb.setValidation(si, [{ r0: 1, c0: 1, r1: 2, c1: 1 }], { type: 'list', f1: '=$A$1:$A$3' });
    wb.setValidation(si, [{ r0: 1, c0: 2, r1: 2, c1: 2 }], { type: 'custom', f1: abs ? '=$E$1>C2' : '=C2>B2' });
    if (over) wb.setValidation(d, [{ r0: 4, c0: 7, r1: 7, c1: 10 }], { type: 'whole', op: 'between', f1: '1', f2: '9' });
    const copy = snapshot(wb, si, { r0: 1, c0: 1, r1: 2, c1: 2 }, cut);
    const area = { r0: 4, c0: 7, r1: 5, c1: 8 };
    if (cut) moveCutParts(wb, d, area, copy); else pasteValidation(wb, d, area, copy);
    assert.deepEqual(rules(wb, si), want(src), `${name}: the sheet copied from`);
    assert.deepEqual(rules(wb, d), want(dest), `${name}: the sheet pasted on`);
  }
});

test("a cut's notes and validation move with its cells, as Excel's (B2:C3 cut to E5: the note to E5, F6's gone; C2:C4's rule C4 F5:F6, =C4>B4; E5's rule gone)", async () => {
  const { serializeXml } = await import('../src/ooxml/xml.js');
  const { moveCutParts } = await import('../src/ooxml/clipops.js');
  const { wb, si } = book();
  wb.setValidation(si, [{ r0: 1, c0: 2, r1: 3, c1: 2 }], { type: 'custom', f1: '=C2>B2' });
  wb.setValidation(si, [{ r0: 4, c0: 4, r1: 4, c1: 4 }], { type: 'whole', op: 'between', f1: '0', f2: '9' });
  // (E5's own note and value aside: Excel's case had F6's note and E5's rule.)
  const cut = snapshot(wb, si, { ar: 1, ac: 1, r0: 1, c0: 1, r1: 2, c1: 2 }, true);
  assert.equal(moveCutParts(wb, si, { r0: 4, c0: 4, r1: 5, c1: 5 }, cut), true);
  const notes = notesOf(wb, si);
  assert.deepEqual(Object.keys(notes).sort(), ['4,4']);
  assert.match(notes['4,4'], /note on b2$/);
  const rules = [...serializeXml(wb._part(si).doc).matchAll(/sqref="([^"]+)"[^>]*><formula1>([^<]*)/g)].map((m) => `${m[1]} | ${m[2]}`);
  assert.deepEqual(rules, ['C4 F5:F6 | C4&gt;B4']);
});

test('cells taken out of a rule rebase its formulas when its first cell changes, as Excel does', () => {
  const { wb, si } = book();
  wb.setValidation(si, [{ r0: 1, c0: 2, r1: 3, c1: 2 }], { type: 'custom', f1: '=C2>B2' });
  wb.setValidation(si, [{ r0: 1, c0: 2, r1: 2, c1: 2 }], null);
  const v = wb.sheet(si).validations[0];
  assert.deepEqual([v.ranges.map((a) => `${a.r0},${a.c0}`), v.formulas], [['3,2'], ['C4>B4']]);
});

test("comment threads come with a paste and move with a cut, as Excel's: replies, authors, times and resolved kept, new ids", async () => {
  const { addThread, replyThread, resolveThread } = await import('../src/ooxml/threadedit.js');
  const { moveCutParts } = await import('../src/ooxml/clipops.js');
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'Threads');
  const when = new Date(2026, 10, 12, 22, 5, 2, 180);
  addThread(wb, si, 1, 1, 'first', 'Avery Chen', { when });
  replyThread(wb, si, 1, 1, 'a reply', 'Blair Diaz', { when: new Date(2026, 10, 12, 22, 5, 2, 200) });
  addThread(wb, si, 1, 2, 'second', 'Avery Chen', { when });
  resolveThread(wb, si, 1, 2, true);
  const copy = snapshot(wb, si, { ar: 1, ac: 1, r0: 1, c0: 1, r1: 1, c1: 2 });
  pasteNotes(wb, si, { r0: 4, c0: 4, r1: 4, c1: 5 }, copy);
  const n = readNotes(wb, si);
  assert.deepEqual(n.get('4,4').items.map((x) => [x.author, x.text, x.date]), [['Avery Chen', 'first', '2026-11-12T22:05:02.18'], ['Blair Diaz', 'a reply', '2026-11-12T22:05:02.20']]);
  assert.equal(n.get('4,5').resolved, true);
  const part = () => wb.pkg.text(Object.keys(wb.pkg.files).find((k) => k.startsWith('xl/threadedComments/')));
  assert.equal(new Set([...part().matchAll(/\sid="(\{[^}]+\})"/g)].map((m) => m[1])).size, 6);
  // (Cut to E11: gone from B2:C2, there at E11:F11.)
  const cut = snapshot(wb, si, { ar: 1, ac: 1, r0: 1, c0: 1, r1: 1, c1: 2 }, true);
  moveCutParts(wb, si, { r0: 10, c0: 4, r1: 10, c1: 5 }, cut);
  const after = readNotes(wb, si);
  assert.deepEqual([after.has('1,1'), after.has('1,2'), after.get('10,4')?.items.length, after.get('10,5')?.resolved], [false, false, 2, true]);
});
