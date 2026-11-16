import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import * as Y from 'yjs';
import { room, host, settle, view, cellOf } from './collab-host.js';
import { connectWorkbook } from '../src/collab.js';
import { LineSeq } from '../src/collab-seq.js';
import { groupLines } from '../src/ooxml/outline.js';

// Co-editing piece by piece (collab.js and its parts), with two Y.Docs relayed to each other in Node: formats, rows'
// and columns' sizes and hiding, rows and columns inserted and deleted, and sheets each go into the room as what they
// are, and concurrent changes end the same in both editors.
const BASIC = fs.readFileSync('test/fixtures/basic.xlsx');
const epochOf = (doc) => doc.getMap('oxe-files').get('epoch');
const bold = (wb, s) => wb.styles.fonts[wb.styles.cellXfs[s].fontId].b;
const styleAt = (api, si, r, c) => api.workbook.sheet(si).cells.get(r)?.get(c)?.s || 0;

test('the line sequence: inserts after one anchor end in one order whatever order they are applied in', () => {
  const ops = [
    { id: 'a.1', ts: 1, cid: 1, after: '1', n: 2 },
    { id: 'b.1', ts: 1, cid: 2, after: '1', n: 1 },
    { id: 'a.2', ts: 2, cid: 1, after: 'a.1:0', n: 1 },
  ];
  const order = (list) => { const s = new LineSeq(); for (const op of list) s.insert(op); return Array.from({ length: 7 }, (_, k) => s.tokAt(k)); };
  const one = order(ops);
  assert.deepEqual(order([ops[1], ops[0], ops[2]]), one);
  assert.deepEqual(one, ['0', '1', 'b.1:0', 'a.1:0', 'a.2:0', 'a.1:1', '2']);
  // (Deleted lines are passed over, and a line inserted after a deleted one still finds its place.)
  const s = new LineSeq();
  assert.deepEqual(s.remove(s.span(1, 2)), [{ pos: 1, n: 2 }]);
  assert.equal(s.tokAt(1), '3');
  assert.equal(s.insert({ id: 'c.1', ts: 1, cid: 3, after: '1', n: 1 }), 1);
  assert.deepEqual(s.posOf('2'), { pos: 2, alive: false });
});

test('the second editor opens the room; contents typed go across cell by cell, not as a new workbook', async () => {
  const { hosts: [a, b], docs } = await room(BASIC);
  assert.deepEqual(b.workbook.sheets.map((s) => s.name), a.workbook.sheets.map((s) => s.name));
  const epoch = epochOf(docs[0]);
  a.type(0, [{ r: 2, c: 7, value: 42 }]);
  b.type(0, [{ r: 3, c: 7, formula: 'H3*2' }]);
  assert.equal(cellOf(b, 0, 2, 7), 42);
  assert.equal(cellOf(a, 0, 3, 7), '=H3*2');
  assert.equal(a.workbook.sheet(0).cells.get(3).get(7).v, 84);
  assert.equal(epochOf(docs[1]), epoch, 'still the same epoch: nothing shared whole');
});

test('a format goes across as the format, matched or added in the other editor\'s styles part', async () => {
  const { hosts: [a, b], docs, net } = await room(BASIC);
  const epoch = epochOf(docs[0]);
  // (Each adds a format of its own at once, so each format has another index in the other.)
  net.hold();
  b.edit((wb) => wb.setStyles(0, [{ r: 9, c: 9, s: wb.styleWith(0, { fill: { pattern: 'solid', fg: { rgb: 'FFFF0000' } } }) }]));
  a.edit((wb) => wb.setStyles(0, [{ r: 2, c: 1, s: wb.styleWith(0, { font: { b: true, i: true } }) }]));
  await settle();
  net.release();
  assert.ok(bold(a.workbook, styleAt(a, 0, 2, 1)));
  assert.ok(bold(b.workbook, styleAt(b, 0, 2, 1)), 'bold in the second');
  assert.notEqual(styleAt(a, 0, 2, 1), styleAt(b, 0, 2, 1), 'at different indexes');
  assert.equal(a.workbook.styles.fills[a.workbook.styles.cellXfs[styleAt(a, 0, 9, 9)].fillId].fg?.rgb, 'FFFF0000', 'the second’s red fill in the first');
  assert.equal(cellOf(b, 0, 2, 1), 2, 'the contents kept');
  assert.equal(epochOf(docs[1]), epoch);
  assert.deepEqual(view(b), view(a));
});

test('row heights, column widths, hiding and line formats go across', async () => {
  const { hosts: [a, b], docs } = await room(BASIC);
  const epoch = epochOf(docs[0]);
  a.edit((wb) => wb.setLines(0, 'row', 4, 5, { height: 30 }));
  a.edit((wb) => wb.setLines(0, 'col', 6, 6, { width: 20 }));
  a.edit((wb) => wb.setLines(0, 'col', 8, 9, { hidden: true }));
  await settle();
  a.edit((wb) => wb.setLineStyles(0, { rows: [{ r: 12, s: wb.styleWith(0, { font: { i: true } }) }] }));
  await settle();
  assert.equal(b.workbook.sheet(0).rows.get(4).ht, 30);
  assert.equal(b.workbook.sheet(0).rows.get(5).ht, 30);
  assert.equal(b.workbook.sheet(0).cols.find((x) => x.min <= 6 && x.max >= 6).width, a.workbook.sheet(0).cols.find((x) => x.min <= 6 && x.max >= 6).width);
  assert.ok(b.workbook.sheet(0).cols.find((x) => x.min <= 8 && x.max >= 8).hidden);
  assert.ok(b.workbook.styles.fonts[b.workbook.styles.cellXfs[b.workbook.sheet(0).rows.get(12).s].fontId].i, 'row 13 italic in the second');
  assert.equal(epochOf(docs[1]), epoch);
  assert.deepEqual(view(b), view(a));
});

test('rows inserted and deleted go across as operations; typing after them stays in step', async () => {
  const { hosts: [a, b], docs } = await room(BASIC);
  const epoch = epochOf(docs[0]);
  a.edit((wb) => wb.changeLines(0, 'row', 2, 2));
  assert.equal(cellOf(b, 0, 4, 0), 'Item 2');
  assert.equal(cellOf(b, 0, 4, 2), '=B5*2', 'formulas moved as the inserting editor moved them');
  b.type(0, [{ r: 2, c: 0, value: 'new' }]);
  assert.equal(cellOf(a, 0, 2, 0), 'new');
  b.edit((wb) => wb.changeLines(0, 'col', 1, -1));
  assert.equal(cellOf(a, 0, 4, 1), '=#REF!*2', 'a column deleted in the second: the first\'s formula reads #REF!');
  assert.equal(epochOf(docs[1]), epoch);
  assert.deepEqual(view(b), view(a));
});

test('two editors inserting rows at the same place at once both get theirs, in one order', async () => {
  const { hosts: [a, b], net } = await room(BASIC);
  net.hold();
  a.edit((wb) => wb.changeLines(0, 'row', 2, 1));
  a.type(0, [{ r: 2, c: 0, value: 'from one' }]);
  b.edit((wb) => wb.changeLines(0, 'row', 2, 2));
  b.type(0, [{ r: 2, c: 0, value: 'from two' }, { r: 3, c: 0, value: 'two again' }]);
  await settle();
  net.release();
  const col = (api) => Array.from({ length: 7 }, (_, r) => cellOf(api, 0, r, 0));
  assert.deepEqual(col(a), col(b));
  assert.deepEqual(col(a).filter((x) => /one|two/.test(x)).sort(), ['from one', 'from two', 'two again']);
  assert.equal(cellOf(a, 0, 5, 0), 'Item 2', 'Item 2 below the three new rows');
  assert.deepEqual(view(b), view(a));
});

test('one editor formatting and typing while the other inserts and deletes rows above: the same cells in both', async () => {
  const { hosts: [a, b], net } = await room(BASIC);
  net.hold();
  // (Item 4 is on row 5 when the first makes it bold and types beside it.)
  a.edit((wb) => wb.setStyles(0, [{ r: 4, c: 0, s: wb.styleWith(0, { font: { b: true } }) }]));
  await settle();
  a.type(0, [{ r: 4, c: 7, value: 'beside 4' }]);
  a.edit((wb) => wb.setLines(0, 'row', 4, 4, { height: 24 }));
  b.edit((wb) => wb.changeLines(0, 'row', 1, 2));
  b.edit((wb) => wb.changeLines(0, 'row', 0, -1));
  await settle();
  net.release();
  for (const api of [a, b]) {
    assert.equal(cellOf(api, 0, 5, 0), 'Item 4');
    assert.ok(bold(api.workbook, styleAt(api, 0, 5, 0)), 'Item 4 bold where it is now');
    assert.equal(cellOf(api, 0, 5, 7), 'beside 4');
    assert.equal(api.workbook.sheet(0).rows.get(5).ht, 24);
  }
  assert.deepEqual(view(b), view(a));
});

test('a formula typed while the other editor inserts rows reads the cell it was typed to read', async () => {
  const { hosts: [a, b], net } = await room(BASIC);
  net.hold();
  a.type(0, [{ r: 9, c: 7, formula: 'B4+1' }]);
  b.edit((wb) => wb.changeLines(0, 'row', 0, 1));
  await settle();
  net.release();
  assert.equal(cellOf(a, 0, 10, 7), '=B5+1');
  assert.equal(cellOf(b, 0, 10, 7), '=B5+1');
  assert.equal(b.workbook.sheet(0).cells.get(10).get(7).v, 4);
});

test('sheets added, renamed, moved, hidden and coloured go across; two added at once both arrive', async () => {
  const { hosts: [a, b], docs, net } = await room(BASIC);
  const epoch = epochOf(docs[0]);
  a.edit((wb) => wb.renameSheet(0, 'Numbers'));
  a.edit((wb) => wb.setTabColor(1, { rgb: 'FF00B050' }));
  a.edit((wb) => wb.setSheetHidden(2, false));
  a.edit((wb) => wb.moveSheet(0, 2));
  await settle();
  assert.deepEqual(b.workbook.sheets.map((s) => [s.name, s.state]), a.workbook.sheets.map((s) => [s.name, s.state]));
  net.hold();
  a.edit((wb) => wb.addSheet(0, 'From one'));
  b.edit((wb) => wb.addSheet(0, 'From two'));
  await settle();
  net.release();
  assert.deepEqual(b.workbook.sheets.map((s) => s.name), a.workbook.sheets.map((s) => s.name));
  assert.ok(a.workbook.sheets.some((s) => s.name === 'From two') && b.workbook.sheets.some((s) => s.name === 'From one'));
  const k = a.workbook.sheets.findIndex((s) => s.name === 'From two');
  b.type(b.workbook.sheets.findIndex((s) => s.name === 'From two'), [{ r: 0, c: 0, value: 'on the new sheet' }]);
  assert.equal(cellOf(a, k, 0, 0), 'on the new sheet');
  a.edit((wb) => wb.deleteSheet(wb.sheets.findIndex((s) => s.name === 'From one')));
  await settle();
  assert.deepEqual(b.workbook.sheets.map((s) => s.name), a.workbook.sheets.map((s) => s.name));
  assert.equal(epochOf(docs[1]), epoch);
  assert.deepEqual(view(b), view(a));
});

test('an editor joining later opens the room with every change since', async () => {
  const { hosts: [a], docs } = await room(BASIC, 1);
  a.edit((wb) => wb.changeLines(0, 'row', 3, 2));
  a.type(0, [{ r: 3, c: 0, value: 'inserted' }, { r: 20, c: 2, formula: 'SUM(B2:B8)' }]);
  a.edit((wb) => { wb.setStyles(0, [{ r: 3, c: 0, s: wb.styleWith(0, { font: { b: true } }) }]); wb.setLines(0, 'col', 3, 3, { width: 30 }); });
  a.edit((wb) => wb.addSheet(1, 'Later'));
  await settle();
  const c = host();
  const doc = new Y.Doc();
  Y.applyUpdate(doc, Y.encodeStateAsUpdate(docs[0]));
  await connectWorkbook(c, doc);
  assert.deepEqual(view(c), view(a));
});

test('a change the room doesn\'t hold piece by piece (a merge) shares the whole workbook as a new epoch', async () => {
  const { hosts: [a, b], docs } = await room(BASIC);
  const epoch = epochOf(docs[0]);
  a.edit((wb) => wb.setMerges(0, [...wb.sheet(0).merges, { r0: 20, c0: 0, r1: 20, c1: 2 }]));
  await settle();
  assert.notEqual(epochOf(docs[1]), epoch);
  assert.ok(b.workbook.sheet(0).merges.some((m) => m.r0 === 20 && m.c1 === 2));
  // (And piece by piece again after it.)
  a.type(0, [{ r: 0, c: 9, value: 'after' }]);
  assert.equal(cellOf(b, 0, 0, 9), 'after');
});

test('rows grouped go piece by piece; a cell style the workbook hadn’t used goes as the whole workbook', async () => {
  const { hosts: [a, b], docs } = await room(BASIC);
  const epoch = epochOf(docs[0]);
  a.edit((wb) => groupLines(wb, 0, 'row', 2, 4));
  await settle();
  assert.equal(b.workbook.sheet(0).rows.get(3)?.level, 1);
  assert.equal(epochOf(docs[1]), epoch);
  a.edit((wb) => wb.setStyles(0, [{ r: 1, c: 1, s: wb.styleWith(0, { cellStyle: 'Good' }) }]));
  await settle();
  assert.notEqual(epochOf(docs[1]), epoch, 'its cell style added to the styles part: a new epoch');
  assert.deepEqual(view(b), view(a));
});

test('room entries that aren\'t what this version writes are left out, not written into the workbook', async () => {
  const { hosts: [a, b], docs } = await room(BASIC);
  const before = JSON.stringify(view(b));
  const roomOf = (doc) => doc.getMap('oxe-rooms').get(epochOf(doc));
  const data = roomOf(docs[0]).get('cells').get('s1');
  docs[0].transact(() => {
    data.set('r 3', { r: '40', ht: '30', customHeight: '1' }); // (a row's own place)
    data.set('c 2', { 'bad name': '1' });
    data.set('v 1,x', { v: 'nowhere' });
    data.set('v 2,1', { v: { nested: true } });
    roomOf(docs[0]).get('styles').set('k1', { xf: ['xf', {}, [['bad name', {}, []]]], font: null, fill: null, border: null, numFmt: { id: 0 } });
    data.set('s 2,2', 'k1');
    roomOf(docs[0]).get('lines').push([{ id: 'x', cid: 1, seq: 1, ts: 1, sheet: 's1', axis: 'row', kind: 'ins', after: null, n: 1e9 }]);
    roomOf(docs[0]).get('props').set('s2/color', { rgb: '"><x/>' });
  }, 'relay-test');
  await settle();
  const now = view(b);
  assert.equal(b.workbook.sheet(0).rows.get(3)?.ht ?? null, a.workbook.sheet(0).rows.get(3)?.ht ?? null, 'the row entry naming its own place left out');
  assert.equal(cellOf(b, 0, 2, 1), 2, 'contents that aren\'t a value left out');
  assert.equal(styleAt(b, 0, 2, 2), 0, 'a format described with a bad element name: the Normal format');
  assert.deepEqual(now[1].color, JSON.parse(before)[1].color, 'a colour that isn\'t hex left out');
  assert.equal(cellOf(b, 0, 5, 0), 'Item 5', 'an insert of a billion rows left out');
});
