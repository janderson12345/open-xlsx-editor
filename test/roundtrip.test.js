import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync, strFromU8 } from 'fflate';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseXml } from '../src/ooxml/xml.js';
import { cellAt } from '../src/ooxml/sheet.js';
import { cellText } from '../src/ooxml/numfmt.js';

const fixtures = fs.readdirSync('test/fixtures').filter((f) => f.endsWith('.xlsx'));
const bytes = (f) => new Uint8Array(fs.readFileSync(`test/fixtures/${f}`));
const same = (a, b) => Buffer.from(a).equals(Buffer.from(b));
/** Every worksheet's model, as plain data to compare. */
const models = (wb) => wb.sheets.map((s, i) => (s.kind === 'worksheet' ? JSON.stringify(wb.sheet(i), (k, v) => (v instanceof Map ? [...v] : v)) : null));
const text = (wb, i, r, c) => cellText(cellAt(wb.sheet(i), r, c));

test('there are fixtures (npm run fixtures makes them)', () => {
  assert.ok(fixtures.length >= 3, fixtures.join());
});

for (const f of fixtures) {
  test(`${f}: open and save without changes gives every part back byte for byte, and the same model`, () => {
    const src = bytes(f);
    const wb = Workbook.fromBytes(src);
    const before = models(wb);
    const out = wb.toBytes();
    const a = unzipSync(src);
    const b = unzipSync(out);
    assert.deepEqual(Object.keys(b).sort(), Object.keys(a).sort());
    for (const k of Object.keys(a)) assert.ok(same(a[k], b[k]), `${k} changed`);
    assert.deepEqual(models(Workbook.fromBytes(out)), before);
  });

  test(`${f}: an edit rewrites only that sheet, which stays well-formed and keeps what isn't modelled`, () => {
    const src = bytes(f);
    const wb = Workbook.fromBytes(src);
    const i = wb.activeSheet;
    const target = wb.sheets[i].path;
    wb.setCell(i, 'Z100', 'edited');
    wb.setCell(i, 'A2', 42);
    const out = wb.toBytes();
    const a = unzipSync(src);
    const b = unzipSync(out);
    for (const k of Object.keys(a)) if (k !== target) assert.ok(same(a[k], b[k]), `${k} changed`);
    const xml = strFromU8(b[target]);
    parseXml(xml);
    assert.match(xml, /^<\?xml version="1.0" encoding="UTF-8" standalone="yes"\?>/);
    const before = strFromU8(a[target]);
    for (const kept of ['<conditionalFormatting', '<dataValidations', '<extLst', '<mergeCells', '<pageMargins']) assert.equal(xml.includes(kept), before.includes(kept), kept);
    const again = Workbook.fromBytes(out);
    assert.equal(text(again, i, 99, 25), 'edited');
    assert.equal(text(again, i, 1, 0), '42');
    // The other cells read as before.
    const m0 = wb.sheet(i);
    const m1 = Workbook.fromBytes(src).sheet(i);
    for (const [r, line] of m1.cells) for (const [c, cell] of line) if (!(r === 1 && c === 0)) assert.deepEqual(cellAt(m0, r, c), cell, `${r},${c}`);
  });
}

test('basic.xlsx: sheets, strings, values, sizes and merges as written', () => {
  const wb = Workbook.fromBytes(bytes('basic.xlsx'));
  assert.deepEqual(wb.sheets.map((s) => [s.name, s.state, s.path]), [
    ['Data', 'visible', 'xl/worksheets/sheet1.xml'],
    ['Notes & more', 'visible', 'xl/worksheets/sheet2.xml'],
    ['Hidden', 'hidden', 'xl/worksheets/sheet3.xml'],
  ]);
  assert.equal(wb.activeSheet, 0);
  const sh = wb.sheet(0);
  assert.equal(text(wb, 0, 0, 0), 'Item');
  assert.equal(cellAt(sh, 0, 0).s, 1);
  assert.equal(text(wb, 0, 1, 1), '1');
  assert.equal(cellAt(sh, 2, 2).f, ''); // (a shared formula's later cells carry only `si`)
  assert.equal(cellAt(sh, 1, 2).f, 'B2*2');
  assert.equal(text(wb, 0, 10, 2), 'TRUE');
  assert.equal(text(wb, 0, 10, 3), '#DIV/0!');
  assert.equal(text(wb, 0, 10, 4), 'inline text');
  assert.equal(text(wb, 0, 10, 5), '1234567.891');
  assert.equal(text(wb, 0, 11, 4), 'Rich bold'); // (the phonetic run's "ruby" is not cell text)
  assert.equal(text(wb, 0, 11, 5), 'Line\rend');
  assert.equal(text(wb, 0, 12, 5), '1.2345E-05');
  assert.equal(text(wb, 0, 39, 7), '1.23457E+11');
  assert.deepEqual(sh.rows.get(3), { ht: 30, hidden: false, customHeight: true });
  assert.equal(sh.rows.get(5).hidden, true);
  assert.deepEqual(sh.cols.map((c) => [c.min, c.max, c.width, c.hidden]), [[1, 1, 20.7109375, false], [3, 3, 9.140625, true], [5, 5, 30.7109375, false]]);
  assert.deepEqual(sh.merges, [{ r0: 11, c0: 0, r1: 12, c1: 2 }]);
  assert.equal(sh.maxRow, 39);
  assert.equal(sh.maxCol, 7);
  assert.equal(text(wb, 1, 1, 0), '  leading spaces kept');
  assert.equal(text(wb, 1, 1, 1), '-3.5');
});

test('prefixed.xlsx: implied row and cell positions follow on from the one before', () => {
  const wb = Workbook.fromBytes(bytes('prefixed.xlsx'));
  assert.deepEqual([[0, 0], [0, 1], [1, 0], [1, 1], [4, 0], [4, 2], [4, 3], [5, 0]].map(([r, c]) => text(wb, 0, r, c)), ['First', 'Second', '1', '2', '5', '3', '4', 'Row six']);
  // An edit pins the implied positions, so the cells after the new one stay where they were.
  wb.setCell(0, 'B5', 'new');
  const again = Workbook.fromBytes(wb.toBytes());
  assert.deepEqual([[4, 0], [4, 1], [4, 2], [4, 3], [5, 0]].map(([r, c]) => text(again, 0, r, c)), ['5', 'new', '3', '4', 'Row six']);
});

test('a cell outside a row\'s `spans` hint drops the hint; one inside keeps it', () => {
  const wb = Workbook.fromBytes(bytes('basic.xlsx'));
  wb.setCell(0, 'F3', 'inside');
  wb.setCell(0, 'Z2', 'outside');
  const xml = strFromU8(unzipSync(wb.toBytes())['xl/worksheets/sheet1.xml']);
  assert.match(xml, /<row r="2">/);
  assert.match(xml, /<row r="3" spans="1:6">/);
});

test('formula cells are refused for now, and nothing is changed', () => {
  const wb = Workbook.fromBytes(bytes('basic.xlsx'));
  assert.throws(() => wb.setCell(0, 'C2', 1), /formula/);
});
