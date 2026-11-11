import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { spawnSync } from 'node:child_process';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { Workbook } from '../src/ooxml/workbook.js';
import { MESSAGES } from '../src/ooxml/sheetops.js';
import { parseRef } from '../src/ooxml/refs.js';

// What Excel did with the same steps over COM (docs/SHEETS.md) is what these expect.
const dir = 'test/out/sheets';
fs.mkdirSync(dir, { recursive: true });
const book = (name, change = null) => {
  const files = unzipSync(new Uint8Array(fs.readFileSync(`test/fixtures/${name}`)));
  if (change) files['xl/workbook.xml'] = strToU8(change(strFromU8(files['xl/workbook.xml'])));
  return Workbook.fromBytes(zipSync(files));
};
const cell = (wb, sheet, ref) => {
  const { r, c } = parseRef(ref);
  return wb.sheet(wb.sheetIndexOf(sheet)).cells.get(r)?.get(c) || null;
};
const names = (wb) => wb.sheets.map((s) => s.name);
const saved = (wb, file) => {
  const bytes = wb.toBytes();
  fs.writeFileSync(`${dir}/${file}`, bytes);
  return { wb: Workbook.fromBytes(bytes), files: unzipSync(bytes) };
};
const refuses = (fn, message) => assert.throws(fn, (e) => e.message === message);

test('new sheets are named as Excel names them: Sheet1, Sheet2..., counting on past names taken or renamed', () => {
  const wb = book('basic.xlsx');
  assert.equal(wb.addSheet(1), 1);
  assert.deepEqual(names(wb), ['Data', 'Sheet1', 'Notes & more', 'Hidden']);
  wb.renameSheet(1, 'X');
  wb.addSheet(1);
  assert.equal(wb.sheets[1].name, 'Sheet2');
  const other = book('basic.xlsx');
  other.renameSheet(0, 'Sheet3');
  for (let k = 0; k < 3; k++) other.addSheet(other.sheets.length);
  assert.deepEqual(names(other).slice(-3), ['Sheet1', 'Sheet2', 'Sheet4']);
  const { wb: back, files } = saved(other, 'added.xlsx');
  assert.deepEqual(names(back), ['Sheet3', 'Notes & more', 'Hidden', 'Sheet1', 'Sheet2', 'Sheet4']);
  assert.deepEqual(back.sheets.map((s) => s.sheetId), [1, 2, 3, 4, 5, 6]);
  assert.match(strFromU8(files['xl/worksheets/sheet4.xml']), /<sheetFormatPr defaultRowHeight="15"\/><sheetData\/><pageMargins/);
  assert.match(strFromU8(files['[Content_Types].xml']), /PartName="\/xl\/worksheets\/sheet6.xml"/);
});

test('sheet names Excel refuses, in its words', () => {
  const wb = book('basic.xlsx');
  for (const bad of ['a/b', 'x'.repeat(32), "'x", "x'", '', 'a[b', 'a:b', 'a*b', 'a?b', 'a\\b']) refuses(() => wb.renameSheet(1, bad), MESSAGES.invalid);
  refuses(() => wb.renameSheet(1, 'data'), MESSAGES.taken);
  refuses(() => wb.renameSheet(1, 'History'), MESSAGES.reserved);
  for (const ok of ['x'.repeat(31), "a'b", ' ', 'Sheet 9', 'NOTES & MORE']) wb.renameSheet(1, ok);
  assert.equal(wb.sheets[1].name, 'NOTES & MORE');
});

test('a rename rewrites every formula and name that reads the sheet; values stay, and a change there still reaches them', () => {
  const wb = book('formulas.xlsx');
  wb.calculate();
  const before = cell(wb, 'Calc', 'B5').v;
  wb.renameSheet(wb.sheetIndexOf('Data'), 'Sales 2024');
  // (Excel: =Data!B2:B9 became ='Sales 2024'!B2:B9; =SUM(Data!B2:B9*1) became =SUM('Sales 2024'!B2:B9*1).)
  assert.equal(cell(wb, 'Calc', 'B2').f, "'Sales 2024'!B2:B9");
  assert.equal(cell(wb, 'Calc', 'B5').f, "SUM('Sales 2024'!B2:B9*1)");
  assert.equal(cell(wb, 'Calc', 'B5').v, before);
  assert.deepEqual(wb.definedNames.map((d) => d.text), ["'Sales 2024'!$A$2:$A$9", "'Sales 2024'!$C$2"]);
  const total = [...wb.sheet(0).cells].find(([, line]) => line.get(1)?.f === "SUM('Sales 2024'!B2:B9)");
  const was = total[1].get(1).v;
  wb.setCell(wb.sheetIndexOf('Sales 2024'), 'B3', 105);
  assert.equal(total[1].get(1).v, was + 100);
  const { files } = saved(wb, 'renamed.xlsx');
  assert.match(strFromU8(files['xl/workbook.xml']), /<sheet name="Sales 2024" sheetId="2"/);
  assert.match(strFromU8(files['xl/workbook.xml']), /<definedName name="Items">'Sales 2024'!\$A\$2:\$A\$9<\/definedName>/);
});

test('a delete turns references to the sheet into #REF!, takes its own names away and moves the others with their sheets', () => {
  // The names Excel was given in the measurement: one on Data, one on My Data, one of the workbook's reading My Data.
  const wb = book('formulas.xlsx', (xml) => xml.replace('<definedName name="Items">',
    '<definedName name="Here" localSheetId="1">Data!$B$3</definedName><definedName name="LocalOnMy">\'My Data\'!$A$1</definedName>' +
    '<definedName name="Mine" localSheetId="2">\'My Data\'!$B$2</definedName><definedName name="Items">'));
  wb.calculate();
  const calc = wb.sheetIndexOf('Calc');
  wb.setFormula(calc, 'H1', "SUM('My Data'!A1:B2,Data!B2)");
  wb.setFormula(calc, 'H2', 'Mine');
  wb.deleteSheet(wb.sheetIndexOf('My Data'));
  assert.deepEqual(names(wb), ['Calc', 'Data']);
  assert.deepEqual(['B95', 'B96', 'H1', 'H2'].map((r) => [cell(wb, 'Calc', r).f, cell(wb, 'Calc', r).v]),
    [['#REF!*2', '#REF!'], ['#REF!', '#REF!'], ['SUM(#REF!,Data!B2)', '#REF!'], ['Mine', '#NAME?']]);
  wb.moveSheet(1, 0);
  assert.deepEqual(names(wb), ['Data', 'Calc']);
  const { files } = saved(wb, 'deleted.xlsx');
  const xml = strFromU8(files['xl/workbook.xml']);
  // (Excel's file: <definedName name="Here" localSheetId="0">Data!$B$3</definedName> ... name="LocalOnMy">#REF!<...)
  assert.match(xml, /<definedName name="Here" localSheetId="0">Data!\$B\$3<\/definedName><definedName name="LocalOnMy">#REF!<\/definedName><definedName name="Items">/);
  assert.doesNotMatch(xml, /Mine/);
  assert.equal(Object.keys(files).filter((p) => /^xl\/worksheets\/[^/]+\.xml$/.test(p)).length, 2, 'two sheet parts left');
  assert.doesNotMatch(strFromU8(files['[Content_Types].xml']), /sheet3\.xml/);
  assert.doesNotMatch(strFromU8(files['xl/_rels/workbook.xml.rels']), /sheet3\.xml/);
});

test('the last sheet shown can be neither hidden nor deleted; hiding the active tab moves it', () => {
  const wb = book('basic.xlsx');
  wb.setSheetHidden(0, true);
  refuses(() => wb.setSheetHidden(1, true), MESSAGES.lastVisible);
  refuses(() => wb.deleteSheet(1), MESSAGES.lastVisible);
  assert.equal(wb.activeSheet, 1);
  wb.deleteSheet(2); // (Hidden: a hidden sheet can go)
  wb.setSheetHidden(0, false);
  const { wb: back, files } = saved(wb, 'hidden.xlsx');
  assert.deepEqual(back.sheets.map((s) => s.state), ['visible', 'visible']);
  assert.match(strFromU8(files['xl/workbook.xml']), /activeTab="1"/);
});

test('tab colours: written first in the sheet, as a colour or a theme colour, and taken off again', () => {
  const wb = book('basic.xlsx');
  wb.setTabColor(0, { rgb: 'FFC00000' });
  wb.setTabColor(1, { theme: 4, tint: 0.3999755851924192 });
  assert.deepEqual([wb.tabColor(0), wb.tabColor(1)], ['#C00000', '#44B3E1']);
  const { files } = saved(wb, 'colours.xlsx');
  assert.match(strFromU8(files['xl/worksheets/sheet1.xml']), /<worksheet [^>]*><sheetPr><tabColor rgb="FFC00000"\/><\/sheetPr><dimension/);
  assert.match(strFromU8(files['xl/worksheets/sheet2.xml']), /<sheetPr><tabColor theme="4" tint="0.3999755851924192"\/><\/sheetPr>/);
  wb.setTabColor(0, null);
  wb.setTabColor(1, null);
  assert.deepEqual([wb.tabColor(0), wb.tabColor(1)], [null, null]);
  assert.doesNotMatch(strFromU8(wb.toBytes() && unzipSync(wb.toBytes())['xl/worksheets/sheet1.xml']), /sheetPr/);
});

test('every workbook these tests wrote passes the package check', () => {
  const r = spawnSync(process.execPath, ['tools/check-ooxml.mjs', dir], { encoding: 'utf8' });
  assert.equal(r.status, 0, r.stdout);
});

test('a snapshot brings back the sheets as they were: added, renamed, moved and deleted ones (undo)', () => {
  const wb = book('basic.xlsx');
  wb.sheet(0); wb.sheet(1); // (parsed, as a shown sheet is)
  const names = () => wb.sheets.map((s) => s.name).join(',');
  const was = names();
  const bytes = unzipSync(wb.toBytes());
  const snap = wb.snapshot();
  const k = wb.addSheet(0, 'Added');
  wb.apply(k, [{ r: 0, c: 0, value: 5 }]);
  wb.renameSheet(1, 'Renamed');
  wb.moveSheet(2, 3);
  wb.deleteSheet(wb.sheetIndexOf('Renamed'));
  assert.notEqual(names(), was);
  const after = wb.snapshot();
  wb.restore(snap);
  assert.equal(names(), was);
  assert.equal(cell(wb, 'Data', 'A1').v, 'Item');
  const again = unzipSync(wb.toBytes());
  for (const name of Object.keys(bytes)) if (name !== 'xl/workbook.xml') assert.deepEqual(again[name], bytes[name], name);
  assert.equal(Object.keys(again).length, Object.keys(bytes).length);
  // (And forward again: redo.)
  wb.restore(after);
  assert.ok(names().startsWith('Added,') && !names().includes('Renamed'));
  assert.equal(wb.sheet(0).cells.get(0).get(0).v, 5);
});
