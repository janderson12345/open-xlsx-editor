import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { build } from './steps.js';
import { serializeXml } from '../src/ooxml/xml.js';

// Format > Default Width... and a row height given to the whole sheet, against what Excel saved
// (test/sizes2-excel.json, tools/excel-compare/excel-sizes2.ps1).
const EXCEL = Object.fromEntries(JSON.parse(fs.readFileSync('test/sizes2-excel.json', 'utf8')).map((c) => [c.name, c]));
const clean = (xml) => xml.replace(/ xmlns(:\w+)?="[^"]*"/g, '').replace(/ x14ac:dyDescent="[^"]*"/g, '').replace(/ spans="[^"]*"/g, '');
const part = (wb, si, name) => [...wb._part(si).doc.documentElement.getElementsByTagNameNS('*', name)].map((n) => clean(serializeXml(n)).replace(/>[\s\S]*$/, '>').replace(/<row([^>]*)>$/, '<row$1>'));
const rowTags = (xml) => [...clean(xml).matchAll(/<row\b[^>]*?\/?>/g)].map((m) => m[0].replace(/\/>$/, '>'));

test('Default Width as Excel writes it: sheetFormatPr defaultColWidth in whole pixels, columns of their own kept', () => {
  for (const [name, chars] of [['default width 12', 12], ['default width 4.5', 4.5]]) {
    const { wb, si } = build({ cells: { A1: 'x' } });
    if (chars === 12) wb.setLines(si, 'col', 2, 2, { width: 20.7109375 });
    wb.setSheetSizes(si, { defaultWidth: chars });
    assert.equal(part(wb, si, 'sheetFormatPr')[0].replace('/>', '>').replace(/>$/, '/>'), clean(EXCEL[name].sheetFormatPr), name);
    assert.equal(wb.sheet(si).format.defaultColWidth, Number(/defaultColWidth="([\d.]+)"/.exec(EXCEL[name].sheetFormatPr)[1]));
  }
});

test('a row height for the whole sheet as Excel writes it: defaultRowHeight with customHeight, each row of the file given it', () => {
  const { wb, si } = build({ cells: { A1: 'x', A3: 'y' } });
  wb.setSheetSizes(si, { rowHeight: 21 });
  const x = EXCEL['whole sheet row height 21'];
  assert.equal(part(wb, si, 'sheetFormatPr')[0].replace(/\/?>$/, '/>'), clean(x.sheetFormatPr));
  assert.deepEqual(rowTags(serializeXml(wb._part(si).doc.documentElement)), rowTags(x.rows));
  wb.setLines(si, 'row', 1, 1, { height: 30 });
  assert.deepEqual(rowTags(serializeXml(wb._part(si).doc.documentElement)), rowTags(EXCEL['whole sheet row height then one row'].rows + '<row r="3" ht="21" customHeight="1">'));
});

test("AutoFit Row of wrapped text: the height found written without customHeight, as Excel's ht=\"150\"", () => {
  const { wb, si } = build({ cells: { A1: 'x' } });
  wb.setLines(si, 'row', 0, 0, { height: 15 });
  wb.setLines(si, 'row', 0, 0, { fitted: 150 });
  assert.deepEqual(rowTags(serializeXml(wb._part(si).doc.documentElement)), rowTags(EXCEL['autofit row with wrapped text'].rows));
});
