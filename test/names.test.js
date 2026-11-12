import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { validName, defineName, editName, deleteName, listNames, NAME_MESSAGES } from '../src/ooxml/names.js';

// test/names-excel.json (tools/excel-compare/excel-names.ps1): which names Excel took with Names.Add, and the
// definedNames it saved for names added out of order (a local one, a hidden one).
const EXCEL = JSON.parse(fs.readFileSync('test/names-excel.json', 'utf8'));
const book = () => Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/formulas.xlsx')));

test('the names Excel takes, and the ones it refuses', () => {
  const wrong = EXCEL.verdicts.filter((v) => validName(v.name) !== v.ok).map((v) => `${v.name.slice(0, 20)}: Excel ${v.ok}`);
  assert.deepEqual(wrong, []);
  assert.ok(EXCEL.verdicts.length >= 30);
});

test('names saved in Excel\'s order: by name, capitals aside, built-in ones by their own name, a sheet\'s first', () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  for (const d of [...wb.definedNames]) deleteName(wb, { name: d.name, sheet: d.sheet ?? null });
  const add = (name, refersTo, sheet = null, hidden = false) => defineName(wb, { name, refersTo, sheet, hidden });
  add('Zeta', '=Data!$C$3');
  add('alpha', '=Data!$A$1:$B$2');
  add('Local', '=Notes!$D$4', 1);
  add('Beta', '=0.07');
  add('Gamma', '=SUM(Data!$A$1:$A$3)*2');
  add('Hidden1', '=Data!$E$5', null, true);
  add('_under', '="text"');
  add('Dup', '=Notes!$A$1', 1);
  add('dup', '=Data!$B$1');
  const ours = new TextDecoder().decode(Workbook.fromBytes(wb.toBytes()).pkg.files['xl/workbook.xml']).match(/<definedNames>[\s\S]*<\/definedNames>/)[0];
  // (Excel's sheet names were Sheet1 and Sheet2, and it had a print area we didn't make.)
  const theirs = EXCEL.definedNames.replace(/Sheet1!/g, 'Data!').replace(/Sheet2!/g, 'Notes!').replace(/<definedName name="_xlnm\.Print_Area"[^>]*>[^<]*<\/definedName>/, '');
  assert.equal(ours, theirs);
});

test('Define Name, Edit Name and Delete as the Name Manager does them: formulas follow, #NAME? when gone', () => {
  const wb = book();
  const data = wb.sheetIndexOf('Data');
  wb.apply(data, [{ r: 20, c: 0, formula: 'Tax*2' }]);
  assert.equal(wb.sheet(data).cells.get(20).get(0).v, '#NAME?');
  defineName(wb, { name: 'Tax', refersTo: '=0.25' });
  assert.equal(wb.sheet(data).cells.get(20).get(0).v, 0.5);
  assert.throws(() => defineName(wb, { name: 'tax', refersTo: '=1' }), { message: NAME_MESSAGES.exists });
  assert.throws(() => defineName(wb, { name: 'my tax', refersTo: '=1' }), { message: NAME_MESSAGES.syntax });
  assert.throws(() => defineName(wb, { name: 'Ok', refersTo: '=1+*2' }), { message: NAME_MESSAGES.formula });
  // (A sheet's own Tax beside the workbook's: allowed, as in Excel.)
  defineName(wb, { name: 'Tax', refersTo: '=1', sheet: data });
  editName(wb, { name: 'Tax', sheet: null }, { name: 'Rate2', refersTo: '=0.5' });
  assert.equal(wb.sheet(data).cells.get(20).get(0).f, 'Tax*2'); // (Data's own Tax still read there)
  deleteName(wb, { name: 'Tax', sheet: data });
  assert.equal(wb.sheet(data).cells.get(20).get(0).v, '#NAME?');
  wb.apply(data, [{ r: 21, c: 0, formula: 'Rate2*4' }]);
  editName(wb, { name: 'Rate2' }, { name: 'Rate3' });
  assert.equal(wb.sheet(data).cells.get(21).get(0).f, 'Rate3*4');
  assert.equal(wb.sheet(data).cells.get(21).get(0).v, 2);
  // (The workbook's LAMBDAs kept in names, Twice, Hyp and WithRate, aside.)
  assert.deepEqual(listNames(wb).filter((n) => !/LAMBDA/.test(n.refersTo)).map((n) => `${n.name}|${n.scope}|${n.refersTo}`), ['Items|Workbook|=Data!$A$2:$A$9', 'Rate|Workbook|=Data!$C$2', 'Rate3|Workbook|=0.5']);
  const again = Workbook.fromBytes(wb.toBytes());
  assert.equal(again.definedName('Rate3', null).text, '0.5');
});
