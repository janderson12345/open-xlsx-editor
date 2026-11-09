import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseRef } from '../src/ooxml/refs.js';
import { shiftFormula, normalizeFormula, displayFormula } from '../src/calc/tokens.js';
import { parseFormula } from '../src/calc/parse.js';
import { numberText } from '../src/calc/values.js';

const book = (name) => Workbook.fromBytes(new Uint8Array(fs.readFileSync(`test/fixtures/${name}`)));
const value = (wb, si, ref) => {
  const { r, c } = parseRef(ref);
  const cell = wb.sheet(si).cells.get(r)?.get(c);
  return cell ? cell.v : null;
};

test('every formula of formulas.xlsx works out as Excel works it out (test/formulas-excel.json)', () => {
  const excel = JSON.parse(fs.readFileSync('test/formulas-excel.json', 'utf8'));
  const wb = book('formulas.xlsx');
  wb.calculate();
  const wrong = [];
  for (const [sheet, ref, formula, t, v] of excel.cells) {
    const { r, c } = parseRef(ref);
    const cell = wb.sheet(wb.sheetIndexOf(sheet)).cells.get(r)?.get(c);
    const ok = t === 'n' ? cell?.t === 'n' && Math.abs(cell.v - Number(v)) <= 1e-15 * Math.abs(Number(v))
      : t === 's' ? cell?.t === 'str' && cell.v === v
        : t === 'b' ? cell?.t === 'b' && cell.v === v
          : cell?.t === 'e' && cell.v === v;
    if (!ok) wrong.push(`${ref} ${formula}: Excel ${t} ${v}, here ${cell?.t} ${cell?.v}`);
  }
  assert.ok(excel.cells.length > 390);
  assert.deepEqual(wrong, []);
});

test('formulas as tokens: moved (shared formulas), as typed, as shown', () => {
  assert.equal(shiftFormula('B2*2+$B$2+B$2+$B2', 3, 1), 'C5*2+$B$2+C$2+$B5');
  assert.equal(shiftFormula('SUM(A:A,2:2,Data!A1:B2)', 1, 1), 'SUM(B:B,3:3,Data!B2:C3)');
  assert.equal(shiftFormula("'My Data'!A1&\"A1\"", 1, 0), "'My Data'!A2&\"A1\"");
  assert.equal(shiftFormula('A1', -1, 0), '#REF!');
  assert.equal(normalizeFormula('sum(a1:b2)+concat("x",true'), 'SUM(A1:B2)+_xlfn.CONCAT("x",TRUE)');
  assert.equal(normalizeFormula('log10(100)'), 'LOG10(100)');
  assert.equal(displayFormula('_xlfn.XLOOKUP(A1,B:B,C:C)'), 'XLOOKUP(A1,B:B,C:C)');
  assert.throws(() => parseFormula('1+'), /ends too soon/);
  assert.throws(() => parseFormula('SUM(1;2)'), /Expected/);
});

test('numbers as text, as Excel writes them into text', () => {
  assert.deepEqual([1 / 3, -2 / 3, 1e-15, 1e-19, 1.23456789012345e-5, 0.000123456789012345, 1e20, 1.23456789012345e19, 0.1 + 0.2].map(numberText),
    ['0.333333333333333', '-0.666666666666667', '0.000000000000001', '1E-19', '1.23456789012345E-05', '0.000123456789012345', '1E+20',
      '12345678901234500000', '0.3']);
});

test('a change works out again the formulas that read it, on any sheet, shared formulas and chains included', () => {
  const wb = book('formulas.xlsx');
  wb.calculate();
  const calc = wb.sheetIndexOf('Calc');
  const data = wb.sheetIndexOf('Data');
  assert.equal(value(wb, calc, 'D11'), 45); // (D2 1, D3:D10 =D2+1 shared: 1 to 9)
  wb.setCell(calc, 'D2', 11);
  assert.equal(value(wb, calc, 'D10'), 19);
  assert.equal(value(wb, calc, 'D11'), 135);
  const sumRow = wb.sheet(calc).cells; // (=SUM(Data!B2:B9) somewhere in column B)
  const at = [...sumRow].find(([, line]) => line.get(1)?.f === 'SUM(Data!B2:B9)')[0];
  assert.equal(value(wb, calc, `B${at + 1}`), 28);
  wb.setCell(data, 'B3', 105);
  assert.equal(value(wb, calc, `B${at + 1}`), 128);
  // A formula entered, then one that reads it.
  wb.setFormula(calc, 'H1', 'D11*2');
  wb.setFormula(calc, 'H2', 'H1+1');
  assert.equal(value(wb, calc, 'H2'), 271);
  wb.setCell(calc, 'D2', 1);
  assert.equal(value(wb, calc, 'H2'), 91);
});

test('circular references are left alone and reported; a formula that reads them still works', () => {
  const wb = book('basic.xlsx');
  const { circular } = wb.setFormula(0, 'H1', 'H2+1');
  assert.equal(circular.length, 0);
  const second = wb.setFormula(0, 'H2', 'H1+1');
  assert.deepEqual(second.circular.map((e) => `${e.r}:${e.c}`).sort(), ['0:7', '1:7']);
});

test('a function the engine does not know keeps its saved value, and Excel is asked to recalculate on opening', () => {
  const wb = book('basic.xlsx');
  wb.setFormula(0, 'H1', 'NOSUCHFN(1)');
  assert.equal(value(wb, 0, 'H1'), '#NAME?');
  wb.setCell(0, 'B2', 5);
  const files = unzipSync(wb.toBytes());
  assert.match(strFromU8(files['xl/workbook.xml']), /<calcPr calcId="191029" fullCalcOnLoad="1"\/>/);
  // (A workbook whose formulas were all worked out here needs nothing of the kind.)
  const plain = book('basic.xlsx');
  plain.setCell(0, 'B2', 5);
  assert.equal(value(plain, 0, 'B11'), 49);
  assert.doesNotMatch(strFromU8(unzipSync(plain.toBytes())['xl/workbook.xml']), /fullCalcOnLoad/);
});

test('saved: a formula entered is written with its result; results of text are t="str"', () => {
  const wb = book('basic.xlsx');
  wb.setFormula(0, 'H2', 'A2&"!"');
  wb.setFormula(0, 'H3', 'B3*10');
  const xml = strFromU8(unzipSync(wb.toBytes())['xl/worksheets/sheet1.xml']);
  assert.match(xml, /<c r="H2" t="str"><f>A2&amp;"!"<\/f><v>Item 1!<\/v><\/c>/);
  assert.match(xml, /<c r="H3"><f>B3\*10<\/f><v>20<\/v><\/c>/);
});
