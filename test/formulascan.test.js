import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { strToU8 } from 'fflate';
import { hasFormulaWithoutResult } from '../src/ooxml/formulascan.js';
import { Workbook } from '../src/ooxml/workbook.js';
import fs from 'node:fs';

// The byte scan Workbook.calculateMissing uses to leave sheets unparsed whose every formula has its saved result.
const sheet = (rows, extra = '') => strToU8(`<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}</sheetData>${extra}</worksheet>`);

test('a formula with its result, or none at all: nothing to work out', () => {
  assert.equal(hasFormulaWithoutResult(sheet('')), false);
  assert.equal(hasFormulaWithoutResult(sheet('<row r="1"><c r="A1"><v>1</v></c></row>')), false);
  assert.equal(hasFormulaWithoutResult(sheet('<row r="1"><c r="A1"><f>B1*2</f><v>4</v></c></row>')), false);
  assert.equal(hasFormulaWithoutResult(sheet('<row r="1"><c r="A1"><f>1&lt;2</f>\n  <v>1</v></c></row>')), false);
  assert.equal(hasFormulaWithoutResult(sheet('<row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">B1</f><v>1</v></c><c r="A2"><f t="shared" si="0"/><v>2</v></c></row>')), false);
  assert.equal(hasFormulaWithoutResult(strToU8('<x:worksheet xmlns:x="m"><x:sheetData><x:row><x:c><x:f>1+1</x:f><x:v>2</x:v></x:c></x:row></x:sheetData></x:worksheet>')), false);
  // (A string result: t="str" with its <v>.)
  assert.equal(hasFormulaWithoutResult(sheet('<row r="1"><c r="A1" t="str"><f>CONCATENATE("a","b")</f><v>ab</v></c></row>')), false);
});

test('a formula saved without its result: to work out', () => {
  assert.equal(hasFormulaWithoutResult(sheet('<row r="1"><c r="A1"><f>B1*2</f></c></row>')), true);
  assert.equal(hasFormulaWithoutResult(sheet('<row r="1"><c r="A1"><f>1</f><v>1</v></c><c r="A2"><f t="shared" si="0"/></c></row>')), true);
  assert.equal(hasFormulaWithoutResult(strToU8('<x:worksheet xmlns:x="m"><x:sheetData><x:row><x:c><x:f>1+1</x:f></x:c></x:row></x:sheetData></x:worksheet>')), true);
  assert.equal(hasFormulaWithoutResult(sheet('<row r="1"><c r="A1"><f>1</f><is><t>x</t></is></c></row>')), true);
});

test('formulas outside sheetData, and tags that only start with f, are not cells\' formulas', () => {
  const ext = '<extLst><ext><x14:dataValidations xmlns:x14="x" xmlns:xm="m"><x14:dataValidation><x14:formula1><xm:f>Lists!$A$1:$A$3</xm:f></x14:formula1></x14:dataValidation></x14:dataValidations></ext></extLst>';
  assert.equal(hasFormulaWithoutResult(sheet('<row r="1"><c r="A1"><v>1</v></c></row>', ext)), false);
  assert.equal(hasFormulaWithoutResult(sheet('', '<conditionalFormatting sqref="A1"><cfRule type="expression"><formula>A1&gt;1</formula></cfRule></conditionalFormatting>')), false);
});

test('calculateMissing leaves a sheet whose formulas all have results unparsed, and still works out one without', () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  wb.calculateMissing();
  // (basic.xlsx's formulas all have their results: no sheet read for it.)
  assert.equal(wb.sheets.filter((info) => wb._parsed.has(info)).length, 0);
  const path = wb.sheets[0].path;
  const xml = wb.pkg.text(path).replace(/<sheetData>/, '<sheetData><row r="100"><c r="A100"><f>1+2</f></c></row>');
  const again = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  again.pkg.setText(path, xml);
  assert.equal(again.calculateMissing(), 1);
  assert.equal(again.sheet(0).cells.get(99).get(0).v, 3);
});
