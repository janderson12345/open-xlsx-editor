import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { serializeXml, unescapeXml } from '../src/ooxml/xml.js';
import { doStep } from './refmoves-steps.js';

// test/cfmoves-excel.json (tools/excel-compare/excel-cfmoves.ps1 on cfmove-cases.mjs, cut down by refmove-results.mjs):
// Excel's conditional formats of both forms after each insert, delete or sort, on the sheet test/fixtures/cfmoves.xlsx
// holds (data bars on B2:B10 in both forms, 3 stars on C2:C10 and =D2>Data!$A$1 on D2:D10 in Excel's 2010 form,
// greater than 5 on E2:E10).
const EXCEL = JSON.parse(fs.readFileSync('test/cfmoves-excel.json', 'utf8'));
// (Each result saved as test/out/cfmoves/<case>.xlsx, for check-ooxml and for Excel to read back: excel-cfmoves.ps1 -Books.)
const OUT = 'test/out/cfmoves';
fs.mkdirSync(OUT, { recursive: true });
const book = () => Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/cfmoves.xlsx')));

/** The sheet's formats in refmove-results.mjs's terms: "ranges | rule types | formulas", sorted, by form. */
function stateOf(wb) {
  const xml = serializeXml(wb._part(0).doc);
  const main = [...xml.matchAll(/<conditionalFormatting\b[^>]*\ssqref="([^"]*)"[^>]*>([\s\S]*?)<\/conditionalFormatting>/g)].map((m) => [
    m[1], [...m[2].matchAll(/<cfRule type="(\w+)"/g)].map((x) => x[1]).join('+'), ...[...m[2].matchAll(/<formula>([^<]*)<\/formula>/g)].map((x) => unescapeXml(x[1])),
  ].join(' | ')).sort();
  const ext = [...xml.matchAll(/<x14:conditionalFormatting\b[^>]*>([\s\S]*?)<\/x14:conditionalFormatting>/g)].map((m) => [
    /<xm:sqref>([^<]*)/.exec(m[1])[1], [...m[1].matchAll(/<x14:cfRule type="(\w+)"/g)].map((x) => x[1]).join('+'), ...[...m[1].matchAll(/<xm:f>([^<]*)<\/xm:f>/g)].map((x) => unescapeXml(x[1])),
  ].join(' | ')).sort();
  return { main, ext };
}

EXCEL.forEach((excel, k) => {
  test(`conditional formats, ${excel.name}: as Excel moves them`, () => {
    const wb = book();
    doStep(wb, excel.step);
    const got = stateOf(wb);
    assert.deepEqual(got.main, excel.main, 'the main form');
    assert.deepEqual(got.ext, excel.ext, 'the 2010 form');
    const bytes = wb.toBytes();
    fs.writeFileSync(`${OUT}/${k + 1}.xlsx`, bytes);
    assert.deepEqual(stateOf(Workbook.fromBytes(bytes)), got);
  });
});

test('a main-form formula rule keeps reading from its first cell when its first rows are deleted, as the 2010 one did', () => {
  // (E2:E10's rule made =E2>5: with rows 2:3 deleted, E2:E8 reads E2>5, as Excel's D2:D10 kept D2>Data!$A$1; it was
  // #REF! before, the deleted E2 taken as a reference.)
  const wb = book();
  const doc = wb._part(0).doc;
  const rule = [...doc.getElementsByTagNameNS('*', 'cfRule')].find((el) => el.getAttribute('type') === 'cellIs');
  rule.setAttribute('type', 'expression');
  rule.removeAttribute('operator');
  rule.getElementsByTagNameNS('*', 'formula')[0].textContent = 'E2>5';
  wb.changeLines(0, 'row', 1, -2);
  assert.ok(stateOf(wb).main.includes('E2:E8 | expression | E2>5'), JSON.stringify(stateOf(wb).main));
});
