import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { serializeXml } from '../src/ooxml/xml.js';
import { parseRange, rangeName } from '../src/ooxml/refs.js';
import { breaksRule, ruleAt } from '../src/ooxml/validation.js';
import { defineName } from '../src/ooxml/names.js';

// Data validation rules in Excel's 2010 form (x14:dataValidation in the sheet's extLst), which Excel writes for a rule
// whose formulas read another sheet: test/dv-x14-excel.json, from tools/excel-compare/excel-dv-x14.ps1 (each rule added
// over COM to B2:B5 of a new workbook with a sheet Data and a name Choices, saved, its rules kept; Excel's xr:uid left
// out). Here the same rules are given from the Data Validation box's terms on test/fixtures/refmoves.xlsx, which has
// the sheet Data, after its own rules are cleared.
const EXCEL = JSON.parse(fs.readFileSync('test/dv-x14-excel.json', 'utf8'));
const OUT = 'test/out/dvx14';
fs.mkdirSync(OUT, { recursive: true });
const BOX = {
  'list of another sheet': { type: 'list', f1: '=Data!$A$1:$A$3' },
  'list of this sheet': { type: 'list', f1: '=$H$1:$H$3' },
  'list of a name': { type: 'list', f1: '=Choices' },
  'whole, max from another sheet': { type: 'whole', op: 'between', f1: '1', f2: '=Data!$B$1' },
  'whole, min from another sheet': { type: 'whole', op: 'greater', f1: '=Data!$B$1' },
  'custom reading another sheet': { type: 'custom', f1: '=B2>Data!B1' },
  'custom reading this sheet': { type: 'custom', f1: '=B2>C2' },
  'length from another sheet': { type: 'length', op: 'equal', f1: '=Data!$B$1' },
  'list of another sheet, with a sparkline': { type: 'list', f1: '=Data!$A$1:$A$3' },
};
const B2B5 = [parseRange('B2:B5')];
const strip = (xml) => (xml || '').replace(/ xmlns(:\w+)?="[^"]*"/g, '');
/** refmoves.xlsx with its rules cleared (its sparklines kept), a name Choices defined. */
function book() {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/refmoves.xlsx')));
  wb.setValidation(0, [parseRange('A1:XFD1048576')], null);
  defineName(wb, { name: 'Choices', refersTo: '=Data!$A$1:$A$3' });
  return wb;
}
const parts = (wb) => {
  const xml = serializeXml(wb._part(0).doc);
  return {
    main: /<dataValidations[\s\S]*?<\/dataValidations>/.exec(xml)?.[0] || '',
    ext: /<x14:dataValidations[\s\S]*?<\/x14:dataValidations>/.exec(xml)?.[0] || '',
    uris: [...xml.matchAll(/<ext uri="([^"]+)"/g)].map((m) => m[1]),
  };
};

test('clearing every rule takes the Excel 2010 one and its ext out too, the sparklines\' ext kept', () => {
  const got = parts(book());
  assert.equal(got.main, '');
  assert.equal(got.ext, '');
  assert.deepEqual(got.uris, ['{05C60535-1F16-4fd2-B633-F4F36F0B64E0}']);
});

EXCEL.forEach((excel, k) => {
  test(`a rule written as Excel writes it: ${excel.name}`, () => {
    const wb = book();
    wb.setValidation(0, B2B5, { ignoreBlank: true, ...BOX[excel.name] });
    const got = parts(wb);
    assert.equal(strip(got.main), strip(excel.main));
    assert.equal(strip(got.ext), strip(excel.ext));
    // (The rules' ext before the sparklines' when there are both, as Excel orders them.)
    if (excel.ext) assert.equal(got.uris[0], excel.uris[0]);
    // Read back as the rule it is, of the form it is in.
    const rule = ruleAt(wb.sheet(0), 1, 1);
    assert.ok(rule, 'read');
    assert.equal(rule.ext, !!excel.ext);
    fs.writeFileSync(`${OUT}/${k + 1}.xlsx`, wb.toBytes());
  });
});

test('a rule of the 2010 form judges its cells, lists from another sheet, and is edited and cleared by the box', () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/refmoves.xlsx')));
  // (E10:E15's list of Data!$A$1:$A$3: p, q, r. E12's zz isn't in it.)
  assert.equal(breaksRule(wb, 0, 11, 4), true);
  assert.equal(breaksRule(wb, 0, 9, 4), false);
  // The box gives E10:E11 a rule of this sheet: they leave the 2010 rule, which keeps E12:E15.
  wb.setValidation(0, [parseRange('E10:E11')], { type: 'whole', op: 'between', f1: '1', f2: '9' });
  const ext = wb.sheet(0).validations.filter((v) => v.ext);
  assert.deepEqual(ext.map((v) => v.ranges.map(rangeName).join(' ')), ['E12:E15']);
  // Clear All over the rest: the rule goes, and with it the ext.
  wb.setValidation(0, [parseRange('E12:E15')], null);
  assert.equal(wb.sheet(0).validations.filter((v) => v.ext).length, 0);
  assert.ok(!serializeXml(wb._part(0).doc).includes('{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}'));
});

test('a cut moves a rule of the 2010 form with its cells; a paste on the same sheet joins it', () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/refmoves.xlsx')));
  wb.moveValidation(0, parseRange('E10:E11'), 0, 4); // (E10:E11 cut to I10:I11)
  const rule = () => wb.sheet(0).validations.find((v) => v.ext);
  assert.deepEqual(rule().ranges.map(rangeName), ['E12:E15', 'I10:I11']);
  assert.ok(wb.addToValidation(0, { first: { r: 11, c: 4 }, type: 'list', formula1: 'Data!$A$1:$A$3' }, [parseRange('K2:K3')]));
  assert.deepEqual(rule().ranges.map(rangeName), ['E12:E15', 'I10:I11', 'K2:K3']);
});
