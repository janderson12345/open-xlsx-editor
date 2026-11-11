import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseXml } from '../src/ooxml/xml.js';
import { deleteRule, moveRule, setStopIfTrue, setRuleRanges, rulePriorities } from '../src/ooxml/cfmanage.js';
import { parseRange } from '../src/ooxml/refs.js';

// The Rules Manager's changes as Excel makes them: test/cfmanage-excel.json holds a sheet's four rules as Excel saved
// them, and the same after each change (tools/excel-compare/excel-cf-manage.ps1). Excel's starting rules are put on a
// sheet here, the change made, and the sheet's rules compared with Excel's text.
const excel = JSON.parse(fs.readFileSync('test/cfmanage-excel.json', 'utf8')).cases;
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
function withExcelRules() {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const doc = wb._part(0).doc;
  const root = doc.documentElement;
  for (const el of [...root.childNodes].filter((n) => n.nodeType === 1 && ['conditionalFormatting', 'extLst'].includes(n.localName))) root.removeChild(el);
  const add = (xml, before) => { const el = doc.importNode(parseXml(`<w xmlns="${NS}">${xml}</w>`).documentElement.firstChild, true); root.insertBefore(el, before); };
  const at = [...root.childNodes].find((n) => n.nodeType === 1 && ['dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'drawing', 'legacyDrawing', 'tableParts'].includes(n.localName));
  for (const cf of excel.start.cf) add(cf, at || null);
  add(excel.start.ext, null);
  wb.dirty.add(wb.sheets[0].path);
  return wb;
}
/** Each workbook's made-up ids out of the way (Excel made each case's afresh). */
const ID = /\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}/g;
const bare = (x) => ({ cf: x.cf.map((c) => c.replace(ID, '{id}')), ext: x.ext.replace(ID, '{id}') });
const rulesText = (wb) => {
  const xml = strFromU8(unzipSync(wb.toBytes())['xl/worksheets/sheet1.xml']);
  return bare({ cf: xml.match(/<conditionalFormatting\b[\s\S]*?<\/conditionalFormatting>/g) || [], ext: /<extLst>(?:(?!<extLst>)[\s\S])*<\/extLst>(?=<\/worksheet>)/.exec(xml)?.[0] || '' });
};
const CHANGES = {
  deleteSecond: (wb) => deleteRule(wb, 0, 2),
  thirdUp: (wb) => moveRule(wb, 0, 3, 2),
  firstLast: (wb) => moveRule(wb, 0, 1, 'last'),
  stopFormula: (wb) => setStopIfTrue(wb, 0, 4, true),
  widen: (wb) => setRuleRanges(wb, 0, 4, [parseRange('A1:B10')]),
};

test('each Rules Manager change made as Excel makes it (5 changes)', () => {
  assert.deepEqual(rulesText(withExcelRules()), bare(excel.start));
  const wrong = [];
  for (const [name, change] of Object.entries(CHANGES)) {
    const wb = withExcelRules();
    change(wb);
    const got = rulesText(wb);
    const want = bare(excel[name]);
    if (JSON.stringify(got.cf) !== JSON.stringify(want.cf)) wrong.push(`${name}: Excel ${want.cf.join('')}\n   here ${got.cf.join('')}`);
    if (got.ext !== want.ext) wrong.push(`${name} extLst: Excel ${want.ext}\n   here ${got.ext}`);
  }
  assert.deepEqual(wrong, []);
});

test('Stop If True off again, the priorities listed in order, and a rule moved back', () => {
  const wb = withExcelRules();
  setStopIfTrue(wb, 0, 1, false);
  assert.match(rulesText(wb).cf[0], /<cfRule type="cellIs" dxfId="1" priority="1" operator="greaterThan">/);
  moveRule(wb, 0, 3, 2);
  assert.deepEqual(rulePriorities(wb, 0), [1, 2, 3, 5]);
  const again = Workbook.fromBytes(wb.toBytes());
  assert.deepEqual(again.sheet(0).condFormats.flatMap((c) => c.rules).sort((a, b) => a.priority - b.priority).map((r) => r.type), ['cellIs', 'iconSet', 'dataBar', 'expression']);
});
