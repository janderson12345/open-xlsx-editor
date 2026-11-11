import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { build, area } from './steps.js';
import { serializeXml } from '../src/ooxml/xml.js';
import { conditionalLooks } from '../src/ooxml/condformat.js';

// Each case of tools/excel-compare/cf-write-cases.json written here as the Conditional Formatting menu adds it,
// compared with what Excel wrote (test/cf-write-excel.json: excel-cf-write.ps1). Excel's COM rules Stop If True and give
// a later rule the lower priority, where the menu does neither: those are set aside here (docs/CONDFORMAT.md).
const CASES = JSON.parse(fs.readFileSync('tools/excel-compare/cf-write-cases.json', 'utf8')).cases;
const EXCEL = JSON.parse(fs.readFileSync('test/cf-write-excel.json', 'utf8'));
const WORDS = ['banana', 'apple', 'cherry', 'band', 'apple', 'kiwi', 'mango', 'banana', 'pear', 'plum'];
const cells = Object.fromEntries(WORDS.flatMap((w, i) => [[`A${i + 1}`, String(i + 1)], [`B${i + 1}`, w]]));
const rules = (xml) => [...xml.matchAll(/<cfRule\b[\s\S]*?(?:\/>|<\/cfRule>)/g)].map((m) => m[0]
  .replace(/ stopIfTrue="1"/, '').replace(/ priority="\d+"/, '').replace(/ dxfId="\d+"/, ' dxfId')).sort();
const dxfs = (xml) => [...xml.matchAll(/<dxf>[\s\S]*?<\/dxf>/g)].map((m) => m[0]).sort();

CASES.forEach((c, k) => {
  test(`conditional format added: ${c.name}, as Excel writes it`, () => {
    const { wb, si } = build({ cells });
    const before = (wb.stylesDoc.documentElement.getElementsByTagNameNS('*', 'dxf') || []).length;
    for (const r of c.rules) wb.setConditional(si, [area(r.range)], r);
    const root = wb._part(si).doc.documentElement;
    const cf = [...root.childNodes].filter((n) => n.localName === 'conditionalFormatting').map((n) => serializeXml(n).replace(/ xmlns(:\w+)?="[^"]*"/g, '')).join('');
    assert.deepEqual(rules(cf), rules(EXCEL[k].conditionalFormatting), 'the rules');
    assert.equal(cf.match(/sqref="[^"]*"/g).join(), EXCEL[k].conditionalFormatting.match(/sqref="[^"]*"/g).join(), 'their ranges');
    const added = [...wb.stylesDoc.documentElement.getElementsByTagNameNS('*', 'dxf')].slice(before).map((n) => serializeXml(n).replace(/ xmlns(:\w+)?="[^"]*"/g, '')).join('');
    assert.deepEqual(dxfs(added), dxfs(EXCEL[k].dxfs), 'their formats');
    if (c.rules.length > 1) {
      // (The menu's order: the rule added last first.)
      const order = [...root.getElementsByTagNameNS('*', 'cfRule')].sort((a, b) => a.getAttribute('priority') - b.getAttribute('priority')).map((r) => r.getAttribute('operator'));
      assert.deepEqual(order, [...c.rules].reverse().map((r) => r.operator));
    }
    // (And drawn: the looks the editor gives the cells come from the rule written.)
    assert.ok(conditionalLooks(wb, si).size > 0 || c.rules.every((r) => r.kind === 'colorScale'));
  });
});

test('Clear Rules from selected cells splits a range; from the entire sheet removes them all', () => {
  const { wb, si } = build({ cells });
  wb.setConditional(si, [area('A1:A10')], CASES[0].rules[0]);
  wb.setConditional(si, [area('A4:A5')], null);
  const sq = () => [...wb._part(si).doc.documentElement.childNodes].filter((n) => n.localName === 'conditionalFormatting').map((n) => n.getAttribute('sqref'));
  assert.deepEqual(sq(), ['A1:A3 A6:A10']);
  wb.setConditional(si, null, null);
  assert.deepEqual(sq(), []);
});
