import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { serializeXml } from '../src/ooxml/xml.js';
import { readSheet } from '../src/ooxml/sheet.js';
import { ruleAt, valid } from '../src/ooxml/validation.js';
import { setValidation, subtract } from '../src/ooxml/dvedit.js';
import { readCases, build, area } from './steps.js';

// Each rule of tools/excel-compare/dv-cases.json written here as Excel's Data Validation box writes it, compared with
// the dataValidations element Excel saved for it (test/dv-excel.json, less Excel's xr:uid), and every cell judged as
// Excel judged it under the rule written here.
const cases = readCases('tools/excel-compare/dv-cases.json', 'test/dv-excel.json');
const strip = (xml) => xml.replace(/ xr:uid="[^"]*"/g, '').replace(/ xmlns(:\w+)?="[^"]*"/g, '');

for (const c of cases) {
  test(`data validation written: ${c.name}, as Excel writes it`, () => {
    const { wb, si } = build(c.setup);
    const { doc } = wb._part(si);
    const r = c.rule;
    setValidation(doc, [area(r.range)], { type: r.type, op: r.op, f1: r.f1, f2: r.f2, ignoreBlank: r.ignoreBlank });
    const el = [...doc.documentElement.childNodes].find((n) => n.localName === 'dataValidations');
    assert.equal(strip(serializeXml(el)), strip(c.excel.dataValidations));
    wb._parsed.set(wb.sheets[si], { doc, model: readSheet(doc, wb.sharedStrings) });
    wb.calculate();
    const sheet = wb.sheet(si);
    const a = area(c.read);
    const mine = {};
    for (let row = a.r0; row <= a.r1; row++) {
      const cell = sheet.cells.get(row)?.get(0);
      const v = cell?.t === 'e' ? { code: String(cell.v) } : cell?.v ?? null;
      mine[`A${row + 1}`] = valid(wb, si, ruleAt(sheet, row, 0), row, 0, v);
    }
    assert.deepEqual(mine, c.excel.valid);
  });
}

test("a rule for part of another's range splits that range; Clear All takes cells out; the last one gone, the element too", () => {
  assert.deepEqual(subtract({ r0: 0, c0: 0, r1: 9, c1: 3 }, { r0: 3, c0: 1, r1: 4, c1: 2 }), [
    { r0: 0, c0: 0, r1: 2, c1: 3 }, { r0: 5, c0: 0, r1: 9, c1: 3 }, { r0: 3, c0: 0, r1: 4, c1: 0 }, { r0: 3, c0: 3, r1: 4, c1: 3 }]);
  const { wb, si } = build({ cells: { A1: '1' } });
  const { doc } = wb._part(si);
  const all = () => [...doc.documentElement.getElementsByTagNameNS('*', 'dataValidation')].map((e) => `${e.getAttribute('type')}:${e.getAttribute('sqref')}`);
  setValidation(doc, [area('A1:A10')], { type: 'whole', op: 'greater', f1: '0' });
  setValidation(doc, [area('A4:A5')], { type: 'list', f1: 'a,b' });
  assert.deepEqual(all(), ['whole:A1:A3 A6:A10', 'list:A4:A5']);
  setValidation(doc, [area('A1:A10')], null);
  assert.deepEqual(all(), []);
  assert.equal(doc.documentElement.getElementsByTagNameNS('*', 'dataValidations').length, 0);
});

test('a rule read back shows in the box as it was typed: dates, times, lists, references, its messages', async () => {
  const { boxRule } = await import('../src/ooxml/dvedit.js');
  const { build: make } = await import('./steps.js');
  for (const rule of [
    { type: 'date', op: 'greaterEqual', f1: '1/1/2024' },
    { type: 'time', op: 'between', f1: '9:00', f2: '17:00' },
    { type: 'list', op: 'between', f1: 'Yes,No', dropDown: false },
    { type: 'decimal', op: 'equal', f1: '=$B$1', prompt: { show: true, title: 'Rate', text: 'A rate from B1' }, error: { show: true, style: 'warning', title: 'Odd rate', text: 'Not B1' } },
  ]) {
    const { wb, si } = make({ cells: { A1: '1' } });
    wb.setValidation(si, [area('A1:A3')], rule);
    const back = boxRule(ruleAt(wb.sheet(si), 1, 0));
    assert.deepEqual([back.type, back.op, back.f1, back.f2 || undefined], [rule.type, rule.op, rule.f1, rule.f2]);
    if (rule.dropDown === false) assert.equal(back.dropDown, false);
    if (rule.prompt) assert.deepEqual([back.prompt.title, back.prompt.text, back.error.style, back.error.title, back.error.text], ['Rate', 'A rate from B1', 'warning', 'Odd rate', 'Not B1']);
  }
});
