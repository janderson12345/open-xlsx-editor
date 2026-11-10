import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseXml, kid } from '../src/ooxml/xml.js';
import { readSheet } from '../src/ooxml/sheet.js';
import { ruleAt, valid } from '../src/ooxml/validation.js';
import { readCases, build, area } from './steps.js';

// Data validation, each case as Excel judged it (Validation.Add over COM, then Range.Validation.Value for each cell:
// tools/excel-compare/excel-dv.ps1 on dv-cases.json, from dv-cases.mjs, giving test/dv-excel.json): the case's sheet
// built here with the dataValidations element Excel saved for it, and every cell's value judged as Excel judged it.
const cases = readCases('tools/excel-compare/dv-cases.json', 'test/dv-excel.json');
const X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

for (const c of cases) {
  test(`data validation: ${c.name}, as Excel`, () => {
    const { wb, si } = build(c.setup);
    const { doc } = wb._part(si);
    const root = doc.documentElement;
    const el = parseXml(c.excel.dataValidations.replace('<dataValidations ', `<dataValidations xmlns="${X}" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" `)).documentElement;
    root.insertBefore(doc.importNode(el, true), kid(root, 'pageMargins') || null);
    wb._parsed.set(wb.sheets[si], { doc, model: readSheet(doc, wb.sharedStrings) });
    wb.calculate();
    const sheet = wb.sheet(si);
    const a = area(c.read);
    const mine = {};
    for (let r = a.r0; r <= a.r1; r++) {
      const cell = sheet.cells.get(r)?.get(0);
      const v = cell?.t === 'e' ? { code: String(cell.v) } : cell?.v ?? null;
      mine[`A${r + 1}`] = valid(wb, si, ruleAt(sheet, r, 0), r, 0, v);
    }
    assert.deepEqual(mine, c.excel.valid);
  });
}
