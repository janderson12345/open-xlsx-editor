import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { Workbook } from '../src/ooxml/workbook.js';
import { XlError, Arr } from '../src/calc/values.js';

// Excel 365's answers for 1972 text formulas over emoji and other characters past U+FFFF (test/emoji-excel.json, from
// tools/excel-compare/emoji-cases.mjs and excel-formulas.ps1): `units` in a workbook without the function
// compatibility version (as this editor writes them), `points` in the same workbook with setVersion 2 (as Excel makes
// them). The last 600 are random cases the rules weren't worked out from.
const data = JSON.parse(fs.readFileSync('test/emoji-excel.json', 'utf8'));
const basic = new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx'));
const COMPAT = '<extLst><ext uri="{D14903EA-33C4-47F7-8F05-3474C54BE107}" xmlns:xlwcv="http://schemas.microsoft.com/office/spreadsheetml/2024/workbookCompatibilityVersion"><xlwcv:version setVersion="VER"/></ext></extLst>';

/** basic.xlsx, with workbook.xml asking for function compatibility version `v` (none when null). */
function book(v) {
  if (v == null) return Workbook.fromBytes(basic);
  const files = unzipSync(basic);
  files['xl/workbook.xml'] = strToU8(strFromU8(files['xl/workbook.xml']).replace('</workbook>', `${COMPAT.replace('VER', v)}</workbook>`));
  return Workbook.fromBytes(zipSync(files));
}

const plain = (v) => {
  if (v instanceof Arr) v = v.rows[0][0];
  if (Array.isArray(v)) v = v[0]?.[0];
  if (v instanceof XlError) return { e: v.code ?? String(v) };
  return v ?? null;
};
const same = (ours, excel) => (excel && typeof excel === 'object' ? ours?.e === excel.e
  : excel === null ? ours === null : typeof excel === 'number' ? Math.abs(ours - excel) < 1e-9 : ours === excel);

for (const [mode, v] of [['units', null], ['points', '2']]) {
  test(`text functions count emoji as Excel does ${mode === 'units' ? 'without' : 'with'} function compatibility version 2`, () => {
    const wb = book(v);
    const si = wb.sheets.length - 1;
    const wrong = [];
    data.formulas.forEach((f, i) => {
      const ours = plain(wb.calc.formulaAt(si, 900 + i, 20, f.slice(1)));
      if (!same(ours, data[mode][i])) wrong.push(`${f}: Excel ${JSON.stringify(data[mode][i])}, ours ${JSON.stringify(ours)}`);
    });
    assert.deepEqual(wrong, []);
  });
}

test('the function compatibility version is read from workbook.xml, 1 where it isn\'t given', () => {
  assert.equal(book(null).calc.formulaAt(0, 900, 20, 'LEN(UNICHAR(128512))'), 2);
  assert.equal(book('1').calc.formulaAt(0, 900, 20, 'LEN(UNICHAR(128512))'), 2);
  assert.equal(book('2').calc.formulaAt(0, 900, 20, 'LEN(UNICHAR(128512))'), 1);
  assert.equal(book('3').calc.formulaAt(0, 900, 20, 'LEN(UNICHAR(128512))'), 1);
});
