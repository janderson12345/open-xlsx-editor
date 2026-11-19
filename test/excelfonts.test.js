import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { EXCEL_FONTS, EXCEL_FONT_FROM, excelFontMetric } from '../src/ooxml/excelfonts.js';
import { fontRowPx } from '../src/editor/fonts.js';
import { digitWidth } from '../src/ooxml/digitwidth.js';

// Excel's row heights and digit widths by font and size, measured over COM (tools/excel-compare/excel-fontmetrics.ps1,
// test/fontmetrics-excel.json) and looked up by fonts.js fontRowPx and digitwidth.js digitWidth.
const measured = JSON.parse(fs.readFileSync('test/fontmetrics-excel.json', 'utf8').trimStart()).fonts;

test('the table is the measurements: every font, every half point, rows regular and bold, digit widths', () => {
  let n = 0;
  for (const [name, v] of Object.entries(measured)) {
    const t = EXCEL_FONTS[name.toLowerCase()];
    assert.ok(t, name);
    for (const [s, x] of Object.entries(v.sizes)) {
      const k = Math.round((Number(s) - EXCEL_FONT_FROM) * 2);
      assert.equal(t.row[k], Math.round(x[0] / 0.75), `${name} ${s}`);
      assert.equal(t.bold[k], Math.round(x[1] / 0.75), `${name} ${s} bold`);
      assert.equal(t.mdw[k], Math.round((x[3] - x[2]) / 0.75 / 9), `${name} ${s} mdw`);
      n++;
    }
  }
  assert.equal(n, 25 * 61);
});

test('row heights as Excel gives them where the old formula was off, and bold where it differs', () => {
  const pt = (name, size, bold) => fontRowPx(name, (size * 4) / 3, bold) * 0.75;
  assert.equal(pt('Calibri', 10), 12.75); // (13.5 by the formula)
  assert.equal(pt('Calibri', 9), 12);
  assert.equal(pt('Tahoma', 9), 11.25); // (12.75)
  assert.equal(pt('Verdana', 10), 12.75);
  assert.equal(pt('Lato', 11), 18); // (15: Lato wasn't in the metrics table, Calibri's were used)
  assert.equal(pt('Arial', 11), 14.25);
  assert.equal(pt('Arial', 11, true), 15);
  assert.equal(pt('Calibri', 11), 15);
  // (A size between half points: to the nearest; a font not measured: the formula, as before.)
  assert.equal(pt('Calibri', 10.4), pt('Calibri', 10.5));
  assert.ok(pt('Wingdings', 11) > 0);
  assert.equal(excelFontMetric('Wingdings', 11, 'row'), null);
  assert.equal(excelFontMetric('Calibri', 72, 'row'), null);
});

test('digit widths for the Normal fonts the old table lacked', () => {
  assert.equal(digitWidth('Lato', 11), 9); // (columns 12.5% narrow before)
  assert.equal(digitWidth('Aptos Narrow', 12), 8); // (11% wide)
  assert.equal(digitWidth('Calibri', 11), 7);
  assert.equal(digitWidth('Arial', 10), 7);
});
