import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { digitWidth, deriveDigitWidth, colPad, setDigitMeasure, MEASURED_MDW } from '../src/ooxml/digitwidth.js';
import { colWidthToPx, charsToColWidth, colWidthToChars, defaultColPx, colChars } from '../src/ooxml/units.js';
import { typedWidth } from '../src/ooxml/sizes.js';
import { blankWorkbook } from '../src/ooxml/csv.js';
import { sheetGeometry } from '../src/editor/geometry.js';

// Excel 365 over COM (tools/excel-compare/excel-digitwidth.ps1): 41 Normal fonts and sizes, columns of 1, 2, 5, 10 and
// 20 characters in points and as saved, the default width, and AutoFit of ten zeros.
const EXCEL = JSON.parse(fs.readFileSync('test/digitwidth-excel.json', 'utf8')).fonts;
const px = (pt) => Math.round(pt / 0.75);
const mdwOf = (f) => (px(f.pt[4]) - px(f.pt[3])) / 10;

test("the maximum digit width of every Normal font measured is Excel's", () => {
  for (const f of EXCEL) assert.equal(digitWidth(f.font, f.size), mdwOf(f), `${f.font} ${f.size}`);
  assert.equal(EXCEL.length, 41);
  // The ones asked for: Arial 10, Times New Roman 12, Cambria 11, Calibri 10 and 12, Aptos 11, Segoe UI 9, Verdana 10,
  // Courier New 10.
  assert.deepEqual([['Arial', 10], ['Times New Roman', 12], ['Cambria', 11], ['Calibri', 10], ['Calibri', 12], ['Aptos', 11], ['Segoe UI', 9], ['Verdana', 10], ['Courier New', 10]].map(([n, s]) => digitWidth(n, s)), [7, 8, 8, 7, 8, 8, 6, 8, 8]);
});

test('columns of n characters, the default width and the stored widths are as Excel has them, for every font', () => {
  for (const f of EXCEL) {
    const mdw = digitWidth(f.font, f.size);
    const name = `${f.font} ${f.size}`;
    f.chars.forEach((n, k) => {
      assert.equal(typedWidth(n, mdw), f.stored[k], `${name}: ${n} characters stored`);
      assert.equal(charsToColWidth(n, mdw), f.stored[k], `${name}: ${n} characters`);
      assert.equal(colWidthToPx(f.stored[k], mdw), px(f.pt[k]), `${name}: ${n} characters in pixels`);
      assert.equal(colWidthToChars(f.stored[k], mdw), n, `${name}: ${n} characters shown`);
    });
    assert.equal(defaultColPx({}, mdw), px(f.defaultPt), `${name}: default column`);
    assert.equal(Math.round(((defaultColPx({}, mdw) - colPad(mdw)) / mdw) * 100) / 100, f.standardWidth, `${name}: StandardWidth`);
    // AutoFit of ten digits: their width and the padding and 2px more.
    assert.equal(px(f.fitPt), 10 * mdw + colPad(mdw) + 2, `${name}: AutoFit`);
  }
  assert.equal(colChars(64, 7), 8);
  assert.equal(colChars(80, 9), 8);
});

test("worked out from the font alone, 32 of the 36 installed fonts are Excel's (the other 4 a pixel off: hinting)", () => {
  const off = [];
  let n = 0;
  for (const f of EXCEL) {
    if (/aptos|ms sans/i.test(f.font)) continue;
    n++;
    if (deriveDigitWidth(f.font, f.size) !== mdwOf(f)) off.push(`${f.font} ${f.size}`);
  }
  assert.equal(n, 36);
  assert.deepEqual(off, ['Calibri 8', 'Calibri 14', 'Arial 14', 'Times New Roman 10']);
  // A font not in the table goes by the browser's measure, rounded to the nearest pixel.
  setDigitMeasure((family, size) => (family === 'Fancy' ? size * 0.62 : null));
  try {
    assert.equal(digitWidth('Fancy', 11), 9); // 14.67 x 0.62 = 9.09
    assert.equal(digitWidth('Fancy', 10), 8); // 13.33 x 0.62 = 8.27
    assert.equal(digitWidth('Arial', 12), MEASURED_MDW.arial[12]);
  } finally {
    setDigitMeasure(null);
  }
});

test("a workbook's columns follow its Normal font: grid, typed widths and the default width", () => {
  const wb = blankWorkbook();
  assert.equal(wb.mdw, 7);
  wb.styles.normal = { ...wb.styles.normal, name: 'Arial', size: 12, scheme: null };
  assert.equal(wb.mdw, 9);
  wb.setLines(0, 'col', 1, 1, { width: typedWidth(10, wb.mdw) });
  const g = sheetGeometry(wb.sheet(0));
  // (Excel: Arial 12's default column 80px, 10 characters 97px.)
  assert.equal(g.cols.size(0), 80);
  assert.equal(g.cols.size(1), 97);
  wb.setSheetSizes(0, { defaultWidth: 5 });
  assert.equal(sheetGeometry(wb.sheet(0)).cols.size(0), 52);
});
