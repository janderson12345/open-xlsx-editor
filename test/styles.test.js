import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { fontFamily, fontOfXf } from '../src/ooxml/styles.js';
import { fontStack, cellLineHeight } from '../src/editor/fonts.js';
import { printRowPx } from '../src/ooxml/units.js';
import { PROBE_FONTS } from './fixtures/make-fixtures.mjs';

const open = (f) => Workbook.fromBytes(new Uint8Array(fs.readFileSync(`test/fixtures/${f}`)));

test('fonts and cell formats from styles.xml, theme fonts from the theme part', () => {
  const wb = open('fonts.xlsx');
  assert.equal(wb.styles.fonts.length, PROBE_FONTS.length);
  assert.deepEqual(wb.theme.fonts, { major: 'Calibri', minor: 'Calibri' });
  assert.equal(wb.theme.colors[4], '4472C4');
  const bold = fontOfXf(wb.styles, 1);
  assert.deepEqual([bold.name, bold.size, bold.b, bold.i], ['Calibri', 11, true, false]);
  assert.deepEqual([fontOfXf(wb.styles, 7).name, fontOfXf(wb.styles, 7).i], ['Cambria', true]);
  assert.equal(fontOfXf(wb.styles, 999), wb.styles.fonts[0]); // (out of range: format 0, as Excel)
});

test('a scheme font is the theme\'s whatever its name says; without a theme part, Excel 365\'s default (Aptos Narrow)', () => {
  const basic = open('basic.xlsx');
  assert.equal(basic.styles.normal.name, 'Calibri');
  assert.equal(fontFamily(basic.styles.normal, basic.theme), 'Aptos Narrow');
  const fonts = open('fonts.xlsx');
  assert.equal(fontFamily(fontOfXf(fonts.styles, 6), fonts.theme), 'Cambria');
});

test('font stacks: metric-compatible stand-ins, Calibri for Aptos Narrow, a generic family last', () => {
  assert.equal(fontStack('Calibri'), 'Calibri, Carlito, sans-serif');
  assert.equal(fontStack('Aptos Narrow'), "'Aptos Narrow', Calibri, Carlito, sans-serif");
  assert.equal(fontStack('Times New Roman'), "'Times New Roman', 'Liberation Serif', Tinos, serif");
  assert.equal(fontStack('Courier New'), "'Courier New', 'Liberation Mono', Cousine, monospace");
  // A font not in the metrics table measures as its stand-in (the font the browser draws).
  assert.equal(cellLineHeight('Aptos Narrow', 14.667), cellLineHeight('Calibri', 14.667));
});

test('printed rows by the Normal style\'s font: Aptos Narrow in 600 dpi device pixels, Calibri x 1.288, others in points', () => {
  assert.equal(printRowPx(15, 'Aptos Narrow'), 127 / 6.25);
  assert.equal(printRowPx(15, 'Calibri'), 15 * 1.288);
  assert.equal(printRowPx(15, 'Verdana'), 20);
});
