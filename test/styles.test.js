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
  // A font not in the metrics table is laid out by its stand-in's metrics (the font the browser draws).
  assert.equal(cellLineHeight('Aptos Narrow', 14.667), cellLineHeight('Calibri', 14.667));
});

test('printed rows by the Normal style\'s font: Aptos Narrow in 600 dpi device pixels, Calibri x 1.291, others in points', () => {
  assert.equal(printRowPx(15, 'Aptos Narrow'), 127 / 6.25);
  assert.equal(printRowPx(15, 'Calibri'), 121 / 6.25);
  assert.equal(printRowPx(15, 'Verdana'), 20);
});

// Excel's printed colours for test/fixtures/fills.xlsx (its PDF, read with tools/excel-compare/shapes.cjs).
test('colours as Excel draws them: theme indexes with the first two pairs swapped, tints in integer HLS, indexed', async () => {
  const { resolveColor, applyTint } = await import('../src/ooxml/theme.js');
  const theme = open('fills.xlsx').theme;
  assert.deepEqual([0, 1, 2, 3, 4, 9].map((t) => resolveColor({ theme: t }, theme)), ['#FFFFFF', '#000000', '#E7E6E6', '#1F3864', '#4472C4', '#70AD47']);
  const tints = [['4472C4', -0.5, '203764'], ['4472C4', -0.25, '305496'], ['4472C4', 0.4, '8EA9DB'], ['4472C4', 0.6, 'B4C6E7'],
    ['4472C4', 0.8, 'D9E1F2'], ['000000', 0.5, '808080'], ['FFFFFF', -0.15, 'D9D9D9'], ['70AD47', -0.249977111117893, '548235']];
  for (const [rgb, t, want] of tints) assert.equal(applyTint(rgb, t), want, `${rgb} ${t}`);
  assert.deepEqual([10, 13, 22, 30, 64].map((i) => resolveColor({ indexed: i }, theme)), ['#FF0000', '#FFFF00', '#C0C0C0', '#0066CC', '#000000']);
  assert.equal(resolveColor({ rgb: 'FF00B050' }, theme), '#00B050');
  assert.equal(resolveColor({ auto: true }, theme), null);
});

test('fills and borders from styles.xml', () => {
  const { styles } = open('fills.xlsx');
  const xf = (s) => styles.cellXfs[s];
  assert.deepEqual(styles.fills[xf(1).fillId].pattern, 'solid');
  assert.deepEqual(styles.fills[xf(1).fillId].fg.rgb, 'FFFF0000');
  const box = styles.borders[styles.cellXfs[styles.cellXfs.length - 1].borderId];
  assert.deepEqual(['left', 'right', 'top', 'bottom'].map((k) => box[k].style), ['thick', 'thick', 'thick', 'thick']);
  assert.equal(box.left.color.indexed, 12);
});

test('alignment from cell formats: horizontal, vertical, wrap, indent', () => {
  const { styles } = open('align.xlsx');
  const a = (s) => styles.cellXfs[s].align;
  assert.deepEqual(a(0), { horizontal: 'general', vertical: 'bottom', wrap: false, indent: 0 });
  assert.deepEqual(a(2), { horizontal: 'center', vertical: 'bottom', wrap: false, indent: 0 });
  assert.deepEqual(a(8), { horizontal: 'left', vertical: 'top', wrap: false, indent: 0 });
  assert.equal(a(13).indent, 2);
  assert.deepEqual([a(15).wrap, a(15).vertical], [true, 'top']);
});
