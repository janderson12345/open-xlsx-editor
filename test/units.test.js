import { test } from 'node:test';
import assert from 'node:assert/strict';
import { emuToPx, pxToEmu, ptToPx, colWidthToPx, pxToColWidth, charsToColWidth, colWidthToChars, defaultColPx, rowHeightToPx, printColPx, printRowPx } from '../src/ooxml/units.js';

test('EMUs and points', () => {
  assert.equal(emuToPx(914400), 96);
  assert.equal(pxToEmu(1), 9525);
  assert.equal(ptToPx(72), 96);
  assert.equal(rowHeightToPx(15), 20);
});

// Calibri 11's maximum digit width is 7px at 96 dpi.
test('column widths in characters of the maximum digit width, as the specification gives them', () => {
  assert.equal(charsToColWidth(8.43, 7), 9.140625);
  assert.equal(colWidthToPx(9.140625, 7), 64);
  assert.equal(colWidthToChars(9.140625, 7), 8.43);
  for (let px = 12; px < 400; px++) assert.equal(colWidthToPx(pxToColWidth(px, 7), 7), px, `${px}px`);
});

test('the default column: 8 characters plus padding, up to a multiple of 8px', () => {
  assert.equal(defaultColPx({}, 7), 64);
  assert.equal(defaultColPx({ baseColWidth: 10 }, 7), 80);
  assert.equal(defaultColPx({ defaultColWidth: 9.140625 }, 7), 64);
});

// Excel's PDFs of test/fixtures/widths.xlsx and large.xlsx (docs/LAYOUT.md).
test('printed sizes: columns x 1.075, rows in whole 600 dpi device pixels of pt x 1.3565', () => {
  assert.equal(printColPx(64), 68.8);
  assert.equal(printRowPx(15), 127 / 6.25);
  assert.equal(printRowPx(60), 508 / 6.25);
});
