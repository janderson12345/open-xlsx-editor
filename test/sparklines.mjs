// Browser check of sparklines (docs/SPARKLINES.md) on test/fixtures/sparklines.xlsx: each drawn in its cell as an SVG
// of the cell's size (lines, columns, win/loss, markers, the axis), redrawn when a value it plots is typed over, in the
// printed pages too, and in the PDF export. Writes a screenshot of the grid to the output folder.
// Run after `npm run build`: node test/sparklines.mjs [outDir]
import { editingPage } from './editing-kit.mjs';
import { openFixture } from './browser.mjs';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';

const { page, check, finish, clickCell, outDir } = await editingPage();
await openFixture(page, 'sparklines.xlsx');
check(await page.evaluate(() => document.compatMode) === 'CSS1Compat', 'standards mode');
// (Column O in view, the rows from the top.)
await page.evaluate(() => { window.oxe.grid.el.scrollLeft = 520; window.oxe.grid.el.scrollTop = 0; });
await page.waitForTimeout(100);

/** The sparklines drawn in the grid: each one's place and size, and its paths' painting. */
const sparks = () => page.evaluate(() => {
  const g = window.oxe.grid.geometry();
  return [...document.querySelectorAll('.oxe-fills svg.oxe-spark')].map((s) => {
    const y = parseFloat(s.style.getPropertyValue('--oxe-y'));
    const row = g.rows.indexAt(y + 1);
    return { row, w: parseFloat(s.style.getPropertyValue('--oxe-w')), h: parseFloat(s.style.getPropertyValue('--oxe-h')), vb: s.getAttribute('viewBox'),
      paths: [...s.querySelectorAll('path')].map((p) => ({ d: p.getAttribute('d'), fill: p.getAttribute('fill'), stroke: p.getAttribute('stroke') })) };
  });
});
const first = await sparks();
const byRow = (list, r) => list.find((s) => s.row === r - 1);
const line = byRow(first, 2);
const column = byRow(first, 6);
const winloss = byRow(first, 10);
const marked = byRow(first, 3);
const axis = byRow(first, 13);
check(first.length >= 12 && line && line.vb === `0 0 ${line.w} ${line.h}`, `sparklines drawn in their cells (${first.length} in view), each the size of its cell`, JSON.stringify(first.map((s) => s.row)));
check(line?.paths.length === 1 && line.paths[0].stroke === '#203764' && line.paths[0].d.split('L').length === 12, 'a line: one stroke through 12 points, in Excel\'s default style\'s colour', JSON.stringify(line?.paths));
check(column?.paths.length === 12 && column.paths.filter((p) => p.fill === '#ED7D31').length === 4, 'a column sparkline: 12 bars, the 4 negative ones in the negative colour', JSON.stringify(column?.paths.map((p) => p.fill)));
check(winloss?.paths.length === 11, 'win/loss: a bar for each value but the 0', String(winloss?.paths.length));
check(marked?.paths.filter((p) => p.fill && p.fill !== 'none').length === 12 && marked.paths.some((p) => p.fill === '#00FF00') && marked.paths.some((p) => p.fill === '#FF00FF'), 'markers on every point, the high and low ones in their colours', JSON.stringify(marked?.paths.map((p) => p.fill)));
check(axis?.paths.some((p) => p.stroke === '#C00000'), 'the horizontal axis, in its colour', JSON.stringify(axis?.paths.map((p) => p.stroke)));
await page.screenshot({ path: `${outDir}/sparklines.png` });

// Typed over: B2 becomes the line's highest point, at the plot's top (3.84px down).
await page.evaluate(() => { window.oxe.grid.el.scrollLeft = 0; });
await page.waitForTimeout(100);
await clickCell(1, 1);
await page.keyboard.type('50');
await page.keyboard.press('Enter');
await page.waitForTimeout(100);
await page.evaluate(() => { window.oxe.grid.el.scrollLeft = 520; });
await page.waitForTimeout(100);
const after = byRow(await sparks(), 2);
check(/^M[\d.]+ 3\.84L/.test(after?.paths[0].d || ''), 'redrawn when a value it plots changes: 50 in B2 puts its first point at the top', after?.paths[0].d.slice(0, 40));
await page.keyboard.press('Control+z');

// Printed: every sparkline on the pages.
const printed = await page.evaluate(() => { window.oxe.preparePrint(); return document.querySelectorAll('.oxe-print svg.oxe-spark').length; });
check(printed === 28, 'in the printed pages: all 28', String(printed));

// The PDF export: the sparklines' fills and strokes (the negative columns' orange among them) on its pages.
const bytes = await page.evaluate(async () => {
  const { exportPdf } = await import('/dist/pdf.js');
  return [...await exportPdf(window.oxe, { localFonts: false })];
});
const doc = await pdfjs.getDocument({ data: new Uint8Array(bytes), disableFontFace: true, verbosity: 0 }).promise;
const colors = new Set();
for (let i = 1; i <= doc.numPages; i++) {
  const ops = await (await doc.getPage(i)).getOperatorList();
  ops.fnArray.forEach((fn, k) => { if (fn === pdfjs.OPS.setFillRGBColor || fn === pdfjs.OPS.setStrokeRGBColor) colors.add(String(ops.argsArray[k][0]).toLowerCase()); });
}
check(['#ed7d31', '#203764', '#00ff00', '#c00000'].every((c) => colors.has(c)), 'in the PDF export: bars, lines, markers and the axis in their colours', [...colors].join(' '));
await finish();
