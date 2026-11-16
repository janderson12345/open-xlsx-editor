// PDF export (src/pdf.js): the sheet shown, and the whole workbook, exported from the demo page; the PDF read back
// here: its pages as the print layout's, at the sheet's paper size, its text real text with each cell's content.
// Run after `npm run build`: node test/pdf.mjs
import { PDFDocument } from 'pdf-lib';
import { openDemo, openFixture, checker } from './browser.mjs';

import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
const t = checker();
const { page, errors, close } = await openDemo();
await openFixture(page, 'basic.xlsx');

const exported = (options) => page.evaluate(async (o) => {
  const { exportPdf } = await import('/dist/pdf.js');
  const bytes = await exportPdf(window.oxe, { localFonts: false, ...o });
  return [...bytes];
}, options);
const texts = async (bytes) => {
  const doc = await pdfjs.getDocument({ data: new Uint8Array(bytes), disableFontFace: true, verbosity: 0 }).promise;
  const out = [];
  for (let i = 1; i <= doc.numPages; i++) out.push((await (await doc.getPage(i)).getTextContent()).items.map((x) => x.str).join(' '));
  return out;
};

const printed = await page.evaluate(() => window.oxe.preparePrint());
const one = await exported({ title: 'Basic' });
const doc = await PDFDocument.load(new Uint8Array(one));
const size = doc.getPage(0).getSize();
t.check(doc.getPageCount() === printed && doc.getTitle() === 'Basic', `the sheet shown: ${printed} pages, as the print layout makes`, `${doc.getPageCount()} / ${printed}`);
t.check(Math.round(size.width) === 612 && Math.round(size.height) === 792, 'Letter paper, in points (612 by 792)', JSON.stringify(size));
const words = (await texts(one)).join(' ');
const shown = await page.evaluate(() => window.oxe.grid.shownText(0, 0));
t.check(words.includes(shown) && /\d/.test(words), `real text: A1's "${shown}" and the numbers are in it`, words.slice(0, 200));

// (basic.xlsx's E12, "Rich bold" with its second word a bold run: printed as its runs, and in the PDF each run is its
// own text in its own font.)
const printedRuns = await page.evaluate(async () => {
  await window.oxe.preparePrint();
  return [...document.querySelectorAll('.oxe-print .oxe-runs > span')].map((x) => [x.textContent, getComputedStyle(x).fontWeight]);
});
t.check(printedRuns.some(([text, w]) => text === 'bold' && Number(w) >= 600), 'print: a rich cell drawn as its runs, the bold run bold', JSON.stringify(printedRuns));
{
  const pdfDoc = await pdfjs.getDocument({ data: new Uint8Array(one), disableFontFace: true, verbosity: 0 }).promise;
  const items = (await (await pdfDoc.getPage(1)).getTextContent()).items;
  const rich = items.find((x) => x.str.trim() === 'Rich');
  const bold = items.find((x) => x.str.trim() === 'bold');
  t.check(rich && bold && rich.fontName !== bold.fontName && Math.abs(rich.transform[5] - bold.transform[5]) < 0.5 && bold.transform[4] > rich.transform[4],
    'PDF: the rich cell’s runs on one line, the bold run after the plain one in a font of its own', JSON.stringify([rich, bold]));
}

const all = await exported({ sheets: 'all' });
const allDoc = await PDFDocument.load(new Uint8Array(all));
const sheetsShown = await page.evaluate(() => window.oxe.workbook.sheets.filter((s, i) => window.oxe.workbook.canShow(i)).length);
t.check(allDoc.getPageCount() > doc.getPageCount() && sheetsShown === 2, 'the whole workbook: the visible sheets’ pages, the hidden one left out', `${allDoc.getPageCount()} pages`);
t.check(await page.evaluate(() => window.oxe.sheetIndex) === 0 && !(await page.evaluate(() => document.querySelector('.oxe-root').classList.contains('oxe-exporting'))), 'afterwards the same sheet shown, the print pages hidden again');

t.check(!errors.length, 'no page errors', errors.join('\n     '));
await close();
console.log(t.failed ? `${t.failed} failed` : 'all passed');
process.exit(t.failed ? 1 : 0);
