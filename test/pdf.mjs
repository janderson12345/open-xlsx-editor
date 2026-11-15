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

const all = await exported({ sheets: 'all' });
const allDoc = await PDFDocument.load(new Uint8Array(all));
const sheetsShown = await page.evaluate(() => window.oxe.workbook.sheets.filter((s, i) => window.oxe.workbook.canShow(i)).length);
t.check(allDoc.getPageCount() > doc.getPageCount() && sheetsShown === 2, 'the whole workbook: the visible sheets’ pages, the hidden one left out', `${allDoc.getPageCount()} pages`);
t.check(await page.evaluate(() => window.oxe.sheetIndex) === 0 && !(await page.evaluate(() => document.querySelector('.oxe-root').classList.contains('oxe-exporting'))), 'afterwards the same sheet shown, the print pages hidden again');

t.check(!errors.length, 'no page errors', errors.join('\n     '));
await close();
console.log(t.failed ? `${t.failed} failed` : 'all passed');
process.exit(t.failed ? 1 : 0);
