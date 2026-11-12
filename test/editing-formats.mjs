// Browser check of formats: rich text drawn, a row's and a column's own format (drawn, given by formatting whole
// lines and the whole sheet), a row grown by its font, and a .csv opened as Excel opens one.
// Run after `npm run build`: node test/editing-formats.mjs [outDir]
import fs from 'node:fs';
import { editingPage } from './editing-kit.mjs';
import { openFixture } from './browser.mjs';

const { page, check, finish, openFixtureFrom } = await editingPage();
await openFixture(page, 'basic.xlsx');
fs.mkdirSync('test/out', { recursive: true });

// Rich text (made here as Excel writes it: richtext.test.js): each run drawn in its own font over the cell's.
{
  const { Workbook } = await import('../src/ooxml/workbook.js');
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  wb.setCell(0, 'J12', 'x');
  const { doc } = wb._part(0);
  const ns = doc.documentElement.namespaceURI;
  const c = [...doc.getElementsByTagNameNS(ns, 'c')].find((e) => e.getAttribute('r') === 'J12');
  while (c.firstChild) c.removeChild(c.firstChild);
  c.setAttribute('t', 'inlineStr');
  const rPr = (inner) => `<rPr>${inner}<sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr>`;
  const { parseXml } = await import('../src/ooxml/xml.js');
  const is = parseXml(`<is xmlns="${ns}"><r><t xml:space="preserve">Plain </t></r><r>${rPr('<b/>')}<t>bold</t></r>`
    + `<r><rPr><i/><sz val="16"/><color rgb="FFFF0000"/><rFont val="Courier New"/><family val="3"/></rPr><t xml:space="preserve"> red</t></r></is>`).documentElement;
  c.appendChild(doc.importNode(is, true));
  wb.dirty.add(wb.sheets[0].path);
  fs.writeFileSync('test/out/rich.xlsx', wb.toBytes());
}
await page.evaluate(() => { document.querySelector('input[type=file]').value = ''; });
await openFixtureFrom(page, 'test/out/rich.xlsx', 'rich.xlsx');
const runs = await page.evaluate(() => {
  const el = document.querySelector('.oxe-row[aria-rowindex="13"] .oxe-cell[aria-colindex="11"] .oxe-runs');
  return el ? [...el.children].map((x) => { const st = getComputedStyle(x); return [x.textContent, st.fontWeight, st.fontStyle, st.fontSize, st.color, st.fontFamily.split(',')[0]]; }) : null;
});
check(runs?.length === 3 && runs[0][1] === '400' && runs[1][0] === 'bold' && runs[1][1] === '700' && runs[2][2] === 'italic' && runs[2][3] === '21.3333px' && runs[2][4] === 'rgb(255, 0, 0)' && /Courier New/.test(runs[2][5]),
  'rich text: each run in its own font (bold; italic, 16pt, red, Courier New) over the cell’s', JSON.stringify(runs));

// A filled column's empty cells drawn filled, as Excel draws a column's own format (linestyles.test.js).
let lineFill;
{
  const { Workbook } = await import('../src/ooxml/workbook.js');
  const { parseXml } = await import('../src/ooxml/xml.js');
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  lineFill = wb.styleWith(0, { fill: { pattern: 'solid', fg: { rgb: 'FFFFFF00' } } });
  const { doc } = wb._part(0);
  const root = doc.documentElement;
  const old = [...root.childNodes].find((n) => n.localName === 'cols');
  if (old) root.removeChild(old);
  const cols = parseXml(`<cols xmlns="${root.namespaceURI}"><col min="12" max="12" width="9.140625" style="${lineFill}"/></cols>`).documentElement;
  root.insertBefore(doc.importNode(cols, true), [...root.childNodes].find((n) => n.localName === 'sheetData'));
  wb.dirty.add(wb.sheets[0].path);
  fs.writeFileSync('test/out/linefill.xlsx', wb.toBytes());
}
await page.evaluate(() => { document.querySelector('input[type=file]').value = ''; });
await openFixtureFrom(page, 'test/out/linefill.xlsx', 'linefill.xlsx');
const filled = await page.evaluate((f) => [...document.querySelectorAll(`.oxe-f${f}`)].filter((e) => !e.classList.contains('oxe-cell')).length, lineFill);
const rowsShown = await page.evaluate(() => document.querySelectorAll('.oxe-row').length);
check(filled >= rowsShown - 2, 'a column’s own fill: its empty cells drawn filled', JSON.stringify({ filled, rowsShown }));

// Column M selected whole, Ctrl+B: the column's own format, as Excel's (a cell typed far below it bold); Ctrl+Z
// takes it away.
await page.evaluate(() => window.oxe.command('selectLines', 'col', 12));
await page.focus('.oxe-grid');
await page.keyboard.press('Control+b');
const colBold = await page.evaluate(() => { const wb = window.oxe.workbook; const sh = wb.sheet(window.oxe.sheetIndex); const col = sh.cols.find((x) => 12 >= x.min && 12 <= x.max); return col ? wb.styles.fonts[wb.styles.cellXfs[col.style].fontId].b : null; });
await page.evaluate(() => window.oxe.enter('M500', 'far'));
const farBold = await page.evaluate(() => { const wb = window.oxe.workbook; const cell = wb.sheet(window.oxe.sheetIndex).cells.get(499).get(12); return wb.styles.fonts[wb.styles.cellXfs[cell.s].fontId].b; });
await page.evaluate(() => { window.oxe.undo(); window.oxe.undo(); });
const colAfterUndo = await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).cols.find((x) => 12 >= x.min && 12 <= x.max)?.style ?? 0);
check(colBold === true && farBold === true && colAfterUndo === 0, 'a whole column made bold: its own format, a cell typed far below bold; Ctrl+Z takes it away', JSON.stringify({ colBold, farBold, colAfterUndo }));

// The whole sheet selected, Ctrl+B: every column its own format (Z9999 bold though empty), as Excel; Ctrl+Z.
await page.evaluate(() => window.oxe.command('selectAll'));
await page.focus('.oxe-grid');
await page.keyboard.press('Control+b');
const sheetBold = await page.evaluate(() => { const wb = window.oxe.workbook; const sh = wb.sheet(window.oxe.sheetIndex); const col = sh.cols.find((x) => 25 >= x.min && 25 <= x.max); return col ? wb.styles.fonts[wb.styles.cellXfs[col.style].fontId].b : null; });
await page.evaluate(() => window.oxe.undo());
const sheetUndone = await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).cols.find((x) => 25 >= x.min && 25 <= x.max)?.style ?? 0);
check(sheetBold === true && sheetUndone === 0, 'the whole sheet made bold: every column its own format; Ctrl+Z takes it away', JSON.stringify({ sheetBold, sheetUndone }));

// Row 30 selected whole, its font made 20pt: the row grows, as Excel's (26.25pt: 35px).
const row30 = () => page.evaluate(() => window.oxe.grid.geometry().rows.size(29));
const row30Before = await row30();
await page.evaluate(() => { window.oxe.command('selectLines', 'row', 29); window.oxe.command('format', { font: { size: 20 } }); });
const row30After = await row30();
await page.evaluate(() => window.oxe.undo());
check(row30Before === 20 && row30After === 35 && await row30() === 20, 'a whole row made 20pt: it grows to 26.25pt as Excel’s; Ctrl+Z back', JSON.stringify({ row30Before, row30After }));


// A .csv (test/csv-excel.csv, its cells as Excel opened them: csv.test.js): one sheet named after the file, a
// percentage, a date, a line break in quotes, a formula.
await page.evaluate(() => { document.querySelector('input[type=file]').value = ''; });
await openFixtureFrom(page, 'test/csv-excel.csv', 'csv-excel.csv');
const csv = await page.evaluate(() => {
  const at = (r, c) => document.querySelector(`.oxe-row[aria-rowindex="${r + 2}"] .oxe-cell[aria-colindex="${c + 2}"]`)?.textContent ?? null;
  return { sheets: window.oxe.workbook.sheets.map((s) => s.name).join(), b6: at(5, 1), b8: at(7, 1), c7: at(6, 2), c4: at(3, 2) };
});
check(csv.sheets === 'csv-excel' && csv.b6 === '50%' && csv.b8 === '3/5/2024' && csv.c7 === '3' && csv.c4 === 'say "hi"',
  'a .csv opened as Excel opens it: one sheet named after the file, 50%, a date, a formula worked out, quotes', JSON.stringify(csv));

// Saved as a CSV again (CSV UTF-8): Excel's bytes for these cells, the byte order mark first.
const resaved = await page.evaluate(() => new TextDecoder('utf-8', { ignoreBOM: true }).decode(window.oxe.saveCsv()));
check(resaved.startsWith('\uFEFFname,num,text\r\nplain,42,"quoted, comma"\r\n') && resaved.includes('pct,50%,FALSE\r\n') && resaved.includes(',"say ""hi"""\r\n'),
  'the CSV saved again as Excel’s CSV UTF-8 writes it', JSON.stringify(resaved.slice(0, 120)));

await finish();
