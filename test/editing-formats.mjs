// Browser check of formats: rich text drawn, a row's and a column's own format (drawn, given by formatting whole
// lines and the whole sheet), a row grown by its font, a .csv opened as Excel opens one, and notes, validation and
// conditional formats pasted with their cells.
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

// (On a new sheet of the CSV's workbook.)
// A note copied with its cell: Ctrl+V brings it, as Excel's Paste does; Ctrl+Z takes it away; Paste Special's
// Comments and Notes brings it alone.
await page.evaluate(() => { window.oxe.command('insertSheet'); window.oxe.select('B2'); window.oxe.enter('B2', 'noted'); });
await page.focus('.oxe-grid');
await page.keyboard.press('Shift+F2');
await page.waitForSelector('dialog.oxe-note-edit');
await page.keyboard.type('copied note');
await page.keyboard.press('Escape');
const marks = () => page.evaluate(() => document.querySelectorAll('.oxe-note-mark').length);
const oneMark = await marks();
await page.evaluate(() => window.oxe.select('B2'));
await page.focus('.oxe-grid');
await page.keyboard.press('Control+c');
await page.evaluate(() => window.oxe.select('D4'));
await page.keyboard.press('Control+v');
const pastedMarks = await marks();
const d4 = await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(3)?.get(3)?.v);
await page.keyboard.press('Control+z');
const undoneMarks = await marks();
await page.evaluate(() => window.oxe.select('F4'));
await page.keyboard.press('Control+Alt+v');
await page.waitForSelector('dialog.oxe-dialog >> text=Comments and Notes');
await page.click('dialog.oxe-dialog >> label:text-is("Comments and Notes")');
await page.click('dialog.oxe-dialog >> button:text-is("OK")');
await page.waitForFunction(() => document.querySelectorAll('.oxe-note-mark').length === 2);
const f4 = await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(3)?.get(5)?.v ?? null);
await page.keyboard.press('Escape');
check(oneMark === 1 && pastedMarks === 2 && d4 === 'noted' && undoneMarks === 1 && f4 === null,
  'a note copied: Ctrl+V brings it with the cell, Ctrl+Z takes it away; Paste Special > Comments and Notes brings it alone', JSON.stringify({ oneMark, pastedMarks, d4, undoneMarks, f4 }));

// A cell's validation copied with it: Ctrl+V adds the pasted cell to the rule, as Excel's Paste does; Paste
// Special's Validation is in the box.
await page.evaluate(() => {
  window.oxe.command('sheetChange', (wb, si) => wb.setValidation(si, [{ r0: 7, c0: 1, r1: 7, c1: 1 }], { type: 'whole', op: 'between', f1: '1', f2: '9' }));
  window.oxe.select('B8');
});
await page.focus('.oxe-grid');
await page.keyboard.press('Control+c');
await page.evaluate(() => window.oxe.select('D8'));
await page.keyboard.press('Control+v');
const ruleRanges = await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).validations.map((v) => v.ranges.length).join());
await page.keyboard.press('Control+Alt+v');
await page.waitForSelector('dialog.oxe-dialog >> text=Comments and Notes');
const hasValidation = await page.isVisible('dialog.oxe-dialog >> label:text-is("Validation")');
await page.keyboard.press('Escape');
await page.keyboard.press('Escape');
check(ruleRanges === '2' && hasValidation, 'a cell with validation copied: Ctrl+V adds the pasted cell to its rule; Paste Special has Validation', JSON.stringify({ ruleRanges, hasValidation }));
// F4's note cut and pasted at H4 (Ctrl+X, Ctrl+V): the note moves with the cell, as Excel's does; Ctrl+Z puts it back.
const noteRefs = () => page.evaluate(() => {
  const wb = window.oxe.workbook;
  const rel = wb.pkg.readRels(wb.sheets[window.oxe.sheetIndex].path).find((r) => /\/comments$/.test(r.type));
  return rel ? [...wb.pkg.text(rel.target).matchAll(/<comment ref="([A-Z]+\d+)"/g)].map((m) => m[1]).sort().join() : '';
});
const refsBefore = await noteRefs();
await page.evaluate(() => window.oxe.select('F4'));
await page.focus('.oxe-grid');
await page.keyboard.press('Control+x');
await page.evaluate(() => window.oxe.select('H4'));
await page.keyboard.press('Control+v');
const refsMoved = await noteRefs();
await page.keyboard.press('Control+z');
check(refsBefore === 'B2,F4' && refsMoved === 'B2,H4' && await noteRefs() === 'B2,F4', 'a note cut and pasted moves with its cell; Ctrl+Z puts it back', JSON.stringify({ refsBefore, refsMoved }));

// Conditional formats copied with their cells (clipcf.js): on a new sheet, "cell value > 5" over A1:A3 (dark red
// text); Paste Special lists Excel's twelve choices; Formats at C1 brings the rule (C1's 8 drawn by it), Values at E1
// none; Ctrl+Z takes the Formats paste back.
await page.evaluate(() => {
  const o = window.oxe;
  o.command('insertSheet');
  for (const ref of ['A1', 'A2', 'A3', 'C1', 'C2', 'C3']) o.enter(ref, ref.startsWith('A') ? '7' : '8');
  o.command('sheetChange', (wb, si) => { wb.setConditional(si, [{ r0: 0, c0: 0, r1: 2, c1: 0 }], { kind: 'cellIs', operator: 'greaterThan', f1: '5', preset: 'lightRed' }); return true; }, 'always');
  o.select('A1'); o.command('extendTo', 2, 0);
});
const cfRanges = () => page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).condFormats.map((x) => x.ranges.map((a) => `${String.fromCharCode(65 + a.c0)}${a.r0 + 1}:${String.fromCharCode(65 + a.c1)}${a.r1 + 1}`).join(' ')).sort().join(','));
const pasteAs = async (ref, label) => {
  await page.evaluate((r) => window.oxe.select(r), ref);
  await page.focus('.oxe-grid');
  await page.keyboard.press('Control+Alt+v');
  await page.waitForSelector('dialog.oxe-dialog >> text=All merging conditional formats');
  const choices = await page.evaluate(() => [...document.querySelectorAll('dialog.oxe-dialog fieldset')][0].querySelectorAll('label').length);
  await page.click(`dialog.oxe-dialog >> label:text-is("${label}")`);
  await page.click('dialog.oxe-dialog >> button:text-is("OK")');
  return choices;
};
await page.focus('.oxe-grid');
await page.keyboard.press('Control+c');
const choices = await pasteAs('C1', 'Formats');
const afterFormats = await cfRanges();
const drawn = await page.evaluate(() => document.querySelectorAll('.oxe-grid .oxe-cf-color').length);
await pasteAs('E1', 'Values');
const afterValues = await cfRanges();
await page.keyboard.press('Control+z');
await page.keyboard.press('Control+z');
const undoneCf = await cfRanges();
await page.keyboard.press('Escape');
check(choices === 12 && afterFormats === 'A1:A3,C1:C3' && drawn >= 6 && afterValues === afterFormats && undoneCf === 'A1:A3',
  'conditional formats: Paste Special\'s twelve choices; Formats brings the rule (the pasted cells drawn by it), Values none; Ctrl+Z takes it back',
  JSON.stringify({ choices, afterFormats, drawn, afterValues, undoneCf }));

await finish();
