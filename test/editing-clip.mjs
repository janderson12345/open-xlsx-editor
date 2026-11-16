// Browser check of pasting from other programs and of links to other workbooks, with two editors on one page (two
// workbooks open, as in Excel): a paste event carrying HTML (a web page's table) pastes its formats, merged cells and
// rich text (cliphtml.js); cells copied in one editor and Paste Link in the other make [1]Sheet1!A1 formulas, shown
// as [oxesrc.xlsx]Sheet1!A1, with the values kept until Data > Edit Links' Update Values; Break Link after Excel's
// question; Undo; a name referring to the other workbook made in the New Name box and listed by the Name Manager.
// Run after `npm run build`: node test/editing-clip.mjs [outDir]
import fs from 'node:fs';
import { csvWorkbook } from '../src/ooxml/csv.js';
import { addSheet, renameSheet } from '../src/ooxml/sheetops.js';
import { entryChanges } from '../src/ooxml/entry.js';
import { parseRef } from '../src/ooxml/refs.js';
import { editingPage } from './editing-kit.mjs';

const { page, check, finish, openFixtureFrom, outDir } = await editingPage();

// (The source workbook, made-up numbers, as test/extlinks.test.js makes it; and an empty one to link from.)
const blank = () => csvWorkbook(new Uint8Array(0), 'Sheet1.csv');
const src = blank();
for (const [ref, text] of [['A1', '5'], ['B1', '2.5'], ['A2', 'text'], ['B2', '=A1*2']]) {
  const { r, c } = parseRef(ref);
  src.apply(0, [entryChanges(src, 0, r, c, text).content]);
}
addSheet(src, 1);
renameSheet(src, 1, 'Data Two');
src.apply(1, [{ r: 0, c: 0, value: 7 }]);
src.calculate();
fs.writeFileSync(`${outDir}/oxesrc.xlsx`, src.toBytes());
await openFixtureFrom(page, `${outDir}/oxesrc.xlsx`, 'oxesrc.xlsx');
await page.evaluate((bytes) => {
  const host = document.createElement('div');
  host.id = 'host2';
  host.style.height = '420px';
  document.body.append(host);
  window.oxe2 = window.mountXlsxEditor(host, { bytes: new Uint8Array(bytes), fileName: 'oxedst.xlsx' });
}, [...blank().toBytes()]);
// (A button of the second editor's ribbon, behind its More button when its group doesn't fit.)
const ribbon2 = async (title) => {
  const btn = `#host2 .oxe-toolbar button[title^="${title}"]`;
  if (!(await page.isVisible(btn)) && await page.isVisible('#host2 .oxe-more')) await page.click('#host2 .oxe-more');
  await page.click(btn);
};
const cellOf = (ed, ref) => page.evaluate(([ed, ref]) => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const cell = window[ed].workbook.sheet(window[ed].sheetIndex).cells.get(Number(m[2]) - 1)?.get(c);
  return cell ? { f: cell.f ?? null, v: cell.v ?? null, runs: cell.runs ? cell.runs.map((x) => x.text) : null } : null;
}, [ed, ref]);

// Copy A1:B2 in the first editor (the keys take the cells), Paste Link in the second at D4.
await page.evaluate(() => { window.oxe.command('select', 0, 0); window.oxe.command('extendTo', 1, 1); });
await page.focus('#host .oxe-grid');
await page.keyboard.press('Control+c');
await page.evaluate(() => { window.oxe2.select('D4'); return window.oxe2.command('pasteSpecial', { link: true }); });
const d4 = await cellOf('oxe2', 'D4');
const e5 = await cellOf('oxe2', 'E5');
const text = await page.evaluate(() => { window.oxe2.select('D4'); return [...document.querySelectorAll('#host2 input, #host2 textarea')].map((x) => x.value).find((v) => v.startsWith('=')) || ''; });
check(d4?.f === '[1]Sheet1!A1' && d4.v === 5 && e5?.f === '[1]Sheet1!B2' && e5.v === 10 && text === '=[oxesrc.xlsx]Sheet1!A1',
  'Paste Link from another open workbook: [1]Sheet1!A1 formulas with its values, the formula bar showing [oxesrc.xlsx]Sheet1!A1', JSON.stringify({ d4, e5, text }));

// The source changes: the link keeps its value until Edit Links' Update Values.
await page.evaluate(() => window.oxe.agent().write({ A1: '50' }));
const kept = await cellOf('oxe2', 'D4');
await page.click('#host2 .oxe-rtab >> text=Data');
await ribbon2('Edit Links');
await page.waitForSelector('dialog.oxe-links-box');
const rows = await page.$$eval('dialog.oxe-links-box tbody tr', (xs) => xs.map((x) => [...x.children].map((td) => td.textContent).join('|')));
await page.click('dialog.oxe-links-box button >> text=Update Values');
const updated = await cellOf('oxe2', 'D4');
check(kept?.v === 5 && rows.join() === 'oxesrc.xlsx|Worksheet|A|OK' && updated?.v === 50 && (await cellOf('oxe2', 'E5'))?.v === 100,
  'the value kept until Update Values; Edit Links lists the source, OK while it is open here; updated from it', JSON.stringify({ kept, rows, updated }));

// Break Link: Excel's question, then the formulas are their values; Undo brings the formulas back.
await page.click('dialog.oxe-links-box button >> text=Break Link');
await page.waitForSelector('dialog[role="alertdialog"]');
const question = await page.textContent('dialog[role="alertdialog"]');
await page.click('dialog[role="alertdialog"] button >> text=Break Links');
await page.click('dialog.oxe-links-box button >> text=Close');
const broken = await cellOf('oxe2', 'D4');
await page.evaluate(() => window.oxe2.undo());
const back = await cellOf('oxe2', 'D4');
check(/Are you sure you want to break the links\?/.test(question || '') && broken?.f === null && broken.v === 50 && back?.f === '[1]Sheet1!A1',
  'Break Link after Excel’s question: the formulas become their values; Undo brings them back', JSON.stringify({ broken, back }));

// A name referring to the other workbook: typed with its file's name in the New Name box, listed so by the Name Manager.
await page.click('#host2 .oxe-rtab >> text=Formulas');
await ribbon2('Define Name');
await page.waitForSelector('dialog.oxe-name-box');
const named = await page.isVisible('dialog.oxe-name-box');
if (named) {
  const inputs = await page.$$('dialog.oxe-name-box input[type=text]');
  await inputs[0].fill('Rate');
  await inputs[inputs.length - 1].fill("='[oxesrc.xlsx]Data Two'!$A$1");
  await page.click('dialog.oxe-name-box button >> text=OK');
}
await ribbon2('Name Manager');
await page.waitForSelector('dialog.oxe-name-manager');
const listed = await page.$$eval('dialog.oxe-name-manager tbody tr', (xs) => xs.map((x) => [...x.children].slice(0, 2).map((td) => td.textContent).join(' ')));
await page.keyboard.press('Escape');
const fileText = await page.evaluate(() => window.oxe2.workbook.definedNames.find((d) => d.name === 'Rate')?.text);
check(named && listed.includes("Rate ='[oxesrc.xlsx]Data Two'!$A$1") && fileText === "'[1]Data Two'!$A$1",
  'a name referring to the other workbook: kept as \'[1]Data Two\'!$A$1, listed with the file’s name', JSON.stringify({ named, listed, fileText }));

// A web page's table pasted (a paste event with its HTML): bold, a fill, a merged row, rich text.
const html = '<table><tr><th style="background-color: rgb(221, 235, 247)">City</th><th>Pop.</th></tr><tr><td><b>Avon</b></td><td style="text-align: right">12,400</td></tr>'
  + '<tr><td colspan="2">Both <span style="font-weight: 700">columns</span></td></tr></table>';
await page.evaluate(() => window.oxe2.select('H1'));
await page.focus('#host2 .oxe-grid');
await page.evaluate((h) => {
  const dt = new DataTransfer();
  dt.setData('text/html', h);
  dt.setData('text/plain', 'City\tPop.\nAvon\t12,400\nBoth columns');
  document.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
}, html);
const pasted = await page.evaluate(() => {
  const wb = window.oxe2.workbook;
  const sheet = wb.sheet(window.oxe2.sheetIndex);
  const at = (r, c) => sheet.cells.get(r)?.get(c);
  const font = (cell) => wb.styles.fonts[wb.styles.cellXfs[cell?.s || 0].fontId];
  return { h1: at(0, 7)?.v, bold: font(at(1, 7)).b, fill: wb.styles.fills[wb.styles.cellXfs[at(0, 7).s].fillId].fg?.rgb, i2: at(1, 8)?.v, merges: sheet.merges.map((m) => `${m.r0},${m.c0}:${m.r1},${m.c1}`), runs: at(2, 7)?.runs?.map((x) => x.text) };
});
check(pasted.h1 === 'City' && pasted.bold && pasted.fill === 'FFDDEBF7' && pasted.i2 === 12400 && pasted.merges.includes('2,7:2,8') && pasted.runs?.join('|') === 'Both |columns',
  'HTML pasted from another program: its formats, a merged row and rich text, as Excel pastes it', JSON.stringify(pasted));

await finish();
