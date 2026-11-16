// Browser suite: other countries' settings, the Text Import box and the Save As box in the demo page. A German locale
// set: numbers and dates shown and typed in it, formulas with semicolons; a semicolon .txt opened through the Text
// Import box with a date column chosen; Data > From Text/CSV putting a fixed-width file at the active cell, one undo
// step; Save As giving a semicolon CSV of the sheet shown (with the active-sheet note) and an .xlsx.
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { editingPage } from './editing-kit.mjs';

const { page, check, outDir, finish, shown, value, clickCell, ribbon } = await editingPage();
const at = (name, text) => { const p = path.join(outDir, name); fs.writeFileSync(p, text); return p; };
const box = 'dialog.oxe-textimport-box';
await page.addScriptTag({ path: createRequire(import.meta.url).resolve('axe-core/axe.min.js') });
/** axe-core's WCAG 2.1 A and AA rules on the page as it is (a box open). */
const axe = async () => (await page.evaluate(() => window.axe.run(document, { runOnly: { type: 'tag', values: ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa'] } }))).violations.map((v) => `${v.id}: ${v.nodes.slice(0, 2).map((n) => n.target.join(' ')).join(', ')}`);

// (A German locale, as the `locale` option or setLocale gives it.)
const tag = await page.evaluate(() => window.oxe.setLocale('de-DE').tag);
check(tag === 'de-DE' && await page.evaluate(() => window.oxe.locale.listSep) === ';', "setLocale('de-DE'): the page's locale, ; as its list separator");

// A semicolon file through the Text Import box (the demo sends a .txt there).
await page.setInputFiles('input[type=file]', at('locale-de.txt', 'item;amount;when\r\nalpha;1.234,5;05.03.2024\r\nbeta;-2,25;13.12.2023\r\n'));
await page.waitForSelector(box);
const guessed = await page.evaluate((b) => {
  const d = document.querySelector(b);
  const checked = [...d.querySelectorAll('.oxe-textimport-delims input[type=checkbox]')].filter((x) => x.checked).map((x) => x.nextSibling.textContent);
  return { checked, cols: d.querySelectorAll('.oxe-textimport-preview thead select').length, decimal: d.querySelectorAll('.oxe-textimport-top input')[0].value, first: d.querySelector('.oxe-textimport-preview tbody td')?.textContent };
}, box);
const importIssues = await axe();
check(!importIssues.length, 'the Text Import box: no axe violations', importIssues.join(String.fromCharCode(10) + '     '));
check(guessed.checked.join() === 'Semicolon' && guessed.cols === 3 && guessed.decimal === ',' && guessed.first === 'item',
  'Text Import: the semicolon detected, three columns previewed, the locale\'s decimal comma', JSON.stringify(guessed));
await page.selectOption(`${box} .oxe-textimport-preview thead th:nth-child(3) select`, 'DMY');
await page.click(`${box} button:text-is("Import")`);
await page.waitForFunction(() => window.oxe.fileName === 'locale-de.txt' && !document.querySelector('dialog[open]'));
await clickCell(1, 2);
const bar = () => page.evaluate(() => document.querySelector('.oxe-bar [aria-label="Formula bar"]')?.value);
const imported = { b2: await value('B2'), c2: await value('C2'), c3: await value('C3'), b2Shown: await shown(1, 1), c2Bar: await bar(), b3Shown: await shown(2, 1) };
check(imported.b2 === 1234.5 && imported.c2 === 45356 && imported.c3 === 45273 && imported.b2Shown === '1.234,50' && imported.c2Bar === '05.03.2024' && imported.b3Shown === '-2,25',
  'imported: 1.234,5 as 1234.5 shown 1.234,50; the DMY column as dates, dd.mm.yyyy in the formula bar', JSON.stringify(imported));

// Typing in the locale: a decimal comma, a date with dots, a formula with semicolons (kept in the file en-US).
await clickCell(4, 0);
await page.keyboard.type('2,5');
await page.keyboard.press('Enter');
await page.keyboard.type('1.5');
await page.keyboard.press('Enter');
await page.keyboard.type('=ROUND(2,25;1)');
await page.keyboard.press('Enter');
await clickCell(6, 0);
const typed = await page.evaluate(() => {
  const sh = window.oxe.workbook.sheet(window.oxe.sheetIndex);
  return { a5: sh.cells.get(4).get(0).v, a6: sh.cells.get(5).get(0).v, f: sh.cells.get(6).get(0).f, bar: document.querySelector('.oxe-bar [aria-label="Formula bar"]')?.value };
});
check(typed.a5 === 2.5 && typed.a6 === 46143 && typed.f === 'ROUND(2.25,1)' && typed.bar === '=ROUND(2,25;1)' && await shown(6, 0) === '2,3',
  'typed: 2,5 is 2.5; 1.5 the 1st of May; =ROUND(2,25;1) stored ROUND(2.25,1), shown so in the formula bar, its value 2,3', JSON.stringify(typed));

// Back to en-US: the same cells drawn the American way.
await page.evaluate(() => window.oxe.setLocale('en-US'));
check(await shown(1, 1) === '1,234.50' && await shown(1, 2) === '3/5/2024' && await shown(4, 0) === '2.5', "setLocale('en-US'): the sheet drawn again, 1,234.50 and 3/5/2024");

// Data > From Text/CSV: a fixed-width file at the active cell, one step of the undo history.
await clickCell(0, 5);
const chooser = page.waitForEvent('filechooser');
await page.click('.oxe-rtab >> text="Data"');
await ribbon('From Text/CSV');
await (await chooser).setFiles(at('fixed.txt', 'ID    NAME    AMOUNT\r\n00012 Widget    15.25\r\n7     Bolt     -3\r\n'));
await page.waitForSelector(box);
await page.check(`${box} input[type=radio][value=fixed]`);
const fixedShown = await page.evaluate((b) => [...document.querySelectorAll(`${b} .oxe-textimport-preview tbody tr:nth-child(2) td`)].map((x) => x.textContent), box);
await page.selectOption(`${box} .oxe-textimport-preview thead th:nth-child(1) select`, 'text');
await page.click(`${box} button:text-is("Import")`);
await page.waitForFunction(() => !document.querySelector('dialog[open]'));
const placed = { f2: await value('F2'), g2: await value('G2'), h2: await value('H2'), h3: await value('H3') };
check(fixedShown.join('|') === '00012|Widget|15.25' && placed.f2 === '00012' && placed.g2 === 'Widget' && placed.h2 === 15.25 && placed.h3 === -3,
  'From Text/CSV: fixed width with the breaks guessed, the first column as Text (00012 kept), put at F1', JSON.stringify({ fixedShown, placed }));
await page.keyboard.press('Control+z');
check(await value('F2') === null && await value('H3') === null, 'From Text/CSV undone in one step');

// Save As: a semicolon CSV of the sheet shown; then an .xlsx, the workbook going by its new name.
await page.evaluate(() => window.oxe.insertSheet({ after: true }));
await page.evaluate(() => window.oxe.showSheet(0));
await page.focus('.oxe-grid');
await page.keyboard.press('F12');
await page.waitForSelector('dialog.oxe-saveas-box');
await page.fill('dialog.oxe-saveas-box .oxe-saveas-name', 'report');
await page.selectOption('dialog.oxe-saveas-box select >> nth=0', 'csv');
const note = await page.textContent('dialog.oxe-saveas-box .oxe-saveas-note');
const saveIssues = await axe();
check(!saveIssues.length, 'the Save As box: no axe violations', saveIssues.join(String.fromCharCode(10) + '     '));
await page.selectOption('dialog.oxe-saveas-box select >> nth=1', ';');
const dl = page.waitForEvent('download');
await page.click('dialog.oxe-saveas-box button:text-is("Save")');
const csv = await dl;
const csvText = fs.readFileSync(await csv.path(), 'utf8');
check(csv.suggestedFilename() === 'report.csv' && /active sheet, locale-de, is saved/.test(note) && csvText.startsWith(String.fromCharCode(0xfeff) + 'item;amount;when')
  && csvText.includes('\r\nalpha;"1,234.50";3/5/2024'),
  'Save As CSV: report.csv, ; between fields (a comma still quoted, as Excel quotes one in a tab-delimited file), the active-sheet note, CSV UTF-8', JSON.stringify({ name: csv.suggestedFilename(), note, start: csvText.slice(0, 60) }));
const saved = await page.evaluate(async () => {
  const out = await window.oxe.saveAs({ fileName: 'report.xlsx', type: 'xlsx' });
  return { name: out.fileName, pk: out.bytes[0] === 0x50 && out.bytes[1] === 0x4b, now: window.oxe.fileName };
});
check(saved.name === 'report.xlsx' && saved.pk && saved.now === 'report.xlsx', 'Save As .xlsx: a workbook package, the editor now on report.xlsx', JSON.stringify(saved));

await finish();
