// Browser check of sheets grouped on their tabs (Ctrl+click, Shift+click, the tab menu's Select All Sheets and Ungroup
// Sheets) and Home > Fill > Across Worksheets over them (All, Contents, Formats, and Undo); and the Insert Hyperlink
// box's Browse for File and Create New Document. Run after `npm run build`: node test/editing-fillacross.mjs [outDir]
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { editingPage } from './editing-kit.mjs';

const { page, check, finish, openFixtureFrom, ribbon } = await editingPage();

// (Three sheets: the first with a bold number, a formula and text in A1:B2; the others with something else there.)
const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
while (wb.sheets.length > 1) wb.deleteSheet(wb.sheets.length - 1);
wb.renameSheet(0, 'One');
const two = wb.addSheet(1, 'Two');
const three = wb.addSheet(2, 'Three');
for (const ref of ['A1', 'B1', 'A2', 'B2', 'C9']) wb.setCell(0, ref, null);
wb.setCell(0, 'A1', 10);
wb.setStyles(0, [{ r: 0, c: 0, s: wb.styleWith(0, { font: { b: true } }) }]);
wb.setFormula(0, 'B1', 'A1*2');
wb.setCell(0, 'A2', 'text');
for (const k of [two, three]) { wb.setCell(k, 'A1', 'old'); wb.setCell(k, 'B2', 'old'); wb.setCell(k, 'D4', 'outside'); }
fs.mkdirSync('test/out', { recursive: true });
fs.writeFileSync('test/out/fillacross.xlsx', wb.toBytes());
await openFixtureFrom(page, 'test/out/fillacross.xlsx', 'fillacross.xlsx');

const tab = (name) => `.oxe-tab:text-is("${name}")`;
const tabs = () => page.$$eval('.oxe-tab', (xs) => xs.map((x) => `${x.textContent}:${x.getAttribute('aria-selected')}${x.classList.contains('oxe-tab-grouped') ? 'g' : ''}`).join('|'));
const cellOf = (si, r, c) => page.evaluate(([si, r, c]) => { const x = window.oxe.workbook.sheet(si).cells.get(r)?.get(c); return x ? { v: x.v ?? null, f: x.f ?? null, bold: !!window.oxe.workbook.styles.fonts[window.oxe.workbook.styles.cellXfs[x.s || 0].fontId]?.b } : null; }, [si, r, c]);
const fillItem = async () => {
  let disabled = null;
  await ribbon('Fill', async () => { disabled = await page.$eval('.oxe-pop button:has-text("Across Worksheets...")', (b) => b.disabled); });
  await page.keyboard.press('Escape');
  return disabled;
};

// Not grouped: Across Worksheets offered but greyed, as Excel's.
const before = await fillItem();
// Ctrl+click on Two's tab: grouped, One still shown; both tabs marked.
await page.click(tab('Two'), { modifiers: ['Control'] });
const grouped = await tabs();
const shown = await page.evaluate(() => window.oxe.sheetIndex);
const after = await fillItem();
check(before === true && after === false && grouped === 'One:trueg|Two:trueg|Three:false' && shown === 0,
  'Ctrl+click on a tab groups its sheet with the one shown (both marked); Fill > Across Worksheets is offered then, greyed before', JSON.stringify({ before, after, grouped, shown }));

// A1:B2 selected, Fill > Across Worksheets > Contents: Two gets the values and the formula as written, not the bold; Three, not grouped, nothing.
await page.evaluate(() => { window.oxe.select('A1'); window.oxe.command('extendTo', 1, 1); });
await ribbon('Fill', () => page.click('.oxe-pop >> text=Across Worksheets...'));
await page.waitForSelector('dialog.oxe-dialog >> text=Fill Across Worksheets');
const choices = await page.$$eval('dialog.oxe-dialog fieldset label', (ls) => ls.map((l) => l.textContent).join('|'));
await page.click('dialog.oxe-dialog label:text-is("Contents")');
await page.click('dialog.oxe-dialog >> button:text-is("OK")');
const c1 = [await cellOf(1, 0, 0), await cellOf(1, 0, 1), await cellOf(1, 1, 0), await cellOf(1, 1, 1), await cellOf(1, 3, 3), await cellOf(2, 0, 0)];
check(choices === 'All|Contents|Formats' && c1[0].v === 10 && !c1[0].bold && c1[1].f === 'A1*2' && c1[1].v === 20 && c1[2].v === 'text' && c1[3].v == null && c1[4].v === 'outside' && c1[5].v === 'old',
  'Fill Across Worksheets, Contents: the grouped sheet gets the values and formulas (a blank cell clears), not the formats; the cells outside and the other sheet stay', JSON.stringify({ choices, c1 }));

// Undo: Two as it was. Then Formats: only the bold, the values kept.
await page.focus('.oxe-grid');
await page.keyboard.press('Control+z');
const undone = [await cellOf(1, 0, 0), await cellOf(1, 1, 1)];
await page.evaluate(() => window.oxe.command('fillAcross', 'formats'));
const formatted = await cellOf(1, 0, 0);
check(undone[0].v === 'old' && undone[1].v === 'old' && formatted.v === 'old' && formatted.bold,
  'Ctrl+Z takes the fill back; Formats brings the bold alone', JSON.stringify({ undone, formatted }));

// Shift+click on Three: One to Three grouped, and All fills them all; the tab menu's Ungroup Sheets ungroups, Select
// All Sheets groups all; a click on a grouped tab shows it, the group kept; ungrouped, a click shows one sheet alone.
await page.click(tab('Three'), { modifiers: ['Shift'] });
const ranged = await tabs();
await page.evaluate(() => { window.oxe.select('A1'); window.oxe.command('extendTo', 1, 1); window.oxe.command('fillAcross', 'all'); });
const all3 = await cellOf(2, 0, 1);
await page.click(tab('One'), { button: 'right' });
const items = await page.$$eval('.oxe-menu [role="menuitem"]', (xs) => xs.map((x) => x.textContent.trim()));
await page.click('.oxe-menu >> text=Ungroup Sheets');
const ungrouped = await tabs();
await page.click(tab('One'), { button: 'right' });
await page.click('.oxe-menu >> text=Select All Sheets');
const everything = await tabs();
await page.click(tab('Three'));
const stillGrouped = await tabs();
await page.click(tab('One'), { button: 'right' });
await page.click('.oxe-menu >> text=Ungroup Sheets');
await page.click(tab('Two'));
const single = await tabs();
check(ranged === 'One:trueg|Two:trueg|Three:trueg' && all3?.f === 'A1*2' && all3.v === 20 && items.includes('Select All Sheets') && items.includes('Ungroup Sheets')
  && ungrouped === 'One:true|Two:false|Three:false' && everything === 'One:trueg|Two:trueg|Three:trueg' && stillGrouped === 'One:trueg|Two:trueg|Three:trueg' && single === 'One:false|Two:true|Three:false',
  'Shift+click groups the tabs from the one shown; All fills them all; the tab menu’s Ungroup Sheets and Select All Sheets; a grouped tab clicked keeps the group',
  JSON.stringify({ ranged, all3, items, ungrouped, everything, stillGrouped, single }));

// Insert Hyperlink > Browse for File: the file picked, by its name, as a relative link; written as Excel writes it
// (a space as %20).
const relOf = (ref) => page.evaluate((ref) => {
  const w = window.oxe.workbook;
  const si = window.oxe.sheetIndex;
  const { doc } = w._part(si);
  const el = [...doc.getElementsByTagNameNS(doc.documentElement.namespaceURI, 'hyperlink')].find((h) => h.getAttribute('ref') === ref);
  if (!el) return null;
  const id = el.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
  const rel = w.pkg.readRels(w.sheets[si].path).find((x) => x.id === id);
  return { target: rel.rawTarget, external: rel.external, value: w.sheet(si).cells.get(Number(ref.slice(1)) - 1)?.get(ref.charCodeAt(0) - 65)?.v ?? null };
}, ref);
fs.writeFileSync('test/out/My Report.xlsx', wb.toBytes());
await page.evaluate(() => window.oxe.select('F2'));
await page.focus('.oxe-grid');
await page.keyboard.press('Control+k');
await page.waitForSelector('dialog.oxe-dialog >> text=Insert Hyperlink');
const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.click('dialog.oxe-dialog >> button:text-is("Browse for File...")')]);
await chooser.setFiles('test/out/My Report.xlsx');
const addr = await page.inputValue('dialog.oxe-dialog input[id$="-addr"]');
const text = await page.inputValue('dialog.oxe-dialog input[id$="-show"]');
await page.click('dialog.oxe-dialog >> button:text-is("OK")');
const browsed = await relOf('F2');
check(addr === 'My Report.xlsx' && text === 'My Report.xlsx' && browsed?.target === 'My%20Report.xlsx' && browsed.external && browsed.value === 'My Report.xlsx',
  'Browse for File: the file’s name as the address and the text to display; the link written relative, a space as %20, TargetMode External', JSON.stringify({ addr, text, browsed }));

// Create New Document: a name without an extension gets .xlsx; Edit later; the host's onCreateDocument told (the demo
// says so in its status); an empty name refused.
await page.evaluate(() => window.oxe.select('F3'));
await page.focus('.oxe-grid');
await page.keyboard.press('Control+k');
await page.waitForSelector('dialog.oxe-dialog >> text=Insert Hyperlink');
await page.click('dialog.oxe-dialog label:text-is("Create New Document")');
await page.click('dialog.oxe-dialog >> button:text-is("OK")');
const refused = await page.textContent('dialog.oxe-dialog [role="alert"]');
await page.fill('dialog.oxe-dialog input[id$="-docname"]', 'Q3 Budget');
await page.click('dialog.oxe-dialog label:text-is("Edit the new document later")');
await page.click('dialog.oxe-dialog >> button:text-is("OK")');
const created = await relOf('F3');
const told = await page.textContent('#status');
check(refused === 'Type a name for the new document.' && created?.target === 'Q3%20Budget.xlsx' && created.value === 'Q3 Budget.xlsx' && told === 'New document: Q3 Budget.xlsx (edit later)',
  'Create New Document: refused without a name; the link to Q3 Budget.xlsx written as Excel writes it; onCreateDocument told the name and Edit later', JSON.stringify({ refused, created, told }));

// Saved and opened again: the links and the filled cells read back.
const bytes = await page.evaluate(() => Array.from(window.oxe.save()));
fs.writeFileSync('test/out/fillacross-saved.xlsx', Uint8Array.from(bytes));
const back = Workbook.fromBytes(Uint8Array.from(bytes));
const rels = back.pkg.readRels(back.sheets[1].path).map((r) => r.rawTarget).sort().join('|');
check(rels === 'My%20Report.xlsx|Q3%20Budget.xlsx' && back.sheet(2).cells.get(0).get(1).f === 'A1*2',
  'saved: the two file links and the filled formula read back', rels);

await finish();
