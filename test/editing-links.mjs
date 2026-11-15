// Browser check of hyperlinks: a click on a hyperlink's cell follows it (an address opened, a place in the workbook
// gone to), a click held selects the cell instead, Excel's ScreenTip and hand over it, and the cell menu's Open
// Hyperlink; and names made by Formulas > Create from Selection, Use in Formula and Paste Names, Home > Fill > Justify, Circle Invalid Data, Apply Names, iterative calculation, and Page Layout's Print Area, Breaks, Print Titles, Orientation and Scale to Fit. Run after `npm run build`:
// node test/editing-links.mjs [outDir]
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { addHyperlink } from '../src/ooxml/links.js';
import { editingPage } from './editing-kit.mjs';

const { page, check, finish, openFixtureFrom, sel, clickCell, pointAt, menu, ribbon } = await editingPage();

// (Made here, as Excel's Hyperlinks.Add writes them: links.test.js.)
const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
addHyperlink(wb, 0, 2, 9, { address: 'https://example.com/page' }); // J3
addHyperlink(wb, 0, 3, 9, { location: "'Notes & more'!C30", text: 'Go there' }); // J4
addHyperlink(wb, 0, 4, 9, { address: 'example.com', tooltip: 'Tip here' }); // J5
fs.mkdirSync('test/out', { recursive: true });
fs.writeFileSync('test/out/links.xlsx', wb.toBytes());
await openFixtureFrom(page, 'test/out/links.xlsx', 'links.xlsx');
await page.evaluate(() => { window.opened = []; window.open = (url) => { window.opened.push(url); return null; }; });
const opened = () => page.evaluate(() => window.opened.slice());

// Over J3: the hand and Excel's ScreenTip; a click opens the address.
await page.evaluate(() => window.oxe.select('J1'));
const p = await pointAt(2, 9);
await page.mouse.move(p.x, p.y);
const tip = await page.getAttribute('.oxe-grid', 'title');
const hand = await page.evaluate(() => getComputedStyle(document.querySelector('.oxe-grid')).cursor);
await clickCell(2, 9);
const first = await opened();
check(tip === 'https://example.com/page - Click once to follow. Click and hold to select this cell.' && hand === 'pointer' && first.join() === 'https://example.com/page',
  'a hyperlink: the hand and Excel’s ScreenTip over it; a click opens its address', JSON.stringify({ tip, hand, first }));

// Held, it selects the cell and opens nothing.
await page.mouse.move(p.x, p.y);
await page.mouse.down();
await page.waitForTimeout(700);
await page.mouse.up();
const held = await opened();
check(held.length === 1 && (await sel()).ar === 2, 'a hyperlink clicked and held: its cell selected, nothing opened', JSON.stringify({ held, sel: await sel() }));

// J4, a place in the workbook: gone to.
await clickCell(3, 9);
await page.waitForFunction(() => window.oxe.sheetIndex === 1);
const there = await sel();
check(there.ar === 29 && there.ac === 2, 'a hyperlink to a place in the workbook: that sheet shown, the cell selected', JSON.stringify(there));

// J5, a file's path: its own ScreenTip; refused in Excel's words; the cell menu's Open Hyperlink does the same.
await page.evaluate(() => window.oxe.showSheet(0));
const q = await pointAt(4, 9);
await page.mouse.move(q.x + 1, q.y + 1);
const ownTip = await page.getAttribute('.oxe-grid', 'title');
await clickCell(4, 9);
await page.waitForSelector('dialog.oxe-dialog[role="alertdialog"]');
const refusal = await page.textContent('dialog.oxe-dialog[role="alertdialog"]');
await page.keyboard.press('Enter');
// (From the keyboard: Shift+F10 opens the cell's menu, its first item focused.)
await page.evaluate(() => window.oxe.select('J3'));
await page.focus('.oxe-grid');
await page.keyboard.press('Shift+F10');
await page.waitForSelector('.oxe-menu >> text=Open Hyperlink');
const menuFocused = await page.evaluate(() => !!document.activeElement.closest('.oxe-menu'));
const items = await page.$$eval('.oxe-menu [role="menuitem"]', (xs) => xs.map((x) => x.textContent.trim()).join('|'));
await page.click('.oxe-menu >> text=Open Hyperlink');
check(ownTip.startsWith('Tip here - ') && /Cannot open the specified file\./.test(refusal || '') && /Open Hyperlink\|Remove Hyperlink/.test(items) && menuFocused && (await opened()).length === 2,
  'a file’s path: its own ScreenTip, refused in Excel’s words; Shift+F10, the cell menu’s Open Hyperlink opens J3’s address', JSON.stringify({ ownTip, refusal, items, menuFocused }));
// (The menu key does the same, on another cell.)
await page.evaluate(() => window.oxe.select('L2'));
await page.focus('.oxe-grid');
await page.keyboard.press('ContextMenu');
await page.waitForSelector('.oxe-menu >> text=Paste Special...');
const keyMenu = await page.$$eval('.oxe-menu [role="menuitem"]', (xs) => xs.map((x) => x.textContent.trim()).includes('Link'));
await page.keyboard.press('Escape');
const backToGrid = await page.evaluate(() => document.activeElement.classList.contains('oxe-grid'));
check(keyMenu && backToGrid, 'the menu key: the cell’s menu (with Link); Esc back to the grid', JSON.stringify({ keyMenu, backToGrid }));

// Ctrl+K on an empty cell: Excel's Insert Hyperlink box; a web address and text to display; then Edit Hyperlink... on
// the cell's menu, to a place in the workbook; then its Remove Link.
const linkOf = (r, c) => page.evaluate(([r, c]) => {
  const wbk = window.oxe.workbook;
  const { doc } = wbk._part(window.oxe.sheetIndex);
  const el = [...doc.getElementsByTagNameNS(doc.documentElement.namespaceURI, 'hyperlink')].find((h) => h.getAttribute('ref') === String.fromCharCode(65 + c) + (r + 1));
  if (!el) return null;
  const id = el.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
  return { location: el.getAttribute('location'), address: id ? wbk.pkg.readRels(wbk.sheets[window.oxe.sheetIndex].path).find((x) => x.id === id).rawTarget : null };
}, [r, c]);
await page.keyboard.press('Escape');
await page.evaluate(() => window.oxe.select('J8'));
await page.focus('.oxe-grid');
await page.keyboard.press('Control+k');
await page.waitForSelector('dialog.oxe-dialog >> text=Insert Hyperlink');
const kinds = await page.$$eval('dialog.oxe-dialog fieldset label', (ls) => ls.map((l) => l.textContent).join('|'));
await page.fill('dialog.oxe-dialog input[id$="-show"]', 'New link');
await page.fill('dialog.oxe-dialog input[id$="-addr"]', 'www.example.com/new');
await page.click('dialog.oxe-dialog >> button:text-is("OK")');
const made = await linkOf(7, 9);
const j8 = await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(7)?.get(9)?.v);
const j8Look = await page.evaluate(() => { const el = document.querySelector('.oxe-row[aria-rowindex="9"] .oxe-cell[aria-colindex="11"]'); return el ? getComputedStyle(el).textDecorationLine : null; });
check(kinds === 'Existing File or Web Page|Place in This Document|E-mail Address' && made?.address === 'http://www.example.com/new' && j8 === 'New link' && /underline/.test(j8Look || ''),
  'Ctrl+K: Excel’s Insert Hyperlink box; a web address (http:// before www.) and its text, underlined in the Hyperlink style', JSON.stringify({ kinds, made, j8, j8Look }));
const j8p = await pointAt(7, 9);
await page.mouse.click(j8p.x, j8p.y, { button: 'right' });
await page.click('.oxe-menu >> text=Edit Hyperlink...');
await page.waitForSelector('dialog.oxe-dialog >> text=Edit Hyperlink');
const shownAddr = await page.inputValue('dialog.oxe-dialog input[id$="-addr"]');
await page.click('dialog.oxe-dialog label:text-is("Place in This Document")');
await page.fill('dialog.oxe-dialog input[id$="-ref"]', 'c5');
await page.selectOption('dialog.oxe-dialog select', 's:Notes & more');
await page.click('dialog.oxe-dialog >> button:text-is("OK")');
const edited = await linkOf(7, 9);
await page.keyboard.press('Control+k');
await page.waitForSelector('dialog.oxe-dialog >> text=Edit Hyperlink');
await page.click('dialog.oxe-dialog >> button:text-is("Remove Link")');
const removed = await linkOf(7, 9);
await page.keyboard.press('Control+z');
const undone = await linkOf(7, 9);
check(shownAddr === 'http://www.example.com/new' && edited?.location === "'Notes & more'!C5" && !edited.address && removed === null && undone?.location === "'Notes & more'!C5",
  'Edit Hyperlink: the link shown, changed to a place in the workbook; Remove Link; Ctrl+Z brings it back', JSON.stringify({ shownAddr, edited, removed, undone }));

// A HYPERLINK formula typed into J10: underlined in the Hyperlink style, as Excel gives it; a click opens its
// link_location, worked out in the cell's place.
await page.evaluate(() => { window.oxe.enter('K10', 'fn'); window.oxe.enter('J10', '=HYPERLINK("https://example.com/"&K10,"Function link")'); });
const j10Look = await page.evaluate(() => { const el = document.querySelector('.oxe-row[aria-rowindex="11"] .oxe-cell[aria-colindex="11"]'); return el ? [el.textContent, getComputedStyle(el).textDecorationLine] : null; });
await page.keyboard.press('Escape');
await clickCell(9, 9);
const fnOpened = (await opened()).at(-1);
check(j10Look?.[0] === 'Function link' && /underline/.test(j10Look[1]) && fnOpened === 'https://example.com/fn',
  'a HYPERLINK formula typed: the Hyperlink style; a click opens its link_location', JSON.stringify({ j10Look, fnOpened }));

// Create from Selection (Ctrl+Shift+F3) on a grid with labels on top and left: the box starts with both checked; OK
// names each column, each row and the whole (createnames.test.js has Excel's); a formula uses one; Ctrl+Z takes them
// all back.
await page.evaluate(() => {
  for (const [ref, v] of [['L20', 'Grid'], ['M20', 'Jan'], ['N20', 'Feb'], ['L21', 'North'], ['L22', 'South'], ['M21', 1], ['N21', 2], ['M22', 3], ['N22', 4]]) window.oxe.enter(ref, v);
  window.oxe.enter('P20', '=SUM(Feb)');
  window.oxe.select('L20');
  window.oxe.command('extendTo', 21, 13);
});
await page.focus('.oxe-grid');
await page.keyboard.press('Control+Shift+F3');
await page.waitForSelector('dialog.oxe-dialog >> text=Create Names from Selection');
const boxChecks = await page.evaluate(() => [...document.querySelectorAll('dialog.oxe-dialog input[type="checkbox"]')].map((x) => `${x.nextSibling.textContent}:${x.checked}`));
await page.click('dialog.oxe-dialog >> button:text-is("OK")');
const named = await page.evaluate(() => ['Jan', 'Feb', 'North', 'South', 'Grid'].map((n) => window.oxe.workbook.definedName(n, null)?.text?.replace(/^[^!]*!/, '')));
const sumFeb = await page.evaluate(() => window.oxe.workbook.sheet(0).cells.get(19).get(15).v);
await page.focus('.oxe-grid');
await page.keyboard.press('Control+z');
const gone = await page.evaluate(() => ['Jan', 'Grid'].every((n) => !window.oxe.workbook.definedName(n, null)));
check(boxChecks.join() === 'Top row:true,Left column:true,Bottom row:false,Right column:false'
  && named.join() === '$M$21:$M$22,$N$21:$N$22,$M$21:$N$21,$M$22:$N$22,$M$21:$N$22' && sumFeb === 6 && gone,
'Ctrl+Shift+F3: Top row and Left column checked; OK names Jan, Feb, North, South and Grid as Excel would; =SUM(Feb) 6; Ctrl+Z takes them back',
JSON.stringify({ boxChecks, named, sumFeb, gone }));

// Paste Names (F3): Paste List writes the names and, as text, what each refers to, from the active cell down (as
// Excel's ListNames; createnames.test.js); F3 while a formula is typed puts the chosen name at the caret.
await page.keyboard.press('Control+y');
await page.evaluate(() => window.oxe.select('R30'));
await page.focus('.oxe-grid');
await page.keyboard.press('F3');
await page.waitForSelector('dialog.oxe-dialog >> text=Paste Name');
const offered = await page.evaluate(() => [...document.querySelectorAll('dialog.oxe-dialog option')].map((o) => o.value));
await page.click('dialog.oxe-dialog >> button:text-is("Paste List")');
const listed = await page.evaluate(() => [29, 30].map((r) => [17, 18].map((c) => { const x = window.oxe.workbook.sheet(0).cells.get(r)?.get(c); return `${x?.t}:${x?.v}`; }).join('|')));
await page.evaluate(() => window.oxe.select('T30'));
await page.focus('.oxe-grid');
await page.keyboard.type('=SUM(');
await page.keyboard.press('F3');
await page.waitForSelector('dialog.oxe-dialog >> text=Paste Name');
const listWhileEditing = await page.evaluate(() => [...document.querySelectorAll('dialog.oxe-dialog button')].some((b) => b.textContent === 'Paste List'));
await page.selectOption('dialog.oxe-dialog select', 'Feb');
await page.click('dialog.oxe-dialog >> button:text-is("OK")');
await page.keyboard.type(')');
await page.keyboard.press('Enter');
const used = await page.evaluate(() => { const x = window.oxe.workbook.sheet(0).cells.get(29).get(19); return `${x.f}=${x.v}`; });
// (basic.xlsx's own Amounts first.)
check(offered.join() === 'Amounts,Feb,Grid,Jan,North,South' && listed[0] === 's:Amounts|s:=Data!$B$2:$B$10' && listed[1] === 's:Feb|s:=Data!$N$21:$N$22'
  && !listWhileEditing && used === 'SUM(Feb)=6',
'F3: Paste Name lists the names; Paste List writes them and their references as text; while typing =SUM( F3 and Feb gives =SUM(Feb) 6',
JSON.stringify({ offered, listed, listWhileEditing, used }));

// Fill > Justify: Excel's own lines for 24 paragraphs in Calibri and Arial (test/justify-excel.json, Range.Justify in
// a hidden Excel; Aptos Narrow's are left out, the test browser having no twin of it), each on a sheet of its own at
// its width; then Excel's refusal of a number, and its question when the text goes on below.
const excelLines = JSON.parse(fs.readFileSync('test/justify-excel.json', 'utf8')).filter((x) => x.font !== 'Aptos Narrow');
const justifyCase = async (cells, range, chars, font) => {
  await page.evaluate(({ cells, range, chars, font }) => {
    const o = window.oxe;
    o.command('insertSheet');
    o.select('A1');
    if (chars) o.command('columnWidth', chars);
    if (font) o.command('format', { font: { name: font } });
    cells.forEach((v, r) => { if (v !== '') o.enter(`A${r + 1}`, String(v)); });
    o.select('A1'); o.command('extendTo', range[0], range[1]);
  }, { cells, range, chars, font });
  await page.click('.oxe-rtab >> text=Home');
  await menu('Fill', 'Justify');
};
const columnA = () => page.evaluate(() => { const out = []; const s = window.oxe.workbook.sheet(window.oxe.sheetIndex); for (let r = 0; r < 40; r++) { const x = s.cells.get(r)?.get(0); out.push(x?.v == null ? '' : String(x.v)); } while (out.length && out[out.length - 1] === '') out.pop(); return out; });
// (Excel's breaks follow the font's own widths: counted only where Calibri and Arial are installed, as on Windows.
// Elsewhere, Linux CI among them, a stand-in measures a little differently, so each paragraph is checked for its words
// only.)
const fontsHere = await page.evaluate(() => ['Calibri', 'Arial'].every((f) => {
  const ctx = document.createElement('canvas').getContext('2d');
  const w = (font) => { ctx.font = `40px ${font}`; return ctx.measureText('mmmmmmmmmmlli1WQ').width; };
  return w(`'${f}', monospace`) !== w('monospace') && w(`'${f}', serif`) !== w('serif');
}));
let same = 0;
let kept = 0;
for (const x of excelLines) {
  await justifyCase([x.text], [39, 0], (x.px - 5) / 7, x.font);
  const got = await columnA();
  if (got.join('|') === x.lines.join('|')) same++;
  if (got.join(' ') === x.text && got.length > 1) kept++;
}
if (fontsHere) check(same >= 20, `Fill > Justify: ${same} of ${excelLines.length} Calibri and Arial paragraphs broken into lines exactly as Excel's (fitted: 20 here, docs/FILL.md)`);
else check(kept === excelLines.length, `Fill > Justify: ${kept} of ${excelLines.length} paragraphs broken into lines, every word kept (Calibri and Arial not installed: ${same} broken exactly as Excel's)`);
await justifyCase(['alpha', 42, 'beta'], [2, 0]);
await page.waitForSelector('dialog.oxe-dialog >> text=Cannot justify cells containing numbers or formulas.');
await page.keyboard.press('Enter');
const refusedLeft = await columnA();
await justifyCase(['one two three four five six seven eight nine ten', '', '', 'below'], [1, 0]);
await page.waitForSelector('dialog.oxe-dialog >> text=Text will extend below range.');
await page.click('dialog.oxe-dialog >> button:text-is("OK")');
const extended = await columnA();
check(refusedLeft.join('|') === 'alpha|42|beta' && extended.length > 4 && !extended.includes('below') && extended.join(' ') === 'one two three four five six seven eight nine ten',
  'Fill > Justify: a number refused in Excel\'s words; text going on below asked first, then written over the cells there', JSON.stringify({ refusedLeft, extended }));

// Data > Data Validation > Circle Invalid Data: whole numbers 1 to 10 over B2:B5 (B5 blank, ignored): ovals round B3's 50
// and B4's text, where Excel draws them (dvcircles.js); B3 put right, its oval goes; kept with the sheet; Clear
// Validation Circles takes them away.
const circlesNow = () => page.evaluate(() => [...document.querySelectorAll('.oxe-circles ellipse')].map((e) => ['cx', 'cy', 'rx', 'ry'].map((k) => Math.round(Number(e.getAttribute(k)) * 10) / 10).join(',')));
await page.evaluate(() => {
  const o = window.oxe;
  o.command('insertSheet');
  o.enter('B2', '5'); o.enter('B3', '50'); o.enter('B4', 'abc');
  o.command('sheetChange', (wb, si) => wb.setValidation(si, [{ r0: 1, c0: 1, r1: 4, c1: 1 }], { type: 'whole', op: 'between', f1: '1', f2: '10', ignoreBlank: true }), 'always');
});
const circleSheet = await page.evaluate(() => window.oxe.sheetIndex);
await page.click('.oxe-rtab >> text=Data');
await page.click('.oxe-toolbar button.oxe-split-more[aria-label^="Data Validation"]');
await page.click('.oxe-pop >> text=Circle Invalid Data');
const circled = await circlesNow();
// (B3's oval in Excel's place: 3.75pt left of the cell and 1.5pt above, 7.5pt wider and 3pt taller.)
const want = await page.evaluate(() => {
  const { rows, cols, rowHeadW } = window.oxe.grid.geometry();
  const w = cols.size(1) + 10; const h = rows.size(2) + 4;
  return [rowHeadW + cols.start(1) - 5 + w / 2, rows.start(2) - 2 + h / 2, w / 2, h / 2].map((v) => Math.round(v * 10) / 10).join(',');
});
await page.evaluate(() => window.oxe.enter('B3', '7'));
const afterFix = await circlesNow();
await page.evaluate((i) => { window.oxe.showSheet(0); window.oxe.showSheet(i); }, circleSheet);
const back = await circlesNow();
await page.click('.oxe-toolbar button.oxe-split-more[aria-label^="Data Validation"]');
await page.click('.oxe-pop >> text=Clear Validation Circles');
const cleared = await circlesNow();
check(circled.length === 2 && circled.includes(want) && afterFix.length === 1 && back.join() === afterFix.join() && !cleared.length,
  'Circle Invalid Data: B3 and B4 circled, B3 put right loses its oval, kept with its sheet, Clear Validation Circles takes them', JSON.stringify({ circled, want, afterFix, back, cleared }));

// Formulas > Define Name > Apply Names...: on a new sheet with Qty (=$B$2), Price (=$C$2) and Pair (=$B$2:$C$2) defined,
// OK from one cell applies them over the whole sheet as Excel's ApplyNames: H8's =B2*C2 Qty*Price, H9's
// =SUM($B$2:$C$2)&"B2" SUM(Pair)&"B2" (the text left), H10's =Sheet!B2 as it was; Ctrl+Z takes them back.
await page.evaluate(() => {
  const o = window.oxe;
  o.command('insertSheet');
  const name = o.workbook.sheets[o.sheetIndex].name;
  o.enter('B2', '4'); o.enter('C2', '5'); o.enter('H8', '=B2*C2'); o.enter('H9', '=SUM($B$2:$C$2)&"B2"'); o.enter('H10', `='${name}'!B2`);
});
// (Named from the Name Box, as typing a name there over a selection does.)
await page.evaluate(() => {
  const o = window.oxe;
  for (const [n, r0, c0, c1] of [['Qty', 1, 1, 1], ['Price', 1, 2, 2], ['Pair', 1, 1, 2]]) { o.select(`${String.fromCharCode(65 + c0)}${r0 + 1}`); if (c1 > c0) o.command('extendTo', r0, c1); o.command('nameBox', n); }
  o.select('A1');
});
const formulasH = () => page.evaluate(() => [7, 8, 9].map((r) => window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(r)?.get(7)?.f));
const beforeApply = await formulasH();
await page.click('.oxe-rtab >> text=Formulas');
await ribbon('Define Name options', () => page.click('.oxe-pop >> text=Apply Names...'));
await page.waitForSelector('dialog.oxe-apply-names');
const offeredNames = await page.evaluate(() => [...document.querySelectorAll('dialog.oxe-apply-names option')].map((o) => `${o.value}${o.selected ? '*' : ''}`));
await page.click('dialog.oxe-apply-names >> button:text-is("OK")');
const applied = await formulasH();
await page.focus('.oxe-grid');
await page.keyboard.press('Control+z');
const unapplied = await formulasH();
check(offeredNames.join() === 'Pair*,Price*,Qty*' && applied.join('|') === `Qty*Price|SUM(Pair)&"B2"|${beforeApply[2]}` && unapplied.join('|') === beforeApply.join('|'),
  'Apply Names: the names offered, all chosen; OK applies them over the sheet as Excel\'s ApplyNames (text and sheet-named references left); Ctrl+Z takes it back',
  JSON.stringify({ offeredNames, beforeApply, applied, unapplied }));

// Formulas > Calculation Options > Iterative Calculation...: on a new sheet A1 =A1+1 is circular (left out, 0); the box
// turned on with Maximum Iterations 50 gives A1 50 (each pass one more, as Excel goes round), and calcPr has
// iterate="1" iterateCount="50" as Excel writes it.
await page.evaluate(() => { const o = window.oxe; o.command('insertSheet'); o.enter('A1', '=A1+1'); o.select('B2'); });
await page.click('.oxe-rtab >> text=Formulas');
await ribbon('Calculation Options', () => page.click('.oxe-pop >> text=Iterative Calculation...'));
await page.waitForSelector('dialog.oxe-iteration-box');
const dimmed = await page.isDisabled('dialog.oxe-iteration-box input.oxe-dialog-field >> nth=0');
await page.check('dialog.oxe-iteration-box input[type=checkbox]');
await page.fill('dialog.oxe-iteration-box input.oxe-dialog-field >> nth=0', '50');
await page.click('dialog.oxe-iteration-box >> button:text-is("OK")');
const iterated = await page.evaluate(() => {
  const o = window.oxe;
  const pr = o.workbook.doc.getElementsByTagNameNS('*', 'calcPr')[0];
  return { a1: o.workbook.sheet(o.sheetIndex).cells.get(0).get(0).v, iterate: pr?.getAttribute('iterate'), count: pr?.getAttribute('iterateCount') };
});
check(dimmed && iterated.a1 === 50 && iterated.iterate === '1' && iterated.count === '50',
  'Iterative Calculation: the box (its numbers dimmed while off); on with 50 iterations, A1 =A1+1 goes round to 50; calcPr as Excel writes it',
  JSON.stringify({ dimmed, ...iterated }));

// Page Layout > Breaks and Print Area: on a new sheet with cells to A40, Insert Page Break at C5 gives a row and a
// column break (the print layout one page more each way), Ctrl+Z takes it back; Set Print Area on B2:D5, and then the
// menu offers Add to Print Area, as Excel's does.
await page.evaluate(() => { const o = window.oxe; o.command('insertSheet'); for (let r = 1; r <= 40; r++) o.enter(`A${r}`, String(r)); o.enter('F1', 'x'); o.select('C5'); });
const printed = () => page.evaluate(() => window.oxe.preparePrint());
const pagesBefore = await printed();
await page.click('.oxe-rtab >> text=Page Layout');
await ribbon('Breaks', () => page.click('.oxe-pop >> text=Insert Page Break'));
const brks = await page.evaluate(() => JSON.stringify(window.oxe.workbook.sheet(window.oxe.sheetIndex).page.breaks));
const pagesWith = await printed();
await page.focus('.oxe-grid');
await page.keyboard.press('Control+z');
const brksUndone = await page.evaluate(() => JSON.stringify(window.oxe.workbook.sheet(window.oxe.sheetIndex).page.breaks));
await page.evaluate(() => { window.oxe.select('B2'); window.oxe.command('extendTo', 4, 3); });
const addBefore = await (async () => { let seen = false; await ribbon('Print Area', async () => { seen = await page.isVisible('.oxe-pop >> text=Add to Print Area'); await page.click('.oxe-pop >> text=Set Print Area'); }); return seen; })();
const areaText = await page.evaluate(() => window.oxe.workbook.definedNames.find((d) => d.name === '_xlnm.Print_Area' && d.sheet === window.oxe.sheetIndex)?.text);
const addAfter = await (async () => { let seen = false; await ribbon('Print Area', async () => { seen = await page.isVisible('.oxe-pop >> text=Add to Print Area'); await page.keyboard.press('Escape'); }); return seen; })();
check(brks === '{"rows":[4],"cols":[2]}' && pagesWith > pagesBefore && brksUndone === '{"rows":[],"cols":[]}' && !addBefore && /!\$B\$2:\$D\$5$/.test(areaText || '') && addAfter,
  'Page Layout: Insert Page Break at C5 (a row and a column break, more pages), Ctrl+Z; Set Print Area, then Add to Print Area offered',
  JSON.stringify({ brks, pagesBefore, pagesWith, brksUndone, addBefore, areaText, addAfter }));

// Page Layout > Print Titles: Excel's Page Setup box at its Sheet tab, the print area set above in it; rows "1:2"
// typed, OK: _xlnm.Print_Titles as Excel saves it; a lone "$3" refused with a note; Ctrl+Z takes the titles back.
await page.click('.oxe-rtab >> text=Page Layout');
await ribbon('Print Titles');
await page.waitForSelector('dialog.oxe-page-setup');
const shownArea = await page.inputValue('dialog.oxe-page-setup input:visible >> nth=0');
await page.fill('dialog.oxe-page-setup input:visible >> nth=1', '$3');
await page.click('dialog.oxe-page-setup >> button:text-is("OK")');
const refusedNote = await page.textContent('dialog.oxe-page-setup .oxe-dialog-note');
await page.fill('dialog.oxe-page-setup input:visible >> nth=1', '1:2');
await page.click('dialog.oxe-page-setup >> button:text-is("OK")');
const titlesText = await page.evaluate(() => window.oxe.workbook.definedNames.find((d) => d.name === '_xlnm.Print_Titles' && d.sheet === window.oxe.sheetIndex)?.text);
await page.focus('.oxe-grid');
await page.keyboard.press('Control+z');
const titlesUndone = await page.evaluate(() => window.oxe.workbook.definedNames.some((d) => d.name === '_xlnm.Print_Titles' && d.sheet === window.oxe.sheetIndex));
check(shownArea === '$B$2:$D$5' && /rows/i.test(refusedNote || '') && /!\$1:\$2$/.test(titlesText || '') && !titlesUndone,
  'Print Titles: the Page Setup box shows the print area; a lone $3 refused; 1:2 saved as $1:$2; Ctrl+Z takes it back',
  JSON.stringify({ shownArea, refusedNote, titlesText, titlesUndone }));

// Page Layout > Orientation and Scale to Fit: Landscape from the menu (the sheet's page landscape, its pages fewer
// down); Width 1 page (fit to pages, height Automatic); Ctrl+Z twice back to portrait and unfitted.
await page.evaluate(() => { const o = window.oxe; o.command('insertSheet'); for (let r = 1; r <= 120; r++) o.enter(`A${r}`, String(r)); o.enter('M1', 'x'); o.select('A1'); });
const pageNow = () => page.evaluate(() => { const p = window.oxe.workbook.sheet(window.oxe.sheetIndex).page; return { landscape: p.landscape, fit: p.fitToPage, w: p.fitToWidth, h: p.fitToHeight }; });
const portraitPages = await printed();
await page.click('.oxe-rtab >> text=Page Layout');
await ribbon('Orientation', () => page.click('.oxe-pop >> text=Landscape'));
const afterLandscape = await pageNow();
const landscapePages = await printed();
await ribbon('Width', () => page.click('.oxe-pop >> text=1 page'));
const afterWidth = await pageNow();
await page.focus('.oxe-grid');
await page.keyboard.press('Control+z');
await page.keyboard.press('Control+z');
const undoneLayout = await pageNow();
check(afterLandscape.landscape && landscapePages !== portraitPages && afterWidth.fit && afterWidth.w === 1 && afterWidth.h === 0 && !undoneLayout.landscape && !undoneLayout.fit,
  'Page Layout: Orientation > Landscape (a different page count), Scale to Fit Width 1 page (height Automatic), Ctrl+Z back',
  JSON.stringify({ portraitPages, afterLandscape, landscapePages, afterWidth, undoneLayout }));

// Margins > Custom Margins...: the Page Setup box at its Margins tab; Left 0.6 and Center horizontally, OK: pageMargins
// left="0.6" and printOptions horizontalCentered="1"; Ctrl+Z takes them back.
await ribbon('Margins', () => page.click('.oxe-pop >> text=Custom Margins...'));
await page.waitForSelector('dialog.oxe-page-setup');
const onMargins = await page.getAttribute('dialog.oxe-page-setup [role=tab] >> nth=0', 'aria-selected');
await page.fill('dialog.oxe-page-setup input:visible >> nth=2', '0.6');
await page.check('dialog.oxe-page-setup input[type=checkbox] >> nth=0');
await page.click('dialog.oxe-page-setup >> button:text-is("OK")');
const marginsNow = await page.evaluate(() => { const p = window.oxe.workbook.sheet(window.oxe.sheetIndex).page; return { left: p.margins.left, h: p.centerH }; });
await page.focus('.oxe-grid');
await page.keyboard.press('Control+z');
const marginsUndone = await page.evaluate(() => { const p = window.oxe.workbook.sheet(window.oxe.sheetIndex).page; return { left: p.margins.left, h: p.centerH }; });
check(onMargins === 'true' && marginsNow.left === 0.6 && marginsNow.h && marginsUndone.left === 0.7 && !marginsUndone.h,
  'Custom Margins...: the Page Setup box at Margins; Left 0.6 and Center horizontally set; Ctrl+Z back',
  JSON.stringify({ onMargins, marginsNow, marginsUndone }));

// Page Layout > Sheet Options > Gridlines: Print: the printed page gets a hairline at each row and column edge of its
// cells (A1:C8 here: 9 across, 4 down), none with it off.
await page.evaluate(() => { const o = window.oxe; o.command('insertSheet'); for (let r = 1; r <= 8; r++) for (const c of ['A', 'B', 'C']) o.enter(`${c}${r}`, String(r)); });
const gridCount = () => page.evaluate(() => { window.oxe.preparePrint(); return [document.querySelectorAll('.oxe-page .oxe-print-grid-h').length, document.querySelectorAll('.oxe-page .oxe-print-grid-v').length]; });
const gridOff = await gridCount();
await page.evaluate(() => window.oxe.command('printGridlines', true));
const gridOn = await gridCount();
check(gridOff.join() === '0,0' && gridOn.join() === '9,4', 'Gridlines: Print draws a hairline at each row and column edge of the printed cells', JSON.stringify({ gridOff, gridOn }));
// Headings: Print: the letters A to C above, the numbers 1 to 8 beside, in a column for three digits.
await page.evaluate(() => window.oxe.command('printHeadings', true));
const heads = await page.evaluate(() => { window.oxe.preparePrint(); const labels = [...document.querySelectorAll('.oxe-page .oxe-print-heads .oxe-cell')].map((e) => e.textContent); return labels.join(','); });
check(heads === 'A,B,C,1,2,3,4,5,6,7,8', 'Headings: Print puts the column letters and row numbers on the printed page', heads);

// A print area of two ranges, Set Print Area on A1:C3 then Add to Print Area on A30:C32: two printed pages, the second
// starting at A30, as Excel prints them.
await page.evaluate(() => { const o = window.oxe; o.command('insertSheet'); for (let r = 1; r <= 40; r++) o.enter(`A${r}`, `a${r}`); o.select('A1'); o.command('extendTo', 2, 2); o.command('printArea', 'set'); o.select('A30'); o.command('extendTo', 31, 2); o.command('printArea', 'add'); });
const twoAreas = await page.evaluate(() => { const n = window.oxe.preparePrint(); const firsts = [...document.querySelectorAll('.oxe-page')].map((p) => p.querySelector('.oxe-cell')?.textContent); return { n, firsts }; });
check(twoAreas.n === 2 && twoAreas.firsts.join() === 'a1,a30', 'A print area of two ranges prints each on a page of its own', JSON.stringify(twoAreas));

await finish();
