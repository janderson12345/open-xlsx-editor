// Browser check of editing, continued on a fresh basic.xlsx: Sort, AutoSum, the Name Box, the agent, Filter, data
// validation, sizes, Find and Replace, View, and entries as Excel reads them.
// Run after `npm run build`: node test/editing-data.mjs [outDir]
import { openFixture } from './browser.mjs';
import { editingPage } from './editing-kit.mjs';

const { page, check, finish, shown, sel, value, ribbon, menu } = await editingPage();
const bar = () => page.evaluate(() => document.querySelector('.oxe-formula').value);

await openFixture(page, 'basic.xlsx');

// Sort: E2:E10 (Yes and No) selected, Sort & Filter > Sort Z to A: Excel's Sort Warning (data beside), Sort with the
// selection expanded: the Yes rows first, each row kept whole, the total under them left; Ctrl+Z. A1:F10, Custom
// Sort...: the headers guessed, Choice Z to A then Amount largest first; OK; Ctrl+Z.
const rowsOf = () => page.evaluate(() => Array.from({ length: 9 }, (_, k) => { const l = window.oxe.workbook.sheet(0).cells.get(k + 1); return `${l?.get(0)?.v}|${l?.get(1)?.v}|${l?.get(4)?.v}`; }));
const rowsWere = await rowsOf();
const totalLabel = await value('A11');
await page.evaluate(() => window.oxe.select('E2'));
await page.focus('.oxe-grid');
for (let k = 0; k < 8; k++) await page.keyboard.press('Shift+ArrowDown');
let warning = '';
await ribbon('Sort & Filter', async () => {
  await page.click('.oxe-pop >> text=Sort Z to A');
  await page.waitForSelector('dialog.oxe-dialog');
  warning = await page.textContent('dialog.oxe-dialog .oxe-dialog-text');
  await page.click('dialog.oxe-dialog >> button:text-is("Sort")');
});
let rowsNow = await rowsOf();
const yesFirst = rowsNow.findIndex((x) => x.endsWith('|No')) > rowsNow.findLastIndex((x) => x.endsWith('|Yes'));
check(/^There is data next to your selection/.test(warning) && yesFirst && [...rowsNow].sort().join() === [...rowsWere].sort().join() && await value('A11') === totalLabel,
  'Sort Z to A on E2:E10: Excel’s Sort Warning, the selection expanded, Yes rows first, rows whole, the total left', rowsNow.join(' '));
await page.keyboard.press('Control+z');
check((await rowsOf()).join() === rowsWere.join(), 'Ctrl+Z: the rows back in their order');
await page.evaluate(() => window.oxe.select('A1'));
await page.focus('.oxe-grid');
for (let k = 0; k < 5; k++) await page.keyboard.press('Shift+ArrowRight');
for (let k = 0; k < 9; k++) await page.keyboard.press('Shift+ArrowDown');
let headed = false;
await ribbon('Sort & Filter', async () => {
  await page.click('.oxe-pop >> text=Custom Sort...');
  await page.waitForSelector('dialog.oxe-sort-box');
  headed = await page.isChecked('dialog.oxe-sort-box >> input[type=checkbox]');
  await page.selectOption('.oxe-sort-level >> nth=0 >> select >> nth=0', { label: 'Choice' });
  await page.selectOption('.oxe-sort-level >> nth=0 >> select >> nth=1', { label: 'Z to A' });
  await page.click('dialog.oxe-sort-box >> button:text-is("Add Level")');
  await page.selectOption('.oxe-sort-level >> nth=1 >> select >> nth=0', { label: 'Amount' });
  await page.selectOption('.oxe-sort-level >> nth=1 >> select >> nth=1', { label: 'Largest to Smallest' });
  await page.click('dialog.oxe-sort-box >> button:text-is("OK")');
});
rowsNow = await rowsOf();
const byBoth = rowsNow.every((x, k) => k === 0 || (() => { const [, a, e] = rowsNow[k - 1].split('|'); const [, b, f] = x.split('|'); return e > f || (e === f && Number(a) >= Number(b)); })());
check(headed && byBoth && await value('A1') === 'Item', 'Custom Sort... on A1:F10: headers guessed, Choice Z to A then Amount largest first, the header row left', rowsNow.join(' '));
await page.keyboard.press('Control+z');
check((await rowsOf()).join() === rowsWere.join(), 'Ctrl+Z: as before');

// AutoSum: 5 and 7 typed in Q40:Q41; on Q42 Alt+= opens the cell with Excel's =SUM(Q40:Q41), Enter enters it.
// Q40:Q42 selected, AutoSum > Average: =AVERAGE(Q40:Q41) in the blank Q42. Ctrl+Z each.
await page.evaluate(() => window.oxe.select('Q40'));
await page.focus('.oxe-grid');
for (const n of ['5', '7']) { await page.keyboard.type(n); await page.keyboard.press('Enter'); }
await page.keyboard.press('Alt+Equal');
await page.waitForSelector('.oxe-editor:visible');
const proposed = await page.evaluate(() => document.querySelector('.oxe-editor').value);
await page.keyboard.press('Enter');
const q42 = () => page.evaluate(() => { const c = window.oxe.workbook.sheet(0).cells.get(41)?.get(16); return c && (c.f != null || c.v != null) ? `${c.f}=${c.v}` : ''; });
check(proposed === '=SUM(Q40:Q41)' && await q42() === 'SUM(Q40:Q41)=12', 'Alt+= under 5 and 7: Excel’s =SUM(Q40:Q41) proposed, Enter: 12', `${proposed} ${await q42()}`);
await page.keyboard.press('Control+z');
await page.evaluate(() => window.oxe.select('Q40'));
await page.focus('.oxe-grid');
await page.keyboard.press('Shift+ArrowDown');
await page.keyboard.press('Shift+ArrowDown');
await menu('AutoSum (Alt+=) options', 'Average');
check(await q42() === 'AVERAGE(Q40:Q41)=6', 'Q40:Q42, AutoSum > Average: =AVERAGE(Q40:Q41) in the blank Q42', await q42());
for (let k = 0; k < 3; k++) await page.keyboard.press('Control+z');
check(await q42() === '' && await value('Q40') === null, 'Ctrl+Z: Q40:Q42 empty again', `${await q42()} ${await value('Q40')} ${await value('Q41')}`);

// The Name Box: C3:D4 typed and Enter: that range selected, C3 active, back on the sheet; "1abc": Excel's words.
await page.click('.oxe-namebox');
await page.keyboard.type('c3:d4');
await page.keyboard.press('Enter');
const went = await page.evaluate(() => ({ sel: window.oxe.selection, box: document.querySelector('.oxe-namebox').value, grid: document.activeElement === window.oxe.grid.el }));
check(went.sel.r0 === 2 && went.sel.c0 === 2 && went.sel.r1 === 3 && went.sel.c1 === 3 && went.sel.ar === 2 && went.box === 'C3' && went.grid, 'Name Box, c3:d4 and Enter: C3:D4 selected, C3 active, back on the sheet', JSON.stringify(went));
await page.click('.oxe-namebox');
await page.keyboard.type('1abc');
await page.keyboard.press('Enter');
await page.waitForSelector('dialog.oxe-dialog');
check(/^You must enter a valid reference/.test(await page.textContent('dialog.oxe-dialog')), 'Name Box, 1abc: Excel’s words', await page.textContent('dialog.oxe-dialog'));
await page.keyboard.press('Enter');
await page.click('.oxe-namebox');
await page.keyboard.press('Escape');
check(await page.evaluate(() => document.querySelector('.oxe-namebox').value) === 'C3', 'Esc: the Name Box shows the active cell again');
// A new name typed in the Name Box names the selection (B2:B4), as Excel's does; a formula uses it; Ctrl+F3 lists it in
// the Name Manager; Edit... renames it and the formula follows; Ctrl+Z takes the rename back.
await page.evaluate(() => { window.oxe.select('B2'); window.oxe.command('extendTo', 3, 1); });
await page.click('.oxe-namebox');
await page.keyboard.type('Amounts2');
await page.keyboard.press('Enter');
await page.evaluate(() => window.oxe.enter('J30', '=SUM(Amounts2)'));
const named = await page.evaluate(() => ({ text: window.oxe.workbook.definedName('Amounts2', null)?.text, v: window.oxe.workbook.sheet(0).cells.get(29).get(9).v, b: [1, 2, 3].map((r) => window.oxe.workbook.sheet(0).cells.get(r).get(1).v) }));
check(named.text === 'Data!$B$2:$B$4' && named.v === named.b.reduce((x, y) => x + y, 0), 'Name Box, Amounts2 on B2:B4: the selection named, =SUM(Amounts2) worked out', JSON.stringify(named));
await page.focus('.oxe-grid');
await page.keyboard.press('Control+F3');
await page.waitForSelector('dialog.oxe-name-manager');
const namesListed = await page.evaluate(() => [...document.querySelectorAll('dialog.oxe-name-manager tbody tr')].map((r) => r.textContent));
await page.click('dialog.oxe-name-manager tbody tr:has-text("Amounts2")');
await page.click('dialog.oxe-name-manager >> button:text-is("Edit...")');
await page.waitForSelector('dialog.oxe-name-box');
await page.fill('dialog.oxe-name-box input[id$="-name"]', 'Totals');
await page.click('dialog.oxe-name-box >> button:text-is("OK")');
const relisted = await page.evaluate(() => [...document.querySelectorAll('dialog.oxe-name-manager tbody tr')].map((r) => r.textContent.split('=')[0]));
await page.click('dialog.oxe-name-manager >> button:text-is("Close")');
const renamed = await page.evaluate(() => window.oxe.workbook.sheet(0).cells.get(29).get(9).f);
check(namesListed.some((x) => x.startsWith('Amounts2=Data!$B$2:$B$4Workbook')) && relisted.includes('Totals') && renamed === 'SUM(Totals)', 'Ctrl+F3: the Name Manager lists it; Edit... renames it Totals and =SUM(Totals) follows', JSON.stringify({ namesListed, relisted, renamed }));
await page.focus('.oxe-grid');
await page.keyboard.press('Control+z');
check(await page.evaluate(() => window.oxe.workbook.sheet(0).cells.get(29).get(9).f === 'SUM(Amounts2)' && !!window.oxe.workbook.definedName('Amounts2', null)), 'Ctrl+Z: the rename taken back, the formula with it');
await page.focus('.oxe-grid');
await page.keyboard.press('F5');
await page.waitForSelector('dialog.oxe-dialog >> text=Go To');
await page.fill('dialog.oxe-dialog input.oxe-dialog-field', 'd4');
await page.keyboard.press('Enter');
check(await page.evaluate(() => window.oxe.selection.ar === 3 && window.oxe.selection.ac === 3), 'F5: Excel’s Go To box; d4 and Enter: D4 selected');

// The agent on the open workbook: its entry drawn at once, one step that Ctrl+Z takes back; a sheet it adds gets a tab.
const agentCell = async () => (await shown(8, 7)) ?? "";
const agentWrote = await page.evaluate(() => window.oxe.agent().write({ H9: '50%' }));
check(agentWrote.written.length === 1 && await agentCell() === '50%', 'api.agent().write: H9 entered as typed, drawn as 50%', JSON.stringify(agentWrote));
await page.focus('.oxe-grid');
await page.keyboard.press('Control+z');
check(await agentCell() === '', 'Ctrl+Z: the agent’s entry taken back');
const agentSheet = await page.evaluate(() => { const name = window.oxe.agent().addSheet('From agent'); return { name, tabs: [...document.querySelectorAll('.oxe-tab')].map((t) => t.textContent) }; });
check(agentSheet.name === 'From agent' && agentSheet.tabs.includes('From agent'), 'api.agent().addSheet: a tab for it', JSON.stringify(agentSheet));
await page.evaluate(() => window.oxe.agent().deleteSheet('From agent'));
// Its review: what one agent changed since it was made, and a cell put back as it was, drawn.
const agentReview = await page.evaluate(() => {
  const a = window.oxe.agent();
  a.write({ H9: 'kiwi' });
  const before = a.changes().cells.filter((c) => c.kind === 'edited').map((c) => `${c.ref}:${c.before.value}>${c.after.value}`);
  const back = a.revert(['H9']);
  return { before, back, after: a.changes().cells.length };
});
check(agentReview.before.length === 1 && /H9:null>kiwi/.test(agentReview.before[0]) && agentReview.back.reverted.length === 1 && agentReview.after === 0 && await agentCell() === '',
  'api.agent().changes: H9 entered; revert puts it back, drawn empty', JSON.stringify(agentReview));

// Filter: B3, Ctrl+Shift+L: buttons on the heading row. Choice's button, "No" unticked, OK: the No rows hidden. E1,
// Alt+Down: the same list; Clear Filter From "Choice": all shown. Ctrl+Shift+L: off, no buttons.
await page.evaluate(() => window.oxe.select('B3'));
await page.focus('.oxe-grid');
await page.keyboard.press('Control+Shift+KeyL');
const head = await page.evaluate(() => window.oxe.workbook.sheet(0).autoFilter?.r0);
check(head === 0 && await page.locator('.oxe-filters .oxe-filter-btn').count() > 0, 'Ctrl+Shift+L on B3: the filter on, buttons on the heading row');
const noRows = await page.evaluate(() => [...window.oxe.workbook.sheet(0).cells].filter(([r, l]) => r > 0 && r < 10 && l.get(4)?.v === 'No').map(([r]) => r));
await page.click('.oxe-grid .oxe-filters .oxe-filter-btn[data-col="4"]');
await page.waitForSelector('dialog.oxe-filter-box');
const listed = await page.evaluate(() => [...document.querySelectorAll('dialog.oxe-filter-box .oxe-filter-list label')].map((l) => l.textContent));
await page.click('dialog.oxe-filter-box >> label:text-is("No")');
await page.click('dialog.oxe-filter-box >> button:text-is("OK")');
const hiddenNow = () => page.evaluate(() => [...window.oxe.workbook.sheet(0).rows].filter(([, x]) => x.hidden).map(([r]) => r).sort((a, b) => a - b));
check(listed.includes('(Select All)') && listed.includes('No') && listed.includes('Yes') && (await hiddenNow()).join() === noRows.join(), 'Choice’s button, No unticked, OK: the No rows hidden', `${listed.join('|')} / ${await hiddenNow()} vs ${noRows}`);
await page.evaluate(() => window.oxe.select('E1'));
await page.focus('.oxe-grid');
await page.keyboard.press('Alt+ArrowDown');
await page.waitForSelector('dialog.oxe-filter-box');
await page.click('dialog.oxe-filter-box >> button:text-is("Clear Filter From \\"Choice\\"")');
check((await hiddenNow()).length === 0, 'E1, Alt+Down, Clear Filter From "Choice": every row shown');
await page.keyboard.press('Control+Shift+KeyL');
check(await page.evaluate(() => window.oxe.workbook.sheet(0).autoFilter) === null && await page.locator('.oxe-filters .oxe-filter-btn').count() === 0, 'Ctrl+Shift+L: the filter off, its buttons gone');

// Data validation: basic.xlsx's E2:E10 is a list (Yes, No). E3 selected: its button; picked No: entered. "Maybe"
// typed into E4: Excel's refusal, Retry or Cancel; Cancel: E4 as it was; Retry: back in the cell's editor.
await page.evaluate(() => window.oxe.select('E3'));
await page.focus('.oxe-grid');
check(await page.isVisible('.oxe-dv-btn'), 'E3, in a list rule: its dropdown button');
await page.click('.oxe-dv-btn');
await page.waitForSelector('dialog.oxe-pick-box');
const offered = await page.evaluate(() => [...document.querySelectorAll('dialog.oxe-pick-box option')].map((o) => o.textContent).join(','));
await page.selectOption('dialog.oxe-pick-box select', { label: 'No' });
await page.keyboard.press('Enter');
check(offered === 'Yes,No' && await value('E3') === 'No', 'its list offers Yes, No; No picked: entered', `${offered} ${await value('E3')}`);
await page.keyboard.press('Control+z');
const e4 = await value('E4');
await page.evaluate(() => window.oxe.select('E4'));
await page.focus('.oxe-grid');
await page.keyboard.type('Maybe');
await page.keyboard.press('Enter');
await page.waitForSelector('dialog.oxe-dialog');
const refusal = await page.textContent('dialog.oxe-dialog');
await page.click('dialog.oxe-dialog >> button:text-is("Cancel")');
check(/doesn't match the data validation restrictions/.test(refusal) && /Retry/.test(refusal) && await value('E4') === e4, 'Maybe typed into E4: Excel’s refusal; Cancel: E4 as it was', refusal);
await page.keyboard.type('Maybe');
await page.keyboard.press('Enter');
await page.waitForSelector('dialog.oxe-dialog');
await page.click('dialog.oxe-dialog >> button:text-is("Retry")');
const again = await page.evaluate(() => document.querySelector('.oxe-editor')?.value);
await page.keyboard.press('Escape');
check(again === 'Maybe' && await value('E4') === e4, 'Retry: back in the editor with Maybe', String(again));

// Sizes: column B's edge dragged 40px wider, stored as Excel stores it; column A's edge double-clicked (AutoFit);
// Ctrl+Z; Home, Format, Row Height... 30 for row 3; Ctrl+9 hides it, Ctrl+Shift+9 shows it again.
const colPx = (c) => page.evaluate((c) => window.oxe.grid.geometry().cols.size(c), c);
const colOf = (c) => page.evaluate((c) => window.oxe.workbook.sheet(0).cols.find((x) => x.min <= c && x.max >= c) || null, c);
await page.evaluate(() => { window.oxe.select('A1'); window.oxe.grid.el.scrollTop = 0; window.oxe.grid.el.scrollLeft = 0; });
const edgeX = (c) => page.evaluate((c) => { const g = window.oxe.grid; const v = g.el.getBoundingClientRect(); const geo = g.geometry(); return { x: v.left + geo.rowHeadW + geo.cols.start(c + 1) - 1, y: v.top + geo.colHeadH / 2 }; }, c);
const bWas = await colPx(1);
let edge = await edgeX(1);
await page.mouse.move(edge.x, edge.y);
await page.mouse.down();
await page.mouse.move(edge.x + 20, edge.y, { steps: 3 });
const tip = await page.textContent('.oxe-guide-tip');
await page.mouse.move(edge.x + 40, edge.y, { steps: 3 });
await page.mouse.up();
const bNow = await colPx(1);
check(bNow === bWas + 40 && (await colOf(1)).width === Math.trunc((bNow / 7) * 256) / 256 && /^Width: [\d.]+ \(\d+ pixels\)$/.test(tip), 'column B dragged 40px wider: its width stored in Excel’s units, Excel’s tip while dragging', `${bWas} -> ${bNow} ${tip}`);
await page.keyboard.press('Control+z');
check(await colPx(1) === bWas, 'Ctrl+Z: column B as it was');
edge = await edgeX(0);
await page.mouse.dblclick(edge.x, edge.y);
const colA = await colOf(0);
check(colA && colA.width > 0 && await colPx(0) > 30, 'column A’s edge double-clicked: AutoFit to its text', JSON.stringify(colA));
await page.keyboard.press('Control+z');
await page.evaluate(() => window.oxe.select('A3'));
await page.focus('.oxe-grid');
await ribbon('Format', async () => { await page.click('.oxe-pop >> text=Row Height...'); await page.fill('.oxe-dialog-field', '30'); await page.keyboard.press('Enter'); });
await page.waitForFunction(() => window.oxe.workbook.sheet(0).rows.get(2)?.ht === 30);
check(await page.evaluate(() => window.oxe.grid.geometry().rows.size(2)) === 40, 'Format, Row Height... 30: row 3 is 30pt (40 pixels)');
await page.focus('.oxe-grid');
await page.keyboard.press('Control+9');
check(await page.evaluate(() => window.oxe.workbook.sheet(0).rows.get(2)?.hidden === true && window.oxe.grid.geometry().rows.size(2) === 0), 'Ctrl+9: row 3 hidden');
await page.keyboard.press('Control+Shift+9');
check(await page.evaluate(() => window.oxe.grid.geometry().rows.size(2)) === 40, 'Ctrl+Shift+9: shown again, still 30pt');

// Find and Replace: Ctrl+F, "Item", Enter goes from A1 to the next match, then the next; Find All counts them. Ctrl+H,
// "Item" with "Thing", Replace All; Esc back to the sheet, Ctrl+Z puts them back.
await page.evaluate(() => window.oxe.select('A1'));
await page.focus('.oxe-grid');
await page.keyboard.press('Control+f');
await page.waitForSelector('.oxe-findbar:not([hidden])');
await page.keyboard.type('Item');
await page.keyboard.press('Enter');
const firstHit = await sel();
await page.keyboard.press('Enter');
const secondHit = await sel();
check(firstHit.ar === 1 && firstHit.ac === 0 && secondHit.ar === 2 && await page.evaluate(() => document.activeElement.classList.contains('oxe-findbox')), 'Ctrl+F, Item, Enter twice: A2 then A3, the box keeping focus', JSON.stringify([firstHit, secondHit]));
await page.click('.oxe-findbar button[title="Find All"]');
const found = await page.textContent('.oxe-findinfo');
check(/^\d+ cell\(s\) found$/.test(found), 'Find All: the cells found counted', found);
await page.keyboard.press('Control+h');
await page.fill('.oxe-replrow .oxe-findbox', 'Thing');
await page.click('.oxe-findbar button[title="Replace All"]');
const done = await page.textContent('.oxe-findinfo');
check(/^All done\. We made \d+ replacements?\./.test(done) && /^Thing/.test(String(await value('A3'))), 'Ctrl+H, Replace All: Item made Thing, said as Excel says it', done);
await page.keyboard.press('Escape');
check(await page.isHidden('.oxe-findbar') && await page.evaluate(() => document.activeElement.classList.contains('oxe-grid')), 'Esc: the bar closed, back in the sheet');
await page.keyboard.press('Control+z');
check(/^Item/.test(String(await value('A3'))), 'Ctrl+Z: every replacement undone at once');

// View: Freeze Panes at C3 from the View tab; scrolled far down and right, a click where A1 shows still selects A1.
// Ctrl+wheel zooms by 10%. Unfreeze Panes, and zoom back to 100%.
await page.evaluate(() => { window.oxe.select('C3'); window.oxe.grid.el.scrollTop = 0; window.oxe.grid.el.scrollLeft = 0; });
await page.click('.oxe-rtab >> text="View"');
await menu('Freeze Panes', 'Freeze Panes');
check(JSON.stringify(await page.evaluate(() => window.oxe.grid.view().freeze)) === '{"rows":2,"cols":2}' && await page.isVisible('.oxe-pane-top') && await page.isVisible('.oxe-pane-corner'), 'View, Freeze Panes at C3: rows 1-2 and columns A-B frozen');
await page.evaluate(() => { window.oxe.grid.el.scrollTop = 2000; window.oxe.grid.el.scrollLeft = 600; });
await page.waitForTimeout(100);
const a1 = await page.evaluate(() => { const g = window.oxe.grid; const v = g.el.getBoundingClientRect(); const geo = g.geometry(); return { x: v.left + geo.rowHeadW + 10, y: v.top + geo.colHeadH + 6 }; });
await page.mouse.click(a1.x, a1.y);
check((await sel()).ar === 0 && (await sel()).ac === 0, 'scrolled away, a click on the frozen corner selects A1', JSON.stringify(await sel()));
await page.mouse.move(a1.x + 200, a1.y + 200);
await page.keyboard.down('Control');
await page.mouse.wheel(0, -100);
await page.keyboard.up('Control');
await page.waitForFunction(() => window.oxe.grid.view().zoom === 110);
check(await page.evaluate(() => window.oxe.workbook.sheet(0).zoom) === 110, 'Ctrl+wheel: zoom 110%, kept for the file');
await menu('Freeze Panes', 'Unfreeze Panes');
await ribbon('100%');
check(JSON.stringify(await page.evaluate(() => window.oxe.grid.view().freeze)) === '{"rows":0,"cols":0}' && await page.evaluate(() => window.oxe.grid.view().zoom) === 100 && await page.isHidden('.oxe-pane-top'), 'Unfreeze Panes and 100%: as before');
await page.click('.oxe-rtab >> text=Home');

// Entries as Excel reads them: a date typed into G40 is a date, shown and in the formula bar as one; F2 and Enter keep
// it; $1,234.50 is currency; Ctrl+Z takes the value and its format away together.
const codeAt = (ref) => page.evaluate((ref) => { const wb = window.oxe.workbook; const cell = wb.sheet(0).cells.get(Number(ref.slice(1)) - 1)?.get(ref.charCodeAt(0) - 65); const xf = wb.styles.cellXfs[cell?.s || 0]; return wb.styles.numFmts[xf.numFmtId] ?? ({ 14: 'm/d/yyyy', 0: 'General', 8: 'currency' })[xf.numFmtId] ?? String(xf.numFmtId); }, ref);
await page.evaluate(() => window.oxe.select('G40'));
await page.focus('.oxe-grid');
await page.keyboard.type('1/2/2024');
await page.keyboard.press('Enter');
check(await value('G40') === 45293 && await shown(39, 6) === '1/2/2024' && await codeAt('G40') === 'm/d/yyyy', 'typed 1/2/2024: the date 45293, shown 1/2/2024', `${await value('G40')} ${await shown(39, 6)}`);
await page.evaluate(() => window.oxe.select('G40'));
check(await bar() === '1/2/2024', 'the formula bar shows it as a date', await bar());
await page.focus('.oxe-grid');
await page.keyboard.press('F2');
await page.keyboard.press('Enter');
check(await value('G40') === 45293, 'F2 and Enter: still the same date');
await page.evaluate(() => window.oxe.select('G41'));
await page.focus('.oxe-grid');
await page.keyboard.type('$1,234.50');
await page.keyboard.press('Enter');
check(await value('G41') === 1234.5 && /\$"?#,##0\.00/.test(await codeAt('G41')), 'typed $1,234.50: 1234.5 as currency (shown ######## in a column this narrow, as in Excel)', await codeAt('G41'));
await page.keyboard.press('Control+z');
check(await value('G41') === null && await codeAt('G41') === 'General', 'Ctrl+Z: the value and its format gone together');

// Formulas > Insert Function: Excel's box, Most Recently Used first (SUM at the top); Statistical, T.TEST picked: its
// syntax under the list; OK: the cell opened with =T.TEST( and the screen tip, array1 in bold, array2 after a comma.
// Shift+F3 in a cell being edited (=1+): the box over it, SUM chosen, typed at the caret; the tip follows. Math & Trig
// > ROUND: =ROUND( in the active cell.
const editorText = () => page.evaluate(() => document.querySelector('.oxe-editor')?.value ?? null);
const tipBold = () => page.evaluate(() => { const t = document.querySelector('.oxe-fntip'); return t ? `${t.textContent}|${t.querySelector('b')?.textContent}` : null; });
await page.evaluate(() => window.oxe.select('J45'));
await page.click('.oxe-rtab >> text=Formulas');
await ribbon('Insert Function (Shift+F3)');
await page.waitForSelector('dialog.oxe-function-box');
const firstListed = await page.evaluate(() => [...document.querySelectorAll('dialog.oxe-function-box .oxe-function-list option')].slice(0, 3).map((o) => o.value).join());
await page.selectOption('dialog.oxe-function-box select.oxe-dialog-select', 'Statistical');
await page.selectOption('dialog.oxe-function-box .oxe-function-list', 'T.TEST');
const syntax = await page.textContent('dialog.oxe-function-box .oxe-function-syntax');
await page.click('dialog.oxe-function-box >> button:text-is("OK")');
await page.waitForFunction(() => document.querySelector('.oxe-editor')?.value === '=T.TEST(');
const tip1 = await tipBold();
await page.keyboard.type('B2:B4,');
const tip2 = await tipBold();
check(firstListed === 'SUM,AVERAGE,IF' && syntax === 'T.TEST(array1, array2, tails, type)' && tip1 === 'T.TEST(array1, array2, tails, type)|array1' && tip2.endsWith('|array2'),
  'Insert Function: Most Recently Used first; Statistical > T.TEST, its syntax; OK: =T.TEST( and the tip, the argument typed in bold', JSON.stringify({ firstListed, syntax, tip1, tip2 }));
await page.keyboard.press('Escape');
await page.focus('.oxe-grid');
await page.keyboard.type('=1+');
await page.keyboard.press('Shift+F3');
await page.waitForSelector('dialog.oxe-function-box');
const recentNow = await page.evaluate(() => document.querySelector('dialog.oxe-function-box .oxe-function-list option')?.value);
await page.fill('dialog.oxe-function-box input.oxe-dialog-field', 'sum');
await page.keyboard.press('Enter');
const searched = await page.evaluate(() => [...document.querySelectorAll('dialog.oxe-function-box .oxe-function-list option')].slice(0, 2).map((o) => o.value).join());
await page.selectOption('dialog.oxe-function-box .oxe-function-list', 'SUM');
await page.click('dialog.oxe-function-box >> button:text-is("OK")');
await page.waitForFunction(() => !document.querySelector('dialog.oxe-function-box'));
const typedIn = await editorText();
const tip3 = await tipBold();
check(recentNow === 'T.TEST' && searched === 'SUM,SUMIF' && typedIn === '=1+SUM(' && tip3 === 'SUM(number1, [number2], ...)|number1',
  'Shift+F3 while editing =1+: T.TEST now most recent; search "sum"; SUM typed at the caret, the cell still being edited', JSON.stringify({ recentNow, searched, typedIn, tip3 }));
await page.keyboard.type('2,3)');
await page.keyboard.press('Enter');
await page.evaluate(() => window.oxe.select('J46'));
await ribbon('Math & Trig', () => page.click('.oxe-pop button:has(span:text-is("ROUND"))'));
await page.waitForFunction(() => document.querySelector('.oxe-editor')?.value === '=ROUND(');
await page.keyboard.type('2.567,1)');
await page.keyboard.press('Enter');
check(await value('J45') === 6 && await value('J46') === 2.6, 'entered: =1+SUM(2,3) is 6; Math & Trig > ROUND: =ROUND(2.567,1) is 2.6', JSON.stringify([await value('J45'), await value('J46')]));

// Formulas > Calculation Options > Manual, as Excel's: a change waits (=J45*2 stays 12, the status bar says Calculate),
// a formula entered is worked out; F9 works out the rest. Back to Automatic.
const statusText = () => page.textContent('.oxe-status');
await page.evaluate(() => window.oxe.select('J47'));
await page.keyboard.type('=J45*2\n');
await ribbon('Calculation Options', () => page.click('.oxe-pop button:has(span:text-is("Manual"))'));
await page.evaluate(() => window.oxe.select('J45'));
await page.keyboard.type('10\n');
await page.keyboard.type('=J45*3\n');
const waiting = [await value('J47'), await value('J46'), await statusText()];
await page.keyboard.press('F9');
const calculated = [await value('J47'), await statusText()];
await ribbon('Calculation Options', () => page.click('.oxe-pop button:has(span:text-is("Automatic"))'));
await page.evaluate(() => window.oxe.select('J45'));
await page.keyboard.type('1\n');
check(waiting[0] === 12 && waiting[1] === 30 && /^Calculate/.test(waiting[2]) && calculated[0] === 20 && !/Calculate/.test(calculated[1]) && await value('J47') === 2,
  'Manual: J47 waits (12) and the status bar says Calculate; the formula entered in J46 is worked out (30); F9: 20; Automatic again: 2',
  JSON.stringify({ waiting, calculated, after: await value('J47') }));

// Formulas > Show Formulas (Ctrl+`): J47 shows =J45*2, its column drawn twice as wide; the button pressed. Again: 2.
const colWidth = () => page.evaluate(() => document.querySelector('[role="columnheader"][aria-colindex="11"]')?.getBoundingClientRect().width ?? null);
const widthBefore = await colWidth();
await page.focus('.oxe-grid');
await page.keyboard.press('Control+Backquote');
const showing = [await shown(46, 9), await colWidth(), await page.getAttribute('button[title="Show Formulas (Ctrl+`)"]', 'aria-pressed')];
await page.keyboard.press('Control+Backquote');
check(showing[0] === '=J45*2' && Math.abs(showing[1] - 2 * widthBefore) < 1 && showing[2] === 'true' && await shown(46, 9) === '2',
  'Ctrl+`: =J45*2 shown, the column twice as wide, Show Formulas pressed; again: the value', JSON.stringify({ showing, widthBefore, after: await shown(46, 9) }));

// Formulas > Trace Precedents on J47 (=J45*2): one arrow, from J45; again: J45 holds a value, nothing more. Trace
// Dependents from J45: J47's arrow (and J46's). Remove Arrows: none left.
const arrowCount = () => page.evaluate(() => document.querySelectorAll('.oxe-arrows polygon').length);
await page.evaluate(() => window.oxe.select('J47'));
await ribbon('Trace Precedents');
const first = await arrowCount();
await ribbon('Trace Precedents');
const second = await arrowCount();
await page.evaluate(() => window.oxe.select('J45'));
await ribbon('Trace Dependents');
const withDependents = await arrowCount();
await ribbon('Remove Arrows');
check(first === 1 && second === 1 && withDependents === 3 && await arrowCount() === 0, 'Trace Precedents: 1 arrow, then no more; Trace Dependents of J45: 2 more; Remove Arrows: none',
  JSON.stringify({ first, second, withDependents }));

// Excel's error indicator: '123 typed is a number stored as text, its green triangle and what it says; 123 typed, none.
const errMark = () => page.evaluate(() => document.querySelector('.oxe-row[aria-rowindex="46"] .oxe-cell[aria-colindex="12"]')?.getAttribute('aria-description') ?? null);
await page.evaluate(() => window.oxe.enter('K45', "'123"));
const flagged = await errMark();
await page.evaluate(() => window.oxe.enter('K45', '123'));
check(/formatted as text/.test(flagged || '') && await errMark() === null, "'123: the number-stored-as-text indicator; 123: none", JSON.stringify(flagged));

// The error button beside it: its menu headed Number Stored as Text; Convert to Number makes it 123 (Ctrl+Z: the text
// again); Alt+Shift+F10 opens the menu, Ignore Error takes the indicator away.
await page.evaluate(() => { window.oxe.enter('K45', "'123"); window.oxe.select('K45'); });
await page.focus('.oxe-grid');
await page.click('.oxe-err-btn');
const menuItems = await page.evaluate(() => [...document.querySelectorAll('.oxe-menu .oxe-menuhead, .oxe-menu .oxe-menuitem')].map((e) => e.textContent).join('|'));
await page.click('.oxe-menu button:text-is("Convert to Number")');
const converted = await value('K45');
await page.keyboard.press('Control+z');
const textAgain = await value('K45');
await page.keyboard.press('Alt+Shift+F10');
await page.click('.oxe-menu button:text-is("Ignore Error")');
check(menuItems === 'Number Stored as Text|Convert to Number|Ignore Error|Edit in Formula Bar' && converted === 123 && textAgain === '123' && await errMark() === null,
  'error button: Convert to Number, Ctrl+Z; Alt+Shift+F10 > Ignore Error', JSON.stringify({ menuItems, converted, textAgain }));

// Formulas > Error Checking from L45 ('7): its cell first; Convert to Number, then on to L46 (=1/0, Divide by Zero
// Error); Next until the whole sheet is checked.
await page.evaluate(() => { window.oxe.enter('L45', "'7"); window.oxe.enter('L46', '=1/0'); window.oxe.select('L45'); });
await ribbon('Error Checking');
await page.waitForSelector('dialog.oxe-errcheck-box');
const boxText = (sel) => page.textContent(`dialog.oxe-errcheck-box ${sel}`);
const firstError = [await boxText('.oxe-errcheck-where'), await boxText('.oxe-errcheck-name')];
await page.click('dialog.oxe-errcheck-box >> button:text-is("Convert to Number")');
const secondError = [await boxText('.oxe-errcheck-where'), await boxText('.oxe-errcheck-name'), await value('L45')];
let finalMessage = null;
for (let k = 0; k < 30 && !finalMessage; k++) {
  await page.click('dialog.oxe-errcheck-box >> button:text-is("Next")');
  finalMessage = await page.evaluate(() => [...document.querySelectorAll('dialog')].find((d) => /complete for the entire sheet/.test(d.textContent))?.textContent ?? null);
}
await page.keyboard.press('Enter');
check(firstError.join() === 'Error in cell L45,Number Stored as Text' && secondError.join() === 'Error in cell L46,Divide by Zero Error,7' && /complete/.test(finalMessage || ''),
  'Error Checking: L45 first, Convert to Number, on to L46; Next until the error check is complete', JSON.stringify({ firstError, secondError, finalMessage }));

// Formulas > Evaluate Formula on L47 (=IF(L45>5,L45*2,0)): L45 underlined first; Evaluate to 7, TRUE, 7*2's 14, then
// IF's 14; Restart at the end.
await page.evaluate(() => { window.oxe.enter('L47', '=IF(L45>5,L45*2,0)'); window.oxe.select('L47'); });
await ribbon('Evaluate Formula');
await page.waitForSelector('dialog.oxe-evaluate-box');
const evalSteps = [await page.evaluate(() => document.querySelector('.oxe-evaluate-text').innerHTML)];
for (let k = 0; k < 6; k++) {
  if (await page.textContent('dialog.oxe-evaluate-box .oxe-dialog-buttons button') === 'Restart') break;
  await page.click('dialog.oxe-evaluate-box >> button:text-is("Evaluate")');
  evalSteps.push(await page.evaluate(() => document.querySelector('.oxe-evaluate-text').textContent));
}
await page.click('dialog.oxe-evaluate-box >> button:text-is("Close")');
check(evalSteps[0] === '=IF(<u>L45</u>&gt;5,L45*2,0)' && evalSteps.slice(1).join('|') === '=IF(7>5,L45*2,0)|=IF(TRUE,L45*2,0)|=IF(TRUE,7*2,0)|=IF(TRUE,14,0)|=14',
  'Evaluate Formula: L45 underlined, then 7, TRUE, 7, 14, IF\'s 14', JSON.stringify(evalSteps));

// Formulas > Watch Window: L47 watched (its value 14 and formula); L45 changed to 9: 18 there at once; Delete Watch.
const watchRows = () => page.evaluate(() => [...document.querySelectorAll('.oxe-watch-table tbody tr')].map((r) => [...r.cells].slice(1).map((c) => c.textContent).join('|')));
await page.evaluate(() => window.oxe.select('L47'));
await ribbon('Watch Window');
await page.click('.oxe-watchbar button:text-is("Add Watch")');
const watched = await watchRows();
await page.evaluate(() => window.oxe.enter('L45', '9'));
const updated = await watchRows();
await page.click('.oxe-watch-table tbody tr');
await page.click('.oxe-watchbar button:text-is("Delete Watch")');
const left = await watchRows();
await ribbon('Watch Window');
check(watched.join() === 'Data||L47|14|=IF(L45>5,L45*2,0)' && updated.join() === 'Data||L47|18|=IF(L45>5,L45*2,0)' && left.length === 0,
  'Watch Window: L47 watched, 14 then 18 as L45 changes; Delete Watch', JSON.stringify({ watched, updated, left }));

await finish();
