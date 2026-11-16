// Browser check of Find by format, @mentions in threaded comments, Freeze Panes and AutoSum, on a fresh basic.xlsx.
// Run after `npm run build`: node test/editing-leftovers.mjs [outDir]
import { openFixture } from './browser.mjs';
import { editingPage } from './editing-kit.mjs';

const { page, check, finish, sel, clickCell, menu } = await editingPage();
await openFixture(page, 'basic.xlsx');

// Find by format: B2 made bold, Ctrl+F, Find Format's From Cell... and a click on B2: the bar says whose format;
// Find All with nothing typed finds B2 and no plain cell; Clear takes the format off.
await page.evaluate(() => window.oxe.select('B2'));
await page.focus('.oxe-grid');
await page.keyboard.press('Control+b');
await page.evaluate(() => window.oxe.select('A1'));
await page.focus('.oxe-grid');
await page.keyboard.press('Control+f');
await page.waitForSelector('.oxe-findbar:not([hidden])');
await page.click('.oxe-findrow button[title="Choose Format From Cell (Find Format)"]');
const picking = await page.evaluate(() => window.oxe.grid.el.classList.contains('oxe-picking'));
await clickCell(1, 1);
await page.waitForFunction(() => /Format of B2/.test(document.querySelector('.oxe-findrow .oxe-find-format')?.textContent || ''));
const shownFormat = await page.textContent('.oxe-findrow .oxe-find-format');
check(picking && /Format of B2/.test(shownFormat) && await page.evaluate(() => document.querySelector('.oxe-findrow .oxe-find-format').classList.contains('oxe-find-format-bold')),
  'From Cell..., a click on B2: picking shown on the grid, then "Format of B2" in bold', shownFormat);
await page.click('.oxe-findbar button[title="Find All"]');
const hits = await page.evaluate(() => [...document.querySelectorAll('.oxe-findall-row span:nth-child(2)')].map((x) => x.textContent));
check(hits.includes('$B$2') && !hits.includes('$A$2') && !hits.includes('$B$3'), 'Find All with B2’s format: B2 found, its plain neighbours not', hits.join(' '));
// (The keyboard way: From Cell..., an arrow to B3 on the sheet, Enter.)
await page.click('.oxe-findrow button[title="Choose Format From Cell (Find Format)"]');
await page.keyboard.press('ArrowDown');
await page.keyboard.press('Enter');
await page.waitForFunction(() => /Format of B3/.test(document.querySelector('.oxe-findrow .oxe-find-format')?.textContent || ''));
check(true, 'From Cell... by keys: an arrow to B3, Enter takes its format');
await page.click('.oxe-findrow button[title="Clear Find Format"]');
check(await page.textContent('.oxe-findrow .oxe-find-format') === 'No Format Set', 'Clear: No Format Set');
await page.keyboard.press('Escape');

// @mentions: New Comment on D3, "Please check @sa": the demo's people narrowed to Sam Lee, Enter puts "@Sam Lee " in;
// posted, the thread's XML has the mention and the person (PeoplePicker), and the box shows the name highlighted.
await page.evaluate(() => window.oxe.select('D3'));
await page.focus('.oxe-grid');
const box = 'dialog.oxe-thread-box';
await page.keyboard.press('Control+Alt+m');
await page.waitForSelector(box);
await page.click(`${box} > textarea`);
await page.keyboard.type('Please check @sa');
await page.waitForSelector(`${box} .oxe-mention-list:not([hidden])`);
const offered = await page.evaluate((b) => [...document.querySelectorAll(`${b} .oxe-mention-option`)].map((x) => x.firstChild.textContent), box);
const combo = await page.evaluate((b) => { const t = document.querySelector(`${b} > textarea`); return [t.getAttribute('role'), t.getAttribute('aria-expanded'), !!document.getElementById(t.getAttribute('aria-activedescendant'))]; }, box);
check(offered.join() === 'Sam Lee' && combo.join() === 'combobox,true,true', '@sa: Sam Lee offered, the box a combobox with the option active', JSON.stringify({ offered, combo }));
await page.keyboard.press('Enter');
await page.keyboard.type('please');
const typed = await page.inputValue(`${box} > textarea`);
check(typed === 'Please check @Sam Lee please', 'Enter: "@Sam Lee " put in, the list closed', typed);
await page.click(`${box} >> button:text-is("Post")`);
await page.waitForSelector(`${box} .oxe-thread-item .oxe-mention`);
const written = await page.evaluate(() => {
  const wb = window.oxe.workbook;
  const rel = wb.pkg.readRels(wb.sheets[window.oxe.sheetIndex].path).find((r) => /threadedComment$/.test(r.type));
  const xml = wb.pkg.text(rel.target);
  const persons = wb.pkg.text('xl/persons/person.xml');
  return { mention: /<mention [^>]*startIndex="13" length="8"\/>/.test(xml), sam: /displayName="Sam Lee"[^>]*userId="sam\.lee@example\.com" providerId="PeoplePicker"/.test(persons) };
});
const mark = await page.textContent(`${box} .oxe-thread-item .oxe-mention`);
check(written.mention && written.sam && mark === '@Sam Lee', 'Posted: the mention written (start 13, length 8), Sam Lee a PeoplePicker person, "@Sam Lee" highlighted', JSON.stringify({ written, mark }));
// (Esc in the list closes the list, not the box.)
await page.click(`${box} > textarea`);
await page.keyboard.type('@r');
await page.waitForSelector(`${box} .oxe-mention-list:not([hidden])`);
await page.keyboard.press('Escape');
check(await page.isVisible(box) && await page.isHidden(`${box} .oxe-mention-list`), 'Esc with the list open: the list closed, the box still open');
await page.click(`${box} >> button:text-is("Close")`);

// Freeze Panes while scrolled, as Excel freezes it (test/freeze-excel.json): the window at C10, E15 active: rows 10-14
// and columns C-D frozen, the sheet view starting at C10; the window can't go above or left of it; a click just under
// the headings selects C10. Unfreeze: the window back at C10, nothing frozen.
await page.evaluate(() => { window.oxe.select('E15'); window.oxe.grid.scrollToCell(9, 2); });
await page.waitForTimeout(50);
await page.click('.oxe-rtab >> text="View"');
await menu('Freeze Panes', 'Freeze Panes');
const frozenNow = await page.evaluate(() => ({ grid: window.oxe.grid.view().freeze, file: window.oxe.workbook.sheet(0).freeze }));
check(JSON.stringify(frozenNow.grid) === '{"rows":5,"cols":2,"top":9,"left":2}' && JSON.stringify(frozenNow.file) === JSON.stringify(frozenNow.grid),
  'scrolled to C10, Freeze Panes at E15: rows 10-14 and columns C-D, from C10', JSON.stringify(frozenNow));
const viewXml = await page.evaluate(() => { const wb = window.oxe.workbook; const p = wb._part(0); return new XMLSerializer().serializeToString(p.doc).match(/<sheetView\b[\s\S]*?<\/sheetView>/)[0]; });
check(/topLeftCell="C10"/.test(viewXml) && /<pane xSplit="2" ySplit="5" topLeftCell="E15" activePane="bottomRight" state="frozen"\/>/.test(viewXml), 'written as Excel: the sheet view from C10, the pane from E15', viewXml);
await page.evaluate(() => { window.oxe.grid.el.scrollTop = 0; window.oxe.grid.el.scrollLeft = 0; });
await page.waitForTimeout(100);
const clamped = await page.evaluate(() => { const g = window.oxe.grid; const geo = g.geometry(); return [g.el.scrollTop >= geo.rows.start(9) - 1, g.el.scrollLeft >= geo.cols.start(2) - 1]; });
check(clamped.join() === 'true,true', 'scrolled to the top: kept at where the freeze starts', clamped.join());
const c10 = await page.evaluate(() => { const g = window.oxe.grid; const v = g.el.getBoundingClientRect(); const geo = g.geometry(); return { x: v.left + geo.rowHeadW + 6, y: v.top + geo.colHeadH + 6 }; });
await page.mouse.click(c10.x, c10.y);
const picked = await sel();
check(picked.ar === 9 && picked.ac === 2, 'a click just under the headings: C10, the frozen corner’s first cell', JSON.stringify(picked));
const paneRows = await page.evaluate(() => [...document.querySelectorAll('.oxe-pane-top .oxe-rowhead')].map((x) => x.textContent).join(','));
check(paneRows.startsWith('10,11,12,13,14'), 'the top pane shows rows 10 to 14', paneRows);
await menu('Freeze Panes', 'Unfreeze Panes');
const back = await page.evaluate(() => ({ freeze: window.oxe.grid.view().freeze, topLeft: window.oxe.grid.view().topLeft }));
check(back.freeze.rows === 0 && back.freeze.cols === 0 && back.topLeft.r === 9 && back.topLeft.c === 2, 'Unfreeze Panes: the window back at C10, as Excel leaves it', JSON.stringify(back));
await page.click('.oxe-rtab >> text=Home');

// AutoSum: 1, 2, 3 typed into H20:H22 and 7, 8 into J20:J21; Alt+= on H23 opens it with Excel's proposal, its range
// outlined on the sheet (the range finder), Enter enters it. H20:H22 and J20:J21 selected (Ctrl+drag): AutoSum puts
// a total under each area at once (Excel's on two areas), both areas grown by it.
const typeDown = async (ref, values) => {
  await page.evaluate((r) => window.oxe.select(r), ref);
  await page.focus('.oxe-grid');
  for (const v of values) { await page.keyboard.type(v); await page.keyboard.press('Enter'); }
};
await typeDown('H20', ['1', '2', '3']);
await typeDown('J20', ['7', '8']);
await page.evaluate(() => window.oxe.select('H23'));
await page.focus('.oxe-grid');
await page.keyboard.press('Alt+=');
await page.waitForSelector('.oxe-mark');
const proposal = await page.evaluate(() => { const e = document.querySelector('.oxe-editor'); return e?.value ?? e?.textContent ?? ''; });
const marks = await page.evaluate(() => document.querySelectorAll('.oxe-overlay .oxe-mark').length);
check(proposal === '=SUM(H20:H22)' && marks >= 1, 'Alt+= on H23: =SUM(H20:H22) proposed, its range outlined', JSON.stringify({ proposal, marks }));
await page.keyboard.press('Enter');
const h23 = await page.evaluate(() => { const c = window.oxe.workbook.sheet(0).cells.get(22)?.get(7); return [c?.f, c?.v]; });
check(h23[0] === 'SUM(H20:H22)' && h23[1] === 6, 'Enter: H23 =SUM(H20:H22), 6', JSON.stringify(h23));
await clickCell(19, 7);
const h22 = await (async () => { const p = await page.evaluate(() => { const g = window.oxe.grid; const v = g.el.getBoundingClientRect(); const { rows, cols, rowHeadW, colHeadH } = g.geometry(); return { x: v.left + rowHeadW + cols.start(7) - g.el.scrollLeft + 6, y: v.top + colHeadH + rows.start(21) - g.el.scrollTop + 6 }; }); return p; })();
await page.keyboard.down('Shift');
await page.mouse.click(h22.x, h22.y);
await page.keyboard.up('Shift');
const j = await page.evaluate(() => { const g = window.oxe.grid; const v = g.el.getBoundingClientRect(); const { rows, cols, rowHeadW, colHeadH } = g.geometry(); const at = (r, c) => ({ x: v.left + rowHeadW + cols.start(c) - g.el.scrollLeft + 6, y: v.top + colHeadH + rows.start(r) - g.el.scrollTop + 6 }); return [at(19, 9), at(20, 9)]; });
await page.keyboard.down('Control');
await page.mouse.move(j[0].x, j[0].y);
await page.mouse.down();
await page.mouse.move(j[1].x, j[1].y, { steps: 3 });
await page.mouse.up();
await page.keyboard.up('Control');
await page.keyboard.press('Alt+=');
const totals = await page.evaluate(() => { const at = (r, c) => window.oxe.workbook.sheet(0).cells.get(r)?.get(c)?.f; return [at(22, 7), at(21, 9)]; });
const areasNow = await page.evaluate(() => { const s = window.oxe.selection; return [...(s.more || []), s].map((a) => `${a.r0},${a.c0}:${a.r1},${a.c1}`).join(' '); });
check(totals[1] === 'SUM(J20:J21)' && areasNow === '19,7:22,7 19,9:21,9', 'two areas, AutoSum: J22 =SUM(J20:J21), H20:H23 (its total already there) and J20:J22 selected', JSON.stringify({ totals, areasNow }));
void sel;

await finish();
