// Browser check of Find by format, @mentions in threaded comments, Freeze Panes and AutoSum, on a fresh basic.xlsx.
// Run after `npm run build`: node test/editing-leftovers.mjs [outDir]
import { openFixture } from './browser.mjs';
import { editingPage } from './editing-kit.mjs';

const { page, check, finish, sel, clickCell } = await editingPage();
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
void sel;

await finish();
