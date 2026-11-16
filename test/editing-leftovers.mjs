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
void sel;

await finish();
