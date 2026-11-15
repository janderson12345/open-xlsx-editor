// Browser check of a chart's Select Data (chartdatabox.js, refpick.js, acts-chartdata.js) on the charts fixture: the
// chart's menu opens the box with its range and series; the range picked on the sheet with the collapse button gives
// the chart a third series; Switch Row/Column; a series added, edited, moved and removed in the Edit Series box and
// the list; Hidden and Empty Cells; Cancel taking every change back, OK keeping them as steps Ctrl+Z undoes one at a
// time; and the saved chart part holding the series as chartdata.js writes them.
// Run after `npm run build`: node test/editing-chartdata.mjs [outDir]
import { editingPage } from './editing-kit.mjs';
import { openFixture } from './browser.mjs';

const { page, check, finish, pointAt } = await editingPage();
await openFixture(page, 'charts.xlsx');
await page.waitForSelector('.oxe-chart');

const box = 'dialog.oxe-select-data';
const listed = () => page.evaluate(() => [...document.querySelectorAll('.oxe-series-list option')].map((o) => o.textContent));
/** Chart 0's series names as drawn (its figure for screen readers: the table's header row). */
const drawn = () => page.evaluate(() => [...document.querySelectorAll('.oxe-chart-figures table')][0]?.rows[0] && [...[...document.querySelectorAll('.oxe-chart-figures table')][0].rows[0].cells].slice(1).map((c) => c.textContent));
const openBox = async () => {
  const b = await page.locator('.oxe-chart[data-oxe-chart="0"]').boundingBox();
  await page.mouse.click(b.x + 80, b.y + 80, { button: 'right' });
  await page.click('.oxe-menu .oxe-menuitem:text-is("Select Data...")');
  await page.waitForSelector(box);
};
const click = (label) => page.click(`${box} >> button:text-is("${label}")`);

await openBox();
check(await page.inputValue(`${box} .oxe-ref-input`) === '=Sales!$A$1:$C$7' && (await listed()).join() === 'North,South',
  'the chart\'s menu, Select Data...: its range and its series', `${await page.inputValue(`${box} .oxe-ref-input`)} ${await listed()}`);

// The collapse button: the box steps aside, A1:D7 dragged over on the sheet, Enter: the range taken, a third series.
await page.click(`${box} .oxe-ref-collapse`);
await page.waitForSelector('.oxe-refbar');
const hidden = await page.evaluate(() => !document.querySelector('dialog.oxe-select-data[open]'));
const a = await pointAt(0, 0);
const d = await pointAt(6, 3);
await page.mouse.move(a.x, a.y);
await page.mouse.down();
await page.mouse.move(d.x, d.y, { steps: 5 });
await page.mouse.up();
await page.waitForFunction(() => document.querySelector('.oxe-refbar input')?.value === '=Sales!$A$1:$D$7');
await page.keyboard.press('Enter');
await page.waitForSelector(`${box}[open]`);
check(hidden && (await listed()).join() === 'North,South,West' && (await drawn())?.join() === 'North,South,West',
  'the range picked on the sheet (the box stepped aside, Enter back): three series, drawn', `${await listed()} / ${await drawn()}`);

await click('Switch Row/Column');
check((await listed()).join() === 'Jan,Feb,Mar,Apr,May,Jun', 'Switch Row/Column: a series for each month', String(await listed()));
await click('Switch Row/Column');

// Add: the Edit Series box, its name typed and its values a reference.
await click('Add');
await page.waitForSelector('dialog.oxe-series-box');
await page.fill('dialog.oxe-series-box input >> nth=0', 'Target');
await page.fill('dialog.oxe-series-box input >> nth=1', '=Sales!$B$2:$B$7');
await page.click('dialog.oxe-series-box >> button:text-is("OK")');
check((await listed()).join() === 'North,South,West,Target', 'Add: a series named as typed, last', String(await listed()));
// Edit it: its name from a cell now; Move Up; Remove the first.
await page.selectOption(`${box} .oxe-series-list`, '3');
await click('Edit');
await page.waitForSelector('dialog.oxe-series-box');
await page.fill('dialog.oxe-series-box input >> nth=0', '=Sales!$A$1');
await page.click('dialog.oxe-series-box >> button:text-is("OK")');
await click('Move Up');
await page.selectOption(`${box} .oxe-series-list`, '0');
await click('Remove');
check((await listed()).join() === 'South,Month,West', 'Edit (its name from a cell), Move Up, Remove: the list', String(await listed()));
await click('Hidden and Empty Cells');
await page.selectOption('dialog.oxe-dialog:has-text("Hidden and Empty Cell Settings") select', 'zero');
await page.click('dialog.oxe-dialog:has-text("Hidden and Empty Cell Settings") >> button:text-is("OK")');

await click('Cancel');
check((await drawn())?.join() === 'North,South', 'Cancel: every change taken back', String(await drawn()));

// Again, kept with OK: each change a step of Undo.
await openBox();
await page.fill(`${box} .oxe-ref-input`, '=Sales!$A$1:$D$7');
await page.press(`${box} .oxe-ref-input`, 'Tab');
await page.selectOption(`${box} .oxe-series-list`, '2');
await click('Move Up');
await click('OK');
check((await drawn())?.join() === 'North,West,South', 'OK: the range typed and a series moved, drawn', String(await drawn()));
const part = () => page.evaluate(() => {
  const wb = window.oxe.workbook;
  const drawing = wb.pkg.readRels(wb.sheets[0].path).find((r) => /drawing$/.test(r.type)).target;
  const rel = wb.pkg.readRels(drawing).find((r) => /chart$/.test(r.type));
  return wb.pkg.text(rel.target);
});
const sers = (xml) => [...xml.matchAll(/<c:ser><c:idx val="(\d+)"\/><c:order val="(\d+)"\/>/g)].map((m) => `${m[1]}/${m[2]}`).join(' ');
const xml = await part();
check(sers(xml) === '0/0 2/1 1/2', 'the chart part: each series keeps its idx, the orders numbered again (as Excel writes PlotOrder)', sers(xml));
await page.focus('.oxe-grid');
await page.keyboard.press('Escape');
await page.keyboard.press('Control+z');
const once = await drawn();
await page.keyboard.press('Control+z');
check(once?.join() === 'North,South,West' && (await drawn())?.join() === 'North,South', 'Ctrl+Z: the move undone, then the range', `${once} / ${await drawn()}`);

await finish();
