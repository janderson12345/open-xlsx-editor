// Browser check of a chart's formatting: a clustered column chart inserted on a new sheet's table and picked, then
// the Format Chart box (Ctrl+1, and the chart's menu): an axis title, the chart area's and a series' fill, the title's
// font, and Cost drawn as a line on the secondary axis; each one step of Undo, the part written as Excel writes it
// (chartformat.test.js), and the saved workbook read back the same.
// Run after `npm run build`: node test/editing-charts.mjs
import { editingPage } from './editing-kit.mjs';

const { page, check, finish, menu, openFixtureFrom } = await editingPage();
const box = 'dialog.oxe-chart-format-box';
await openFixtureFrom(page, 'test/fixtures/basic.xlsx', 'basic.xlsx');

await page.evaluate(() => {
  const o = window.oxe;
  o.command('insertSheet');
  [['Region', 'Sales', 'Cost'], ['North', 120, 80], ['South', 95, 70], ['East', 140, 90], ['West', 60, 45]].forEach((row, r) => row.forEach((v, c) => o.enter(String.fromCharCode(65 + c) + (r + 1), String(v))));
  o.select('A1');
});
await page.click('.oxe-rtab >> text=Insert');
await menu('Insert Column or Bar Chart', 'Clustered Column');
await page.waitForSelector('.oxe-chart');
const part = () => page.evaluate(() => {
  const wb = window.oxe.workbook;
  const drawing = wb.pkg.readRels(wb.sheets[window.oxe.sheetIndex].path).find((r) => /drawing$/.test(r.type));
  const chart = wb.pkg.readRels(drawing.target).find((r) => /chart$/.test(r.type));
  return wb.pkg.text(chart.target);
});
const svgText = () => page.evaluate(() => [...document.querySelectorAll('.oxe-chart svg text')].map((t) => t.textContent));
const pickChart = async () => { const b = await page.locator('.oxe-chart').boundingBox(); await page.mouse.click(b.x + 60, b.y + 60); };
/** The Format Chart box opened with Ctrl+1 on the picked chart, the part chosen and `fill` done, then OK. */
async function format(partKey, fill) {
  await pickChart();
  await page.keyboard.press('Control+1');
  await page.waitForSelector(box);
  await page.selectOption(`${box} select >> nth=0`, partKey);
  await fill();
  await page.click(`${box} >> button:text-is("OK")`);
  await page.waitForSelector(box, { state: 'detached' });
}

// (The vertical axis's title: Excel's 10pt title turned to read upwards, drawn left of the labels.)
await format('v', async () => {
  await page.selectOption(`${box} .oxe-chart-format-title select`, 'text');
  await page.fill(`${box} .oxe-chart-format-title input`, 'Amount');
});
const amount = await page.evaluate(() => [...document.querySelectorAll('.oxe-chart svg text')].find((t) => t.textContent === 'Amount')?.getAttribute('transform') || '');
check(/^rotate\(-90 /.test(amount) && /<c:valAx>(?:(?!<\/c:valAx>)[\s\S])*<c:title><c:tx><c:rich><a:bodyPr rot="-5400000"/.test(await part()),
  'Format Chart, Vertical Axis, title "Amount": drawn turned beside the axis, written as Excel writes it', amount);

// (The chart area's fill and a series' fill, from Excel's standard colours.)
await format('chart', () => page.selectOption(`${box} .oxe-chart-format-look select >> nth=0`, 'FFC000'));
await format('s0', () => page.selectOption(`${box} .oxe-chart-format-look select >> nth=0`, 'C00000'));
const fills = await page.evaluate(() => [...document.querySelectorAll('.oxe-chart svg rect')].map((r) => (r.getAttribute('fill') || '').toUpperCase()));
check(fills[0] === '#FFC000' && fills.filter((f) => f === '#C00000').length === 5,
  'Format Chart: the chart area Orange, Sales Dark Red (its four columns and its legend key)', fills.slice(0, 3).join(' '));

// (The title's font: 18pt, bold.)
await format('title', async () => {
  await page.fill(`${box} .oxe-chart-format-font input[type=number]`, '18');
  await page.selectOption(`${box} .oxe-chart-format-font select >> nth=1`, '1');
});
const title = await page.evaluate(() => { const t = [...document.querySelectorAll('.oxe-chart svg text')].find((x) => x.textContent === 'Chart Title'); return [t.getAttribute('font-size'), t.getAttribute('font-weight')].join(' '); });
check(title === '24 bold', 'Format Chart, Chart Title, 18pt bold: drawn at 24px, bold', title);

// (Cost as a line on the secondary axis: Change Series Chart Type, one step of Undo.)
const before = await part();
await format('s1', async () => {
  await page.selectOption(`${box} .oxe-chart-format-series select`, 'line');
  await page.check(`${box} .oxe-chart-format-series input[type=checkbox]`);
});
const comboPart = await part();
const line = await page.evaluate(() => !![...document.querySelectorAll('.oxe-chart svg path')].find((p) => p.getAttribute('stroke') && p.getAttribute('fill') === 'none'));
const texts = await svgText();
check(/<\/c:barChart><c:lineChart>/.test(comboPart) && /<c:axPos val="r"\/>/.test(comboPart) && line && texts.includes('100') && texts.includes('160'),
  'Format Chart, Series "Cost", Line on the secondary axis: a line group, the secondary axis drawn (0 to 100) beside the primary (0 to 160)', texts.join(','));
await page.keyboard.press('Escape');
await page.focus('.oxe-grid');
await page.keyboard.press('Control+z');
await page.waitForFunction((b) => {
  const wb = window.oxe.workbook;
  const drawing = wb.pkg.readRels(wb.sheets[window.oxe.sheetIndex].path).find((r) => /drawing$/.test(r.type));
  return wb.pkg.text(wb.pkg.readRels(drawing.target).find((r) => /chart$/.test(r.type)).target) === b;
}, before);
check(true, 'Ctrl+Z: the series back as columns on the primary axis, the part as it was, in one step');
await page.keyboard.press('Control+y');

// (The chart's menu offers the box; the saved workbook reads back with its formatting.)
const b = await page.locator('.oxe-chart').boundingBox();
await page.mouse.click(b.x + 60, b.y + 60, { button: 'right' });
await page.waitForSelector('.oxe-menu');
const items = await page.evaluate(() => [...document.querySelectorAll('.oxe-menu .oxe-menuitem')].map((x) => x.textContent.trim()));
check(items.includes('Format Chart...'), 'the chart\'s menu: Format Chart...', items.join(' | '));
await page.keyboard.press('Escape');
const [bytes, shownSheet] = await page.evaluate(() => [[...window.oxe.save()], window.oxe.sheetIndex]);
const { Workbook } = await import('../src/ooxml/workbook.js');
const { readCharts } = await import('../src/ooxml/charts.js');
const again = Workbook.fromBytes(new Uint8Array(bytes));
const c = readCharts(again, shownSheet)[0];
const got = [c.axis.title?.text, c.look.area.fill?.rgb, c.series[0].look.fill?.rgb, c.look.fonts.title.size, c.series.map((s) => `${s.type}${s.secondary ? '/2' : ''}`).join(' ')].join(', ');
check(got === 'Amount, FFC000, C00000, 18, bar line/2', 'saved and opened again: the title, fills, font and the line on the secondary axis read back', got);

await finish();
