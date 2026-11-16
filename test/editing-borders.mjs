// Browser check of Home > Font's Draw Borders and More Colors (docs/FORMATS.md): a drag in Draw Border Grid draws
// every line and selects nothing, one Ctrl+Z takes it back, a click in Erase Border takes a cell's lines off, Esc ends
// the mode, Line Style and More Colors... on Line Color draw a gridline clicked in that line, and More Colors... on Fill
// Color and Font Color give the selection a custom colour, written rgb="FFRRGGBB", kept in Recent Colors.
// Run after `npm run build`: node test/editing-borders.mjs [outDir]
import { editingPage } from './editing-kit.mjs';
import { openFixture } from './browser.mjs';

const { page, check, finish, pointAt, clickCell } = await editingPage();
await openFixture(page, 'basic.xlsx');
check(await page.evaluate(() => document.compatMode) === 'CSS1Compat', 'standards mode');

/** A choice of the Borders menu (or a menu it opened), by its exact text. */
const bordersMenu = async (...items) => {
  await page.click('.oxe-toolbar button[title="Bottom Border options"]');
  for (const item of items) await page.click(`.oxe-pop button:has(span:text-is("${item}"))`);
};
const colorMenu = (title) => page.click(`.oxe-toolbar button[title="${title} options"]`);
/** A cell's border sides as 'style' or '-' (left right top bottom), and its colour on one side. */
const sides = (ref) => page.evaluate((ref) => {
  const { r, c } = { r: Number(ref.slice(1)) - 1, c: ref.charCodeAt(0) - 65 };
  const wb = window.oxe.workbook;
  const s = wb.sheet(window.oxe.sheetIndex).cells.get(r)?.get(c)?.s || 0;
  const b = wb.styles.borders[wb.styles.cellXfs[s].borderId];
  return ['left', 'right', 'top', 'bottom'].map((k) => b[k]?.style || '-').join(' ');
}, ref);
const drag = async (from, to) => {
  const a = await pointAt(...from);
  const b = await pointAt(...to);
  await page.mouse.move(a.x, a.y);
  await page.mouse.down();
  await page.mouse.move((a.x + b.x) / 2, (a.y + b.y) / 2, { steps: 3 });
  await page.mouse.move(b.x, b.y, { steps: 3 });
  await page.mouse.up();
};
const mode = () => page.evaluate(() => document.querySelector('.oxe-root').classList.contains('oxe-drawing-borders') ? 'draw' : document.querySelector('.oxe-root').classList.contains('oxe-erasing-borders') ? 'erase' : null);

await clickCell(0, 0);
await bordersMenu('Draw Border Grid');
const on = await mode();
// (Column D of basic.xlsx is hidden: E3:F5.)
await drag([2, 4], [4, 5]);
const grid = await Promise.all(['E3', 'F4', 'F5'].map(sides));
const selection = await page.evaluate(() => window.oxe.selection);
check(on === 'draw' && grid.every((x) => x === 'thin thin thin thin') && selection.r0 === 0 && selection.c1 === 0,
  'Draw Border Grid: a drag over E3:F5 draws every line there; A1 stays selected', JSON.stringify({ on, grid, selection }));

await page.keyboard.press('Control+z');
const undone = await Promise.all(['E3', 'F5'].map(sides));
await page.keyboard.press('Control+y');
check(undone.every((x) => x === '- - - -') && await sides('F4') === 'thin thin thin thin', 'the drag is one step of the history (Ctrl+Z, Ctrl+Y)', JSON.stringify(undone));

await bordersMenu('Erase Border');
const erasing = await mode();
await clickCell(3, 5);
const erased = await sides('F4');
const beside = await sides('E4');
await page.focus('.oxe-grid');
await page.keyboard.press('Escape');
check(erasing === 'erase' && erased === '- - - -' && beside === 'thin - thin thin' && await mode() === null,
  'Erase Border: a click inside F4 takes its lines off and E4\'s facing right; Esc ends the mode', JSON.stringify({ erasing, erased, beside }));

// Line Style, then Line Color's More Colors...: hex 123456, and Draw Border on; a click at C8's bottom gridline.
await bordersMenu('Line Style ▸', 'Medium dashed');
await bordersMenu('Line Color ▸', 'More Colors...');
await page.waitForSelector('dialog.oxe-colorbox[open]');
await page.fill('dialog.oxe-colorbox input[id$="-hex"]', '#123456');
await page.click('dialog.oxe-colorbox >> button:text-is("OK")');
const lineMode = await mode();
const edge = await page.evaluate(() => {
  const g = window.oxe.grid;
  const v = g.el.getBoundingClientRect();
  const { rows, cols, rowHeadW, colHeadH } = g.geometry();
  return { x: v.left + rowHeadW + cols.start(2) - g.el.scrollLeft + 20, y: v.top + colHeadH + rows.start(8) - g.el.scrollTop - 2 };
});
await page.mouse.click(edge.x, edge.y);
const line = await page.evaluate(() => {
  const wb = window.oxe.workbook;
  const b = wb.styles.borders[wb.styles.cellXfs[wb.sheet(window.oxe.sheetIndex).cells.get(7)?.get(2)?.s || 0].borderId];
  return [b.bottom?.style, b.bottom?.color?.rgb, b.top?.style || '-'].join(' ');
});
await page.focus('.oxe-grid');
await page.keyboard.press('Escape');
check(lineMode === 'draw' && line === 'mediumDashed FF123456 -', 'Line Style and Line Color > More Colors...: Draw Border on, a gridline clicked drawn in that line', JSON.stringify({ lineMode, line }));

// Fill Color > More Colors...: Red 18, Green 52, Blue 86 on A1; then in Recent Colors, and written as Excel writes it.
await clickCell(0, 0);
const a1Was = await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(0)?.get(0)?.s || 0);
await colorMenu('Fill Color');
await page.click('.oxe-pop button:has(span:text-is("More Colors..."))');
await page.waitForSelector('dialog.oxe-colorbox[open]');
for (const [k, v] of [['red', '18'], ['green', '52'], ['blue', '86']]) await page.fill(`dialog.oxe-colorbox input[id$="-${k}"]`, v);
const hex = await page.inputValue('dialog.oxe-colorbox input[id$="-hex"]');
await page.click('dialog.oxe-colorbox >> button:text-is("OK")');
const fill = await page.evaluate(() => {
  const wb = window.oxe.workbook;
  const s = wb.sheet(window.oxe.sheetIndex).cells.get(0)?.get(0)?.s || 0;
  return wb.styles.fills[wb.styles.cellXfs[s].fillId];
});
await colorMenu('Font Color');
const recent = await page.evaluate(() => [...document.querySelectorAll('.oxe-pop .oxe-cap')].map((x) => x.textContent).includes('Recent Colors') && !!document.querySelector('.oxe-pop button[title="#123456"]'));
await page.keyboard.press('Escape');
// (The styles part as it will be saved.)
const xml = await page.evaluate(() => new XMLSerializer().serializeToString(window.oxe.workbook.stylesDoc));
check(hex === '#123456' && fill?.fg?.rgb === 'FF123456' && recent && /<mruColors><color rgb="FF123456"\/><\/mruColors>/.test(xml) && /<fgColor rgb="FF123456"\/>/.test(xml),
  'Fill Color > More Colors...: RGB 18, 52, 86 (hex #123456) fills A1 rgb="FF123456", shown in Recent Colors and kept in mruColors',
  JSON.stringify({ hex, fill, recent, mru: xml.match(/<colors>[\s\S]*?<\/colors>/)?.[0] }));

await page.keyboard.press('Control+z');
const unfilled = await page.evaluate(() => window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(0)?.get(0)?.s || 0);
check(unfilled === a1Was, 'More Colors\' fill is one step of the history (Ctrl+Z: A1 as it was)', String(unfilled));

// Font Color > More Colors...: hex abcdef as A1's text colour; the button keeps it as the last colour.
await colorMenu('Font Color');
await page.click('.oxe-pop button:has(span:text-is("More Colors..."))');
await page.fill('dialog.oxe-colorbox input[id$="-hex"]', 'abcdef');
await page.keyboard.press('Enter');
const font = await page.evaluate(() => {
  const wb = window.oxe.workbook;
  const s = wb.sheet(window.oxe.sheetIndex).cells.get(0)?.get(0)?.s || 0;
  return wb.styles.fonts[wb.styles.cellXfs[s].fontId].color?.rgb;
});
const swatch = await page.evaluate(() => getComputedStyle(document.querySelector('.oxe-fontcolor')).getPropertyValue('--oxe-swatch').trim());
check(font === 'FFABCDEF' && /abcdef|171, 205, 239/i.test(swatch), 'Font Color > More Colors...: abcdef on A1, and the button\'s colour', JSON.stringify({ font, swatch }));

// A merged range draws only its outline: H20:J22 given All Borders (each cell's four sides, 36 edges), then merged,
// keeps the 12 sides on the range's edge; the covered cells' sides inside it aren't drawn, as in Excel.
const edgeCount = () => page.evaluate(() => document.querySelectorAll('.oxe-edges > *').length);
const before = await edgeCount();
await page.evaluate(() => { window.oxe.command('select', 19, 7); window.oxe.command('extendTo', 21, 9); window.oxe.command('borders', 'all'); });
await page.waitForTimeout(100);
const bordered = await edgeCount() - before;
await page.evaluate(() => window.oxe.command('merge', 'cells'));
await page.waitForTimeout(100);
const merged = await edgeCount() - before;
check(bordered === 36 && merged === 12, 'All Borders then Merge Cells: the 3x3 range\'s 36 edges become its 12 outline edges', JSON.stringify({ bordered, merged }));

await finish();
