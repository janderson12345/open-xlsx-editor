// Browser check of pictures, shapes, icons and groups on the grid (docs/LAYOUT.md, Pictures, shapes and icons) on
// test/fixtures/drawings.xlsx: each drawn where its anchor puts it, following the zoom and frozen panes; picked,
// moved with the mouse and the arrow keys, deleted, each one step of the history; Go To Special > Objects picking
// them all, and Excel's message when there are none; printed and exported to PDF.
// Run after `npm run build`: node test/drawings.mjs [outDir]
import { editingPage } from './editing-kit.mjs';
import { openFixture } from './browser.mjs';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';

const { page, check, finish, outDir, ribbon } = await editingPage();
await openFixture(page, 'drawings.xlsx');
await page.waitForSelector('.oxe-obj');

/** Object k's box on the page and where the grid's geometry puts its anchor's cells. */
const placed = (k) => page.evaluate((k) => {
  const g = window.oxe.grid;
  const el = document.querySelector(`.oxe-charts > .oxe-obj[data-oxe-ord="${k}"]`);
  if (!el) return null;
  const b = el.getBoundingClientRect();
  const v = g.el.getBoundingClientRect();
  const { rowHeadW, colHeadH } = g.geometry();
  const z = (g.view().zoom || 100) / 100;
  return { x: (b.left - v.left + g.el.scrollLeft) / z - rowHeadW, y: (b.top - v.top + g.el.scrollTop) / z - colHeadH, w: b.width / z, h: b.height / z };
}, k);
const near = (a, b) => a && Object.keys(b).every((k) => Math.abs(a[k] - b[k]) <= 1);
/** Where the grid's geometry puts cells c0, r0 to c1, r1 (their top left corners), in sheet pixels. */
const cells = (c0, r0, c1, r1) => page.evaluate(([c0, r0, c1, r1]) => {
  const { rows, cols } = window.oxe.grid.geometry();
  return { x: cols.start(c0), y: rows.start(r0), w: cols.start(c1) - cols.start(c0), h: rows.start(r1) - rows.start(r0) };
}, [c0, r0, c1, r1]);
/** The anchor in the drawing part holding the object named `name`. */
const anchorOf = (xml, name) => xml.split(/(?=<xdr:(?:twoCellAnchor|oneCellAnchor|absoluteAnchor)\b)/).find((a) => a.includes(`name="${name}"`)) || '';
const fromCol = (xml, name) => Number(/<xdr:from><xdr:col>(\d+)<\/xdr:col>/.exec(anchorOf(xml, name))?.[1]);
const fromRow = (xml, name) => Number(/<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/.exec(anchorOf(xml, name))?.[1]);
const drawing = () => page.evaluate(() => window.oxe.workbook.pkg.text('xl/drawings/drawing1.xml'));
const count = () => page.evaluate(() => document.querySelectorAll('.oxe-charts > .oxe-obj').length);
const picked = () => page.evaluate(() => [...document.querySelectorAll('.oxe-charts > .oxe-obj.oxe-chart-on')].map((e) => Number(e.dataset.oxeOrd)));

// Drawn: all but the hidden one, each where its anchor puts it (column C 120px, G hidden, row 6 hidden).
check(await count() === 15, 'fifteen objects drawn, the hidden one not', String(await count()));
const q0 = await cells(1, 1, 4, 8);
check(near(await placed(0), q0) && q0.h === 120, 'Quarters over its cells, B2 to E8, row 6 hidden under it', JSON.stringify([await placed(0), q0]));
const q1 = await cells(5, 1, 9, 8);
check(near(await placed(1), q1) && q1.w === 192, 'Cropped over F2 to J8, hidden column G under it', JSON.stringify([await placed(1), q1]));
const cellBox = await page.evaluate(() => { const c = document.querySelector('.oxe-row[aria-rowindex="3"] .oxe-cell[aria-colindex="2"]').getBoundingClientRect(); const o = document.querySelector('.oxe-charts > .oxe-obj[data-oxe-ord="4"]'); return [c.left, o && o.getBoundingClientRect().left]; });
check(Math.abs(cellBox[1] - cellBox[0] - 64) <= 1, 'Rectangle 5\'s left edge on column B\'s, past the row headings', JSON.stringify(cellBox));
const images = await page.evaluate(() => [...document.querySelectorAll('.oxe-obj image')].map((i) => i.getAttribute('href').slice(0, 22)));
check(images.length === 4 && images.filter((s) => s.startsWith('data:image/png')).length === 3 && images.some((s) => s.startsWith('data:image/svg+xml')), 'the pictures drawn from their PNGs, the icon from its SVG', JSON.stringify(images));
const texts = await page.evaluate(() => [...document.querySelectorAll('.oxe-obj-text')].map((t) => t.textContent));
check(texts.includes('Rounded') && texts.includes('Oval') && texts.some((t) => t.startsWith('A text box')), 'the shapes\' text', JSON.stringify(texts));
const fills = await page.evaluate(() => [...document.querySelectorAll('.oxe-obj[data-oxe-ord="7"] path, .oxe-obj[data-oxe-ord="5"] path')].map((p) => `${p.getAttribute('fill')}/${p.getAttribute('stroke')}`));
check(fills.includes('#4472C4/#2F528F') && fills.includes('#C55A11/null'), 'the style\'s accent 1 and its darker line; accent 2 at 75%', JSON.stringify(fills));
const figures = await page.evaluate(() => [...document.querySelectorAll('.oxe-chart-figures figcaption')].map((f) => f.textContent));
check(figures.includes('Picture: Quarters. Four coloured quarters'), 'for a screen reader: each object a figure with its alternative text', JSON.stringify(figures.slice(0, 4)));
await page.screenshot({ path: `${outDir}/drawings.png` });

// The zoom: at 150%, the same place on the sheet.
await page.evaluate(() => window.oxe.command('zoom', 150));
const zoomed = await page.evaluate(() => document.querySelector('.oxe-charts > .oxe-obj[data-oxe-ord="0"]').getBoundingClientRect().width);
check(Math.abs(zoomed - q0.w * 1.5) <= 1.5, 'zoomed to 150%, Quarters 1.5 times as wide', String(zoomed));
await page.evaluate(() => window.oxe.command('zoom', 100));

// Frozen panes: rows 1 to 4 frozen, the sheet scrolled down: Quarters still drawn in the frozen rows.
await page.evaluate(() => { window.oxe.select('A5'); window.oxe.command('freeze', 'panes'); });
await page.evaluate(() => { window.oxe.grid.el.scrollTop = 300; });
await page.waitForTimeout(150);
const inPane = await page.evaluate(() => {
  const el = document.querySelector('.oxe-pane-top .oxe-obj[data-oxe-ord="0"]');
  const pane = document.querySelector('.oxe-pane-top').getBoundingClientRect();
  const b = el?.getBoundingClientRect();
  return b ? { top: Math.round(b.top - pane.top), pane: Math.round(pane.height) } : null;
});
check(inPane && inPane.top === 20, 'frozen: Quarters kept in place in the frozen rows', JSON.stringify(inPane));
await page.evaluate(() => { window.oxe.grid.el.scrollTop = 0; window.oxe.command('freeze', 'none'); });
await page.waitForTimeout(100);

// Picked and moved: Rectangle 5 (B11:D15) dragged a column right, then back with Undo; the arrow key moves it a row.
const at = async (k) => { const b = await page.locator(`.oxe-charts > .oxe-obj[data-oxe-ord="${k}"]`).boundingBox(); return { x: b.x + b.width / 2, y: b.y + b.height / 2 }; };
let p = await at(4);
await page.mouse.click(p.x, p.y);
check(JSON.stringify(await picked()) === '[4]', 'a click picks the rectangle', JSON.stringify(await picked()));
const before = await drawing();
await page.mouse.move(p.x, p.y);
await page.mouse.down();
await page.mouse.move(p.x + 40, p.y + 2, { steps: 4 });
await page.mouse.move(p.x + 120, p.y + 2, { steps: 4 });
await page.mouse.up();
const movedXml = await drawing();
check(fromCol(before, 'Rectangle 5') === 1 && fromCol(movedXml, 'Rectangle 5') === 2, 'dragged 120px right: from column B to C', `${fromCol(before, 'Rectangle 5')} -> ${fromCol(movedXml, 'Rectangle 5')}`);
check(movedXml.replace(anchorOf(movedXml, 'Rectangle 5'), '') === before.replace(anchorOf(before, 'Rectangle 5'), ''), 'the rest of the drawing as it was');
await page.evaluate(() => window.oxe.command('undo'));
check(await drawing() === before, 'Undo puts it back, byte for byte');
p = await at(4);
await page.mouse.click(p.x, p.y);
await page.keyboard.press('ArrowDown');
check(fromRow(await drawing(), 'Rectangle 5') === 11, 'the down arrow moves it a row', String(fromRow(await drawing(), 'Rectangle 5')));
await page.evaluate(() => window.oxe.command('undo'));

// Deleted with Delete; Ctrl+click picks two, Delete deletes both in one step.
p = await at(4);
await page.mouse.click(p.x, p.y);
await page.keyboard.press('Delete');
check(await count() === 14 && !(await drawing()).includes('Rectangle 5'), 'Delete deletes the rectangle');
await page.evaluate(() => window.oxe.command('undo'));
check(await drawing() === before && await count() === 15, 'Undo brings it back');
p = await at(4);
await page.mouse.click(p.x, p.y);
const q = await at(6);
await page.keyboard.down('Control');
await page.mouse.click(q.x, q.y);
await page.keyboard.up('Control');
check(JSON.stringify((await picked()).sort()) === '[4,6]', 'Ctrl+click adds the oval', JSON.stringify(await picked()));
await page.keyboard.press('Delete');
check(await count() === 13, 'both deleted');
await page.evaluate(() => window.oxe.command('undo'));
check(await drawing() === before, 'one Undo brings both back');

// Go To Special > Objects: every one picked; deleted; then Excel's message, none being left.
const goToObjects = async () => {
  await page.click('.oxe-rtab >> text="Home"');
  await ribbon('Find & Select', () => page.click('.oxe-pop >> text=Go To Special...'));
  await page.locator('dialog[open]').getByLabel('Objects').check();
  await page.locator('dialog[open]').getByRole('button', { name: 'OK' }).click();
};
await page.mouse.click(10, 10);
await page.evaluate(() => window.oxe.select('A1'));
await goToObjects();
check((await picked()).length === 15, 'Go To Special > Objects picks all fifteen', JSON.stringify(await picked()));
await page.keyboard.press('Delete');
check(await count() === 0, 'and Delete deletes them all');
await goToObjects();
const message = await page.locator('dialog[open]').textContent();
check(/No objects were found\./.test(message), 'none left: "No objects were found."', message);
await page.locator('dialog[open]').getByRole('button', { name: 'OK' }).click();
await page.evaluate(() => window.oxe.command('undo'));
check(await count() === 15, 'Undo brings them all back');

// Printed: the objects on the pages, over their cells; and in the PDF export, their pictures and text.
const printedObjects = await page.evaluate(() => { window.oxe.preparePrint(); return [...document.querySelectorAll('.oxe-print .oxe-page')].map((pg) => pg.querySelectorAll('.oxe-page-objclip .oxe-obj').length); });
check(printedObjects.length === 2 && printedObjects[0] >= 12 && printedObjects[1] >= 3, 'printed: the objects on both pages, those across the break on each', JSON.stringify(printedObjects));
const pdf = await page.evaluate(async () => {
  const { exportPdf } = await import('/dist/pdf.js');
  return [...await exportPdf(window.oxe, { localFonts: false })];
});
const doc = await pdfjs.getDocument({ data: new Uint8Array(pdf), disableFontFace: true, verbosity: 0 }).promise;
const first = await doc.getPage(1);
const ops = await first.getOperatorList();
const imgs = ops.fnArray.filter((f) => f === pdfjs.OPS.paintImageXObject).length;
const words = (await first.getTextContent()).items.map((x) => x.str).join(' ');
check(doc.numPages === 2 && imgs >= 3 && /Rounded/.test(words) && /Oval/.test(words) && /text box/.test(words), 'PDF: two pages, the pictures as images, the shapes\' text as text', `${doc.numPages} pages, ${imgs} images: ${words.slice(0, 160)}`);
await finish();
