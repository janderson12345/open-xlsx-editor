import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Axis } from '../src/editor/axis.js';
import { breakAxis, paginate } from '../src/editor/print.js';
import { Workbook } from '../src/ooxml/workbook.js';

const sheetOf = (f) => Workbook.fromBytes(new Uint8Array(fs.readFileSync(`test/fixtures/${f}`))).sheet(0);

test('an axis breaks into runs that fit, each taking at least one index, hidden ones costing nothing', () => {
  const a = new Axis(10, 20, new Map([[2, 0], [5, 100]]));
  assert.deepEqual(breakAxis(a, 9, 60), [[0, 3], [4, 4], [5, 5], [6, 8], [9, 9]]);
  assert.deepEqual(breakAxis(a, 0, 60), [[0, 0]]);
});

test('page setup read from the sheet, Normal margins when it has none', () => {
  const basic = sheetOf('basic.xlsx').page;
  assert.deepEqual(basic.margins, { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 });
  assert.equal(basic.paperSize, 1);
  assert.equal(sheetOf('prefixed.xlsx').page.margins.left, 0.7);
});

test('pages down then over: large.xlsx splits its 5,000 rows down and its 20 columns across', () => {
  const sheet = sheetOf('large.xlsx');
  const { pages, paper, area } = paginate(sheet);
  assert.deepEqual(paper, { w: 816, h: 1056 });
  assert.equal(area.x, 0.7 * 96);
  // 9.5in of rows printed at 20.32px (15pt) is 44 rows a page; 7.1in of columns printed at 68.8px (64 x 1.075) is 9
  // columns: 114 down by 3 across, Excel's 342 pages.
  assert.deepEqual(pages[0], { r0: 0, r1: 43, c0: 0, c1: 8 });
  assert.deepEqual(pages[1], { r0: 44, r1: 87, c0: 0, c1: 8 });
  assert.equal(pages.length, 342);
  assert.equal(pages[pages.length / 3].c0, 9);
});

test('print area, titles and manual breaks (print-area.xlsx): pages as Excel prints them', async () => {
  const { printSetup } = await import('../src/ooxml/printarea.js');
  const { sheetGeometry } = await import('../src/editor/geometry.js');
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/print-area.xlsx')));
  const setup = printSetup(wb, 0);
  assert.deepEqual(setup, { area: { r0: 2, c0: 1, r1: 69, c1: 7 }, areas: [{ r0: 2, c0: 1, r1: 69, c1: 7 }], titles: { rows: [0, 1], cols: null } });
  const sheet = wb.sheet(0);
  assert.deepEqual(sheet.page.breaks, { rows: [29], cols: [5] });
  const { pages } = paginate(sheet, sheetGeometry(sheet, { print: true, normalFamily: 'Calibri' }), setup);
  // (Excel's PDF: B:E then F:H across, rows 3:29 then 30:70 down, rows 1:2 above each.)
  assert.deepEqual(pages, [{ r0: 2, r1: 28, c0: 1, c1: 4 }, { r0: 29, r1: 69, c0: 1, c1: 4 }, { r0: 2, r1: 28, c0: 5, c1: 7 }, { r0: 29, r1: 69, c0: 5, c1: 7 }]);
});

test('scale and fit to page: Excel picks the whole percentage that fits, rows of whole printer dots', async () => {
  const { printScale } = await import('../src/editor/print.js');
  const { sheetGeometry } = await import('../src/editor/geometry.js');
  const fit = sheetOf('print-fit.xlsx');
  const range = { r0: 0, c0: 0, r1: fit.maxRow, c1: fit.maxCol };
  assert.equal(printScale(fit, sheetGeometry(fit, { print: true, normalFamily: 'Calibri' }), range), 0.41, '24 columns on one page wide: 41%, as Excel');
  const scaled = sheetOf('print-scale.xlsx');
  assert.equal(printScale(scaled, sheetGeometry(scaled, { print: true, normalFamily: 'Calibri' }), range), 0.6);
  // (A 15pt Calibri row prints as 121 dots at 600 dpi; at 60% as 73, at 75% 91: here before the 60% zoom.)
  const g = sheetGeometry(scaled, { print: true, normalFamily: 'Calibri', scale: 0.6 });
  assert.ok(Math.abs(g.rows.size(0) * 0.6 * 6.25 - 73) < 1e-9);
  const { pages } = paginate(scaled, g, { scale: 0.6 });
  assert.equal(pages[0].r1 - pages[0].r0 + 1, 77, '77 rows on the first page, as Excel (78 would fill 5,694 of 5,700 dots)');
});

test("headers and footers: Excel's codes read into sections, fields filled in per page", async () => {
  const { parseHeaderFooter, fillFields } = await import('../src/ooxml/headerfooter.js');
  const s = parseHeaderFooter('&LQuarterly stock&CPage &P of &N&R&A');
  assert.deepEqual(s.left.map((r) => r.text), ['Quarterly stock']);
  assert.deepEqual(fillFields(s.center, { page: 2, pages: 5 }).map((r) => r.text).join(''), 'Page 2 of 5');
  assert.equal(fillFields(s.right, { page: 1, pages: 1, sheet: 'Probe' })[0].text, 'Probe');
  const f = parseHeaderFooter('Plain && simple &B&14Bold big&B &"Times New Roman,Italic"italic &P+1');
  assert.equal(f.center[0].text, 'Plain & simple ');
  assert.deepEqual([f.center[1].bold, f.center[1].size], [true, 14]);
  assert.deepEqual([f.center[3].font, f.center[3].italic], ['Times New Roman', true]);
  assert.equal(fillFields(f.center, { page: 3, pages: 9 }).at(-1).text, '4', '&P+1');
  assert.deepEqual(sheetOf('print-headers.xlsx').page.headerFooter.odd, { header: '&LQuarterly stock&CPage &P of &N&R&A', footer: '&C&F' });
});

test('header and footer pictures (&G): the VML shapes named for their sections, sizes in points, as Excel writes them', async () => {
  const { parseHeaderFooter } = await import('../src/ooxml/headerfooter.js');
  const { headerPictures } = await import('../src/ooxml/hfpictures.js');
  const { hfPictureBook } = await import('./hfpicbook.js');
  const s = parseHeaderFooter('&L&G&CTitle &G');
  assert.deepEqual(s.left.map((r) => !!r.picture), [true]);
  assert.deepEqual(s.center.map((r) => r.text ?? 'pic'), ['Title ', 'pic']);
  const wb = Workbook.fromBytes(hfPictureBook({ hf: { oddHeader: '&L&G', oddFooter: '&R&G', firstHeader: '&C&G', differentFirst: true }, pictures: [{ id: 'LH', w: 60, h: 30 }, { id: 'RF', w: 40.5, h: 20.25 }, { id: 'CHFIRST', w: 30, h: 15 }] }));
  const pics = headerPictures(wb, 0);
  assert.deepEqual(Object.keys(pics).sort(), ['CHFIRST', 'LH', 'RF']);
  assert.deepEqual([pics.RF.w, pics.RF.h], [40.5, 20.25]);
  assert.match(pics.LH.src, /^data:image\/png;base64,iVBORw0KGgo/);
  assert.equal(wb.sheet(0).page.headerFooter.first.header, '&C&G');
  assert.deepEqual(headerPictures(Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/print-headers.xlsx'))), 0), {});
});

test('header and footer pictures in sections of several lines: each line a box under the one before, as Excel prints them', async () => {
  const { pictureLines, linesOf } = await import('../src/editor/printhf.js');
  const { parseHeaderFooter } = await import('../src/ooxml/headerfooter.js');
  const em = (11 * 4) / 3;
  const at = (text, which, edge, h) => pictureLines(linesOf(parseHeaderFooter(text).left), { which, edge, h, em, sem: em, pitch: 20.32 });
  const near = (a, b, what) => assert.ok(Math.abs(a - b) <= 0.5, `${what}: ${a} against Excel's ${b}`);
  // (Excel's PDF, ..\runs\hfpic2 lines.xlsx: 60 by 30pt pictures 40.64px tall, 30 by 15pt 20.3px; margins 28.8px.)
  const left = at('&LTop line\n&G', 'header', 28.8, 40.64);
  near(left[0].base, 43.7, '"Top line"');
  near(left[1].picTop, 49.12, 'the picture under it');
  const centre = at('&L&G\nBelow', 'header', 28.8, 40.64);
  near(centre[0].picTop, 31.52, 'a picture with its line break: dropped');
  near(centre[1].base, 89.8, '"Below"');
  const foot = at('&LFoot one\nFoot &G two', 'footer', 1027.2, 20.3);
  near(foot[1].base, 1021.9, '"Foot  two"');
  near(foot[1].picTop, 1004.16, 'its picture');
  near(foot[0].base, 996.2, '"Foot one"');
  // (One line: alone, the picture on the margin; beside text, Excel's single-line places.)
  near(at('&L&G', 'footer', 1027.2, 27.36)[0].picTop, 999.84, 'a footer picture alone');
  near(at('&LTitle &G', 'header', 28.8, 40.64)[0].base, 69.5, '"Title" beside a picture');
});
