import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { Workbook } from '../src/ooxml/workbook.js';

// What Excel wrote for the same steps (its ribbon's Freeze Panes, Freeze Top Row, Freeze First Column over
// CommandBars.ExecuteMso; Window.Zoom; docs/VIEW.md).
const fresh = () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/formulas.xlsx')));
  return { wb, si: wb.addSheet(wb.sheets.length) };
};
const views = (wb, si) => {
  const xml = strFromU8(unzipSync(wb.toBytes())[wb.sheets[si].path]);
  return xml.match(/<sheetView\b.*<\/sheetViews>/)[0].replace(/^<sheetView [^>]*>/, '').replace(/<\/sheetView><\/sheetViews>$/, '');
};

test('freezing at C5, the top row and the first column, as Excel writes them', () => {
  let { wb, si } = fresh();
  wb.setView(si, { freeze: { rows: 4, cols: 2 }, active: { r: 4, c: 2 } });
  assert.equal(views(wb, si), '<pane xSplit="2" ySplit="4" topLeftCell="C5" activePane="bottomRight" state="frozen"/><selection pane="topRight" activeCell="C1" sqref="C1"/><selection pane="bottomLeft" activeCell="A5" sqref="A5"/><selection pane="bottomRight" activeCell="C5" sqref="C5"/>');
  assert.deepEqual(Workbook.fromBytes(wb.toBytes()).sheet(si).freeze, { rows: 4, cols: 2 });
  ({ wb, si } = fresh());
  wb.setView(si, { freeze: { rows: 1, cols: 0 }, active: { r: 4, c: 2 } });
  assert.equal(views(wb, si), '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="C5" sqref="C5"/>');
  ({ wb, si } = fresh());
  wb.setView(si, { freeze: { rows: 0, cols: 1 }, active: { r: 4, c: 2 } });
  assert.equal(views(wb, si), '<pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/><selection pane="topRight" activeCell="C5" sqref="C5"/>');
  wb.setView(si, { freeze: { rows: 0, cols: 0 }, active: { r: 3, c: 2 } });
  assert.equal(views(wb, si), '<selection activeCell="C4" sqref="C4"/>');
});

test('zoom as zoomScale and zoomScaleNormal; 100 writes neither', () => {
  const { wb, si } = fresh();
  wb.setView(si, { zoom: 150 });
  assert.match(strFromU8(unzipSync(wb.toBytes())[wb.sheets[si].path]), /<sheetView [^>]*zoomScale="150" zoomScaleNormal="150"/);
  assert.equal(Workbook.fromBytes(wb.toBytes()).sheet(si).zoom, 150);
  wb.setView(si, { zoom: 5 });
  assert.equal(wb.sheet(si).zoom, 10);
  wb.setView(si, { zoom: 100 });
  assert.doesNotMatch(strFromU8(unzipSync(wb.toBytes())[wb.sheets[si].path]), /zoomScale/);
});

test('Show Formulas as showFormulas="1" on the sheet view, as Excel writes it; off takes it away; read back', () => {
  const { wb, si } = fresh();
  const viewTag = () => strFromU8(unzipSync(wb.toBytes())[wb.sheets[si].path]).match(/<sheetView\b[^>]*>/)[0];
  wb.setView(si, { showFormulas: true });
  assert.match(viewTag(), / showFormulas="1"/);
  assert.equal(Workbook.fromBytes(wb.toBytes()).sheet(si).showFormulas, true);
  wb.setView(si, { showFormulas: false });
  assert.doesNotMatch(viewTag(), /showFormulas/);
  assert.equal(wb.sheet(si).showFormulas, false);
});

test('View > Gridlines and Headings off as showGridLines="0" and showRowColHeaders="0", as Excel writes them; on again takes them away', () => {
  // (Excel, Window.DisplayGridlines and DisplayHeadings set false then true, saved: <sheetView showGridLines="0"
  // showRowColHeaders="0" ...>, then neither.)
  const { wb, si } = fresh();
  const viewTag = () => strFromU8(unzipSync(wb.toBytes())[wb.sheets[si].path]).match(/<sheetView\b[^>]*>/)[0];
  wb.setView(si, { showGridLines: false });
  assert.match(viewTag(), / showGridLines="0"/);
  assert.doesNotMatch(viewTag(), /showRowColHeaders/);
  wb.setView(si, { showHeadings: false });
  assert.match(viewTag(), / showRowColHeaders="0"/);
  const back = Workbook.fromBytes(wb.toBytes()).sheet(si);
  assert.deepEqual([back.showGridLines, back.showHeadings], [false, false]);
  wb.setView(si, { showGridLines: true, showHeadings: true });
  assert.doesNotMatch(viewTag(), /showGridLines|showRowColHeaders/);
  assert.deepEqual([wb.sheet(si).showGridLines, wb.sheet(si).showHeadings], [true, true]);
});

test('Page Layout > Sheet Options, Print: <printOptions headings="1" gridLines="1"/> before pageMargins, gone when neither, as Excel writes it', async () => {
  // (Excel, PageSetup.PrintGridlines and PrintHeadings set true then false, saved: the element, then none.)
  const { setPrintOptions } = await import('../src/ooxml/printopts.js');
  const { wb, si } = fresh();
  const xml = () => strFromU8(unzipSync(wb.toBytes())[wb.sheets[si].path]);
  setPrintOptions(wb, si, { gridLines: true });
  assert.match(xml(), /<printOptions gridLines="1"\/><pageMargins /);
  setPrintOptions(wb, si, { headings: true });
  assert.match(xml(), /<printOptions gridLines="1" headings="1"\/><pageMargins /);
  const back = Workbook.fromBytes(wb.toBytes()).sheet(si).page;
  assert.deepEqual([back.gridLines, back.headings], [true, true]);
  setPrintOptions(wb, si, { gridLines: false, headings: false });
  assert.doesNotMatch(xml(), /printOptions/);
  assert.deepEqual([wb.sheet(si).page.gridLines, wb.sheet(si).page.headings], [false, false]);
});
