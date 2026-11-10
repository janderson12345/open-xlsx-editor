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
