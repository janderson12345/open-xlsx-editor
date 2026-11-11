import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { shiftCells } from '../src/ooxml/rowcol.js';
import { TABLE_MESSAGES } from '../src/ooxml/anchors.js';
import { readTables } from '../src/ooxml/tables.js';
import { readNotes } from '../src/ooxml/notes.js';
import { parseRange, refName, rangeName, colIndex } from '../src/ooxml/refs.js';
import { cellText } from '../src/ooxml/numfmt.js';

// test/tables-excel.json (tools/excel-compare/excel-anchors.ps1 on anchor-cases.json): the table Stock, a formula
// reading it, the chart, rectangle and notes after each step Excel did (or its refusal), on test/fixtures/anchors-table.xlsx.
const EXCEL = JSON.parse(fs.readFileSync('test/tables-excel.json', 'utf8'));
const CASES = JSON.parse(fs.readFileSync('tools/excel-compare/anchor-cases.json', 'utf8')).cases;
const EMU_PT = 12700;

function step(wb, s) {
  const lines = /^(\d+):(\d+)$/.exec(s.range) || /^([A-Z]+):([A-Z]+)$/.exec(s.range);
  if (!s.shift && lines) {
    const rows = /^\d/.test(s.range);
    const a = rows ? Number(lines[1]) - 1 : colIndex(lines[1]);
    const b = rows ? Number(lines[2]) - 1 : colIndex(lines[2]);
    return wb.changeLines(0, rows ? 'row' : 'col', a, (s.do === 'insert' ? 1 : -1) * (b - a + 1));
  }
  return shiftCells(wb, 0, parseRange(s.range), s.do === 'insert' ? 1 : -1, s.shift === 'down' || s.shift === 'up' ? 'row' : 'col');
}
/** The formula reading the table, as Excel's Find gives it: "A20 =SUM(...) = 85". */
function formula(wb) {
  const sh = wb.sheet(0);
  for (const [r, line] of sh.cells) for (const [c, cell] of line) if (cell.f && cell.f.includes('SUM(')) return `${refName(r, c)} =${cell.f} = ${cellText(cell)}`;
  return null;
}
function boxes(wb) {
  const text = wb.pkg.text('xl/drawings/drawing1.xml');
  const out = { chart: [], shape: [] };
  for (const m of text.matchAll(/<xdr:twoCellAnchor>([\s\S]*?)<\/xdr:twoCellAnchor>/g)) {
    const n = (tag, side) => Number(new RegExp(`<xdr:${side}>[\\s\\S]*?<xdr:${tag}>(-?\\d+)</xdr:${tag}>`).exec(m[1])[1]);
    const [c0, r0, c1, r1] = [n('col', 'from'), n('row', 'from'), n('col', 'to'), n('row', 'to')];
    out[m[1].includes('graphicFrame') ? 'chart' : 'shape'].push(`${refName(r0, c0)}-${refName(r1, c1)} h${(r1 - r0) * 15 + (n('rowOff', 'to') - n('rowOff', 'from')) / EMU_PT}`);
  }
  return out;
}
const excelBoxes = (list) => [].concat(list || []).map((b) => `${b.from}-${b.to} h${b.height}`);

CASES.forEach((c, k) => {
  test(`table, ${c.name}: as Excel does it`, () => {
    const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/anchors-table.xlsx')));
    const e = EXCEL[k];
    assert.equal(e.name, c.name);
    if (e.refused) {
      // (Excel's own words where COM gives them; for a header row, COM only says the Delete failed.)
      const words = /move cells in a table/.test(e.refused) ? TABLE_MESSAGES.shift : /header/.test(c.name) ? TABLE_MESSAGES.header : null;
      assert.throws(() => step(wb, c.step), words ? { message: words } : Error);
      return;
    }
    step(wb, c.step);
    const t = readTables(wb).find((x) => x.name === 'Stock');
    assert.equal(t ? rangeName(t) : null, e.table, 'the table');
    if (t) assert.deepEqual(t.columns, [].concat(e.columns), 'its columns');
    assert.equal(formula(wb), e.formula, 'the formula reading it');
    assert.deepEqual(boxes(wb).chart, excelBoxes(e.charts), 'the chart');
    assert.deepEqual(boxes(wb).shape, excelBoxes(e.shapes), 'the rectangle');
    const notes = [...readNotes(wb, 0).keys()].map((key) => { const [r, col] = key.split(',').map(Number); return refName(r, col); });
    assert.deepEqual(notes.sort(), [].concat(e.notes || []).sort(), 'the notes');
  });
});
