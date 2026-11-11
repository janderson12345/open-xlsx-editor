import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { pivotBook } from './pivots.js';
import { readPivots } from '../src/ooxml/pivot.js';
import { PIVOT_LINE_MESSAGE } from '../src/ooxml/anchors.js';
import { shiftCells } from '../src/ooxml/rowcol.js';
import { parseRange, colIndex, colName } from '../src/ooxml/refs.js';

// test/pivot-lines-excel.json (tools/excel-compare/excel-pivot-lines.ps1): a pivot table at Pivot!C5 (rows Region, sum
// of Units) of Src!A1:E19, and where it and its source went after each step Excel did on either sheet, or its refusal.
const EXCEL = JSON.parse(fs.readFileSync('test/pivot-lines-excel.json', 'utf8'));

/** pivots.js's book for the first case, its pivot moved to C5 as Excel's was, "beside" in Pivot!A1. */
function book() {
  const b = pivotBook(0);
  const { wb, piv } = b;
  const p = readPivots(wb)[0];
  const sh = wb.sheet(piv);
  const cells = [];
  for (const [r, line] of sh.cells) for (const [c, cell] of line) if (cell.v != null) cells.push({ r, c, value: cell.v });
  wb.apply(piv, [...cells.map(({ r, c }) => ({ r, c, value: null })), ...cells.map(({ r, c, value }) => ({ r: r + 2, c: c + 2, value })), { r: 0, c: 0, value: 'beside' }]);
  wb.pkg.setText(p.path, wb.pkg.text(p.path).replace(/(<location\b[^>]*\sref=")[^"]*"/, '$1C5:D9"'));
  return b;
}
const r1c1 = (ref) => {
  const a = parseRange(ref);
  return `R${a.r0 + 1}C${a.c0 + 1}:R${a.r1 + 1}C${a.c1 + 1}`;
};

EXCEL.forEach((e) => {
  const name = `${e.sheet} ${e.do} ${e.range}${e.shift ? ` ${e.shift}` : ''}`;
  test(`pivot table, ${name}: as Excel does it`, () => {
    const { wb, src, piv } = book();
    const si = e.sheet === 'Src' ? src : piv;
    const run = () => {
      const lines = /^(\d+):(\d+)$/.exec(e.range) || /^([A-Z]+):([A-Z]+)$/.exec(e.range);
      if (!e.shift && lines) {
        const rows = /^\d/.test(e.range);
        const a = rows ? Number(lines[1]) - 1 : colIndex(lines[1]);
        const b = rows ? Number(lines[2]) - 1 : colIndex(lines[2]);
        return wb.changeLines(si, rows ? 'row' : 'col', a, (e.do === 'insert' ? 1 : -1) * (b - a + 1));
      }
      return shiftCells(wb, si, parseRange(e.range), e.do === 'insert' ? 1 : -1, e.shift === 'down' || e.shift === 'up' ? 'row' : 'col');
    };
    // (One step Excel takes its own way: a cell inserted inside the pivot, shifting down, leaves Excel's pivot with
    // an empty cell in it; here, as for any cells shifted through part of a pivot, it is refused.)
    if (e.refused || name === 'Pivot insert C6 down') {
      assert.throws(run, { message: PIVOT_LINE_MESSAGE });
      return;
    }
    run();
    const p = readPivots(wb)[0];
    const loc = p && /<location\b[^>]*\sref="([^"]+)"/.exec(wb.pkg.text(p.path))[1];
    assert.equal(loc ?? null, e.pivot, 'the pivot table');
    if (e.source) {
      const ref = /<worksheetSource\b[^>]*\sref="([^"]+)"/.exec(wb.pkg.text(p.cachePath))[1];
      assert.equal(`Src!${r1c1(ref)}`, e.source, 'its source');
    }
    const beside = [...wb.sheet(piv).cells].flatMap(([r, line]) => [...line].filter(([, cell]) => cell.v === 'beside').map(([c]) => `${colName(c)}${r + 1}`));
    assert.deepEqual(beside, e.beside ? [e.beside] : [], 'the value beside it');
  });
});
