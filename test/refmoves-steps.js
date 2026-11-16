// For the tests of what moves with rows, columns and cells (refmoves.test.js, slicers.test.js): one of the Excel
// probes' steps (refmove-cases.mjs, slicer-cases.mjs) done the editor's way.
import assert from 'node:assert/strict';
import { shiftCells } from '../src/ooxml/rowcol.js';
import { sortChanges } from '../src/ooxml/sortops.js';
import { parseRange, colIndex } from '../src/ooxml/refs.js';

/** Do one of the probe's steps (insert, delete, sort) the editor's way. */
export function doStep(wb, s) {
  const si = s.sheet === 'Data' ? 1 : 0;
  if (s.do === 'none') return;
  const dir = s.do === 'insert' ? 1 : -1;
  const rows = /^(\d+):(\d+)$/.exec(s.range);
  const cols = /^([A-Z]+):([A-Z]+)$/.exec(s.range);
  if (rows) return wb.changeLines(si, 'row', Number(rows[1]) - 1, dir * (Number(rows[2]) - Number(rows[1]) + 1));
  if (cols) return wb.changeLines(si, 'col', colIndex(cols[1]), dir * (colIndex(cols[2]) - colIndex(cols[1]) + 1));
  const area = parseRange(s.range);
  if (s.do === 'sort') {
    const key = parseRange(s.key);
    const step = sortChanges(wb, si, area, { keys: [{ at: s.leftToRight ? key.r0 : key.c0, desc: !!s.desc }], leftToRight: !!s.leftToRight });
    assert.ok(!step.error, step.error);
    wb.apply(si, step.contents);
    wb.setStyles(si, step.styles);
    return;
  }
  shiftCells(wb, si, { ...area, ar: area.r0, ac: area.c0 }, dir, s.shift === 'down' || s.shift === 'up' ? 'row' : 'col');
}
