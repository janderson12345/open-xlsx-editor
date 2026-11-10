import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { rangeName, colIndex } from '../src/ooxml/refs.js';
import { snapshot, pasteCopy, pasteSpecial, pasteWidths, pasteLink } from '../src/ooxml/clipops.js';
import { transposeFormula } from '../src/calc/tokens.js';
import { readCases, area, build, cellsIn, excelCells, saveCase } from './steps.js';

// Paste Special and Paste Link, each case as Excel did it over COM (Range.Copy, then PasteSpecial or Worksheet.Paste
// with Link: tools/excel-compare/excel-steps.ps1 on paste-cases.json, from paste-cases.mjs, giving
// test/paste-excel.json): the same sheet built here, the same paste done, and every cell of A1:H16 (formula and its
// value, format, fill, bold, bottom border), the merged ranges and the column widths compared with Excel's.
const cases = readCases('tools/excel-compare/paste-cases.json', 'test/paste-excel.json');

/** The steps of a case done here, as the editor applies a paste. */
function run(c, { wb, si }) {
  for (const st of c.steps) {
    const clip = snapshot(wb, si, area(st.from));
    if (st.paste === 'widths') {
      for (const { c: col, width } of pasteWidths(area(st.to), clip)) wb.setLines(si, 'col', col, col, { width });
      continue;
    }
    const steps = st.do === 'link' ? pasteLink(wb, si, area(st.to), clip) : pasteSpecial(wb, si, area(st.to), clip, { paste: st.paste, op: st.op, skip: st.skip, transpose: st.transpose });
    for (const x of steps) {
      if (x.contents.length) wb.apply(x.si, x.contents);
      wb.setMerges(x.si, x.merges);
      if (x.styles.length) wb.setStyles(x.si, x.styles);
    }
  }
  wb.calculate();
}

for (const c of cases) {
  test(`paste special: ${c.name}, as Excel`, () => {
    const built = build(c.setup);
    const { wb, si } = built;
    run(c, built);
    saveCase(c, built);
    assert.deepEqual(cellsIn(wb, si, c.read, { results: true }), excelCells(c.excel, { results: true }));
    assert.deepEqual(wb.sheet(si).merges.map(rangeName).sort(), [...(c.excel.merges || [])].sort());
    // (Excel's widths in characters; the file's here: the same column widths are what is compared.)
    const width = (col) => wb.sheet(si).cols.find((x) => x.min <= col && x.max >= col)?.width ?? null;
    const custom = Object.keys(c.excel.widths || {}).sort();
    assert.deepEqual(['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'].filter((l) => width(colIndex(l)) != null), custom);
    for (const l of custom) assert.equal(width(colIndex(l)), width(colIndex(custom[0])));
  });
}

test('paste special: Transpose with references to the copy, outside it, $ and not, and off the sheet', () => {
  const from = { sheet: 'S', r0: 0, c0: 0, r1: 2, c1: 2 };
  const t = (f, src, dst) => transposeFormula(f, { here: 'S', from, src, dst, to: { r: 0, c: 4 } });
  assert.equal(t('$A$1+B1', { r: 1, c: 1 }, { r: 1, c: 5 }), '$E$1+E2');
  assert.equal(t('SUM(A1:B1)', { r: 2, c: 1 }, { r: 1, c: 6 }), 'SUM(E1:E2)');
  assert.equal(t('B1+D9', { r: 1, c: 1 }, { r: 1, c: 5 }), 'E2+M4');
  assert.equal(t('$D$9+1', { r: 1, c: 0 }, { r: 0, c: 5 }), '$D$9+1');
  assert.equal(t('S!A1+Other!A1', { r: 1, c: 0 }, { r: 0, c: 5 }), 'S!E1+Other!E1');
  assert.equal(t('A1', { r: 2, c: 0 }, { r: 0, c: 6 }), 'E1');
  assert.equal(transposeFormula('A1', { here: 'S', from: { sheet: 'S', r0: 5, c0: 5, r1: 6, c1: 6 }, src: { r: 6, c: 5 }, dst: { r: 0, c: 1 }, to: { r: 0, c: 0 } }), '#REF!');
});

test('paste: an error value copied stays an error, not text', () => {
  const { wb, si } = build({ cells: {} });
  wb.apply(si, [{ r: 0, c: 0, value: { code: '#DIV/0!' } }]);
  for (const x of pasteCopy(wb, si, area('C1'), snapshot(wb, si, area('A1')))) wb.apply(x.si, x.contents);
  assert.deepEqual([wb.sheet(si).cells.get(0).get(2).t, wb.sheet(si).cells.get(0).get(2).v], ['e', '#DIV/0!']);
});
