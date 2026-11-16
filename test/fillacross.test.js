import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { readSheet } from '../src/ooxml/sheet.js';
import { parseRef, parseRange } from '../src/ooxml/refs.js';
import { formatCodeOf } from '../src/ooxml/styles.js';
import { addHyperlink, hyperlinkAt } from '../src/ooxml/links.js';
import { setValidation } from '../src/ooxml/dvedit.js';
import { addConditional } from '../src/ooxml/cfedit.js';
import { setNote } from '../src/ooxml/noteedit.js';
import { readNotes } from '../src/ooxml/notes.js';
import { fillAcrossSheets, fillAcrossSteps } from '../src/ooxml/fillacross.js';

// (Excel 365 over COM: tools/excel-compare/excel-fillacross.ps1, FillAcrossSheets of Sheet1!A1:C4 to Sheet2 and Sheet3
// with xlFillWithAll, xlFillWithContents and xlFillWithFormats.)
const excel = JSON.parse(fs.readFileSync('test/fillacross-excel.json', 'utf8'));
const area = (t) => parseRange(t);

/** The workbook excel-fillacross.ps1 builds: Sheet1's source, Sheet2 and Sheet3 with italic text in the range. */
function build() {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const at = wb.sheets.length;
  const [s1, s2, s3] = [0, 1, 2].map((k) => wb.addSheet(at + k, `F${k + 1}`));
  const set = (si, ref, v) => { const { r, c } = parseRef(ref); wb.apply(si, [typeof v === 'string' && v.startsWith('=') ? { r, c, formula: v.slice(1) } : { r, c, value: v }]); };
  const style = (si, ref, change) => { const { r, c } = parseRef(ref); wb.setStyles(si, [{ r, c, s: wb.styleWith(wb.sheet(si).cells.get(r)?.get(c)?.s || 0, change) }]); };
  set(s1, 'A1', 1234.5);
  style(s1, 'A1', { numFmt: '#,##0.00', font: { b: true }, fill: { pattern: 'solid', fg: { rgb: 'FFFFFF00' } }, border: { bottom: { style: 'thin' } } });
  set(s1, 'B1', '=A1*2');
  set(s1, 'C1', '=$A$1+F2!A1');
  style(s1, 'C1', { numFmt: '#,##0.00' });
  set(s1, 'A2', 'text');
  set(s1, 'C2', '=SUM(A1:B1)');
  style(s1, 'C2', { numFmt: '#,##0.00' });
  set(s1, 'A3', 'noted');
  setNote(wb, s1, 2, 0, 'a note');
  addHyperlink(wb, s1, 2, 1, { address: 'https://example.com/', text: 'link' });
  set(s1, 'C3', 5);
  set(s1, 'A4', 7);
  addConditional(wb, s1, [area('A4')], { kind: 'cellIs', operator: 'equal', f1: '0', preset: 'lightRed' });
  const { doc } = wb._part(s1);
  setValidation(doc, [area('C3')], { type: 'whole', op: 'between', f1: '1', f2: '10' });
  wb._parsed.set(wb.sheets[s1], { doc, model: readSheet(doc, wb.sharedStrings) });
  set(s1, 'B4', 'merged');
  wb.setMerges(s1, [area('B4:C4')]);
  for (const k of [s2, s3]) {
    for (const ref of ['A1', 'B2', 'C3', 'B4']) { set(k, ref, `old${k - s1 + 1}`); style(k, ref, { font: { i: true } }); }
    set(k, 'D1', `outside${k - s1 + 1}`);
    set(k, 'A5', `below${k - s1 + 1}`);
  }
  wb.calculate();
  return { wb, sheets: [s1, s2, s3] };
}

/** Cell `ref` of sheet si read as excel-fillacross.ps1 reads Excel's. */
function read(wb, si, ref) {
  const { r, c } = parseRef(ref);
  const cell = wb.sheet(si).cells.get(r)?.get(c);
  const xf = wb.styles.cellXfs[cell?.s || 0];
  const font = wb.styles.fonts[xf.fontId];
  const doc = wb._part(si).doc;
  const covers = (tag, attr) => [...doc.getElementsByTagNameNS('*', tag)].some((el) => (el.getAttribute(attr) || '').split(' ').some((t) => { const a = parseRange(t); return a && r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1; }));
  const value = cell?.f ? `=${cell.f}` : cell?.v ?? '';
  return {
    formula: String(value).replace(/F(\d)!/g, 'Sheet$1!'),
    numFmt: formatCodeOf(wb.styles, cell?.s || 0),
    bold: !!font?.b, italic: !!font?.i,
    fill: wb.styles.fills[xf.fillId]?.pattern === 'solid',
    bottom: !!wb.styles.borders[xf.borderId]?.bottom?.style,
    merged: wb.sheet(si).merges.some((m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1),
    note: readNotes(wb, si).has(`${r},${c}`),
    links: hyperlinkAt(wb, si, r, c) ? 1 : 0,
    cf: covers('conditionalFormatting', 'sqref') ? 1 : 0,
    dv: covers('dataValidation', 'sqref'),
  };
}
const excelOf = (x) => ({
  formula: x.formula, numFmt: x.numFmt, bold: x.bold, italic: x.italic, fill: x.fill !== -4142, bottom: x.bottom !== -4142,
  merged: x.merged, note: x.note, links: x.links, cf: x.cf, dv: x.dv > 0,
});

for (const type of ['all', 'contents', 'formats']) {
  test(`Fill Across Worksheets, ${type}: each grouped sheet's cells as Excel's FillAcrossSheets left them`, () => {
    const { wb, sheets } = build();
    fillAcrossSheets(wb, sheets[0], area('A1:C4'), sheets, type);
    let same = 0;
    let all = 0;
    for (const [k, name] of [[1, 'Sheet2'], [2, 'Sheet3'], [0, 'Sheet1']]) {
      for (const [ref, x] of Object.entries(excel[type][name])) {
        if (ref.startsWith('_')) continue;
        all++;
        const ours = read(wb, sheets[k], ref);
        assert.deepEqual(ours, excelOf(x), `${type} ${name}!${ref}`);
        same++;
      }
    }
    assert.equal(same, all);
  });
}

test('Fill Across Worksheets: the shown sheet itself and chart sheets are passed over; a whole column is cut to the cells filled', () => {
  const { wb, sheets } = build();
  const steps = fillAcrossSteps(wb, sheets[0], { r0: 0, c0: 0, r1: 1048575, c1: 0 }, [sheets[0], sheets[1]], 'contents');
  assert.equal(steps.length, 1);
  assert.equal(steps[0].si, sheets[1]);
  assert.ok(steps[0].area.r1 < 10);
  assert.throws(() => fillAcrossSteps(wb, sheets[0], area('A1'), sheets, 'values'));
});
