import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { invalidCells } from '../src/ooxml/validation.js';
import { serializeXml } from '../src/ooxml/xml.js';
import { refName } from '../src/ooxml/refs.js';
import { doStep } from './refmoves-steps.js';

// test/refmoves-excel.json (tools/excel-compare/excel-refmoves.ps1 on refmove-cases.mjs, cut down by
// refmove-results.mjs): Excel's sparklines, data validation rules, notes, threaded comments, the notes' boxes and
// Circle Invalid Data's ovals after each insert, delete or sort, on the sheet test/fixtures/refmoves.xlsx holds.
const EXCEL = JSON.parse(fs.readFileSync('test/refmoves-excel.json', 'utf8'));
// (Each result saved as test/out/refmoves/<case>.xlsx, for check-ooxml and for Excel to read back: excel-refmoves.ps1 -Books.)
const OUT = 'test/out/refmoves';
fs.mkdirSync(OUT, { recursive: true });
const book = () => Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/refmoves.xlsx')));
const unescape = (s) => s.replace(/&gt;/g, '>').replace(/&lt;/g, '<').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/** The workbook's state in refmove-results.mjs's terms. */
function stateOf(wb) {
  const sheet = serializeXml(wb._part(0).doc);
  const formulas = (body, tag) => [...body.matchAll(new RegExp(`<${tag}(\\d)>(?:<xm:f>)?([^<]*)`, 'g'))].map((m) => unescape(m[2]));
  const text = (path) => wb.pkg.text(path) || '';
  return {
    sparklines: [...sheet.matchAll(/<xm:f>([^<]*)<\/xm:f><xm:sqref>([^<]*)<\/xm:sqref>/g)].map((m) => `${m[2]} <- ${unescape(m[1])}`),
    rules: [...sheet.matchAll(/<dataValidation\b[^>]*\ssqref="([^"]*)"[^>]*>([\s\S]*?)<\/dataValidation>/g)].map((m) => [m[1], ...formulas(m[2], 'formula')].join(' | ')).sort(),
    x14: [...sheet.matchAll(/<x14:dataValidation\b[^>]*>([\s\S]*?)<\/x14:dataValidation>/g)].map((m) => [/<xm:sqref>([^<]*)/.exec(m[1])[1], ...formulas(m[1], 'x14:formula')].join(' | ')).sort(),
    notes: [...text('xl/comments1.xml').matchAll(/<comment ref="([^"]+)"/g)].map((m) => m[1]).sort(),
    threads: [...text('xl/threadedComments/threadedComment1.xml').matchAll(/<threadedComment ref="([^"]+)"/g)].map((m) => m[1]).sort(),
    boxes: [...text('xl/drawings/vmlDrawing1.vml').matchAll(/<x:Anchor>([^<]*)<\/x:Anchor>[\s\S]*?<x:Row>(\d+)<\/x:Row>[\s\S]*?<x:Column>(\d+)<\/x:Column>/g)].map((m) => `${m[2]},${m[3]}: ${m[1].trim()}`).sort(),
  };
}

test('the probe\'s sheet: as Excel had it before each step, its invalid cells circled as Excel circled them', () => {
  const wb = book();
  fs.writeFileSync(`${OUT}/1.xlsx`, wb.toBytes());
  const start = EXCEL[0];
  const got = stateOf(wb);
  for (const k of ['sparklines', 'rules', 'x14', 'notes', 'threads', 'boxes']) assert.deepEqual(got[k], start[k], k);
  // (E12 breaks the rule reading another sheet, kept in the Excel 2010 form the reader leaves aside: not circled here.)
  assert.deepEqual(invalidCells(wb, 0).map(({ r, c }) => refName(r, c)).sort(), start.ovals.filter((x) => x !== 'E12').sort());
});

// Where the editor knowingly differs (docs/ROWCOL.md): with B12:C12 deleted shifting left, Excel took B12 out of
// B10:B15's list rule, whose formula is absolute, though it left C12 in C10:C15's (and with C12:E12, E12 in E10:E15's);
// the editor leaves an absolute rule whole in each case.
const KNOWN = {
  'delete B12:C12 shifting left (one rule deleted, one moved)': { rules: ['B10:B15 | $H$10:$H$12', 'B12 | B12>#REF!', 'C10:C15 | 1 | $I$10', 'D10:D11 D13:D15 | D10>C10'] },
};

EXCEL.slice(1).map((x) => ({ ...x, ...KNOWN[x.name] })).forEach((excel, k) => {
  test(`${excel.name}: as Excel does it`, () => {
    const wb = book();
    doStep(wb, excel.step);
    const got = stateOf(wb);
    // (Sparklines in a group as Excel lists them, which can differ from their order in the file: compared as a set.)
    assert.deepEqual([...got.sparklines].sort(), [...excel.sparklines].sort(), 'sparklines');
    for (const key of ['rules', 'x14', 'notes', 'threads', 'boxes']) assert.deepEqual(got[key], excel[key], key);
    // And the workbook saved reads back the same.
    const bytes = wb.toBytes();
    fs.writeFileSync(`${OUT}/${k + 2}.xlsx`, bytes);
    const again = Workbook.fromBytes(bytes);
    assert.deepEqual(stateOf(again), got);
  });
});
