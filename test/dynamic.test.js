import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { entryChanges } from '../src/ooxml/entry.js';
import { parseRef, refName, rangeName } from '../src/ooxml/refs.js';
import { displayFormula, normalizeFormula } from '../src/calc/fileform.js';
import { editText } from '../src/ooxml/input.js';

// test/dynamic-excel.json (tools/excel-compare/excel-dynamic.ps1 on dynamic-cases.mjs): each formula typed in D1 of a
// new sheet in Excel 365 (Formula2), A1:A6 numbers and B1:B6 words: its spill range, the cells' values and texts, the
// formula as Excel shows it, and the cells' XML as Excel saved them.
const EXCEL = JSON.parse(fs.readFileSync('test/dynamic-excel.json', 'utf8'));
const CASES = JSON.parse(fs.readFileSync('tools/excel-compare/dynamic-cases.json', 'utf8')).cases;
const NUMS = [5, 3, 8, 3, 1, 8];
const WORDS = ['pear', 'apple', 'fig', 'apple', 'kiwi', 'plum'];

/** A new sheet with the probe's data, `typed` entered at refs as a person types them. */
function sheet(typed) {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'Dyn');
  wb.apply(si, NUMS.flatMap((n, r) => [{ r, c: 0, value: n }, { r, c: 1, value: WORDS[r] }]));
  const type = (ref, text) => { const { r, c } = parseRef(ref); wb.apply(si, [entryChanges(wb, si, r, c, text).content]); };
  for (const [ref, text] of typed) type(ref, text);
  return { wb, si, type };
}
const shown = (wb, si, ref) => {
  const { r, c } = parseRef(ref);
  const cell = wb.sheet(si).cells.get(r)?.get(c);
  return cell?.v == null ? null : cell.t === 'e' ? cell.v : cell.v;
};

CASES.forEach((k, n) => {
  test(`dynamic arrays as Excel 365: ${k.name}`, () => {
    const want = EXCEL[n];
    assert.equal(want.name, k.name);
    const typed = [];
    if (k.block) typed.push([k.block.ref, k.block.value]);
    typed.push(['D1', k.formula]);
    if (k.then) typed.push([k.then.ref, k.then.formula]);
    const { wb, si } = sheet(typed);
    const anchor = wb.sheet(si).cells.get(0).get(3);
    // The spill range (Excel's SpillingToRange; none for one value or a blocked spill).
    const spill = anchor.spill && (anchor.spill.r1 > 0 || anchor.spill.c1 > 3) && anchor.v !== '#SPILL!' ? rangeName(anchor.spill) : null;
    assert.equal(spill, want.spill, 'spill');
    // Every cell Excel had something in, D1:H8: the same value (errors by their text), and nothing else filled.
    for (const x of want.cells) {
      const v = shown(wb, si, x.ref);
      if (typeof x.value === 'number' && !/^#/.test(x.text)) assert.ok(Math.abs(v - x.value) < 1e-9, `${x.ref} ${v} ${x.value}`);
      else assert.equal(v, /^#/.test(x.text) ? x.text : x.value, x.ref);
      if (x.formula2?.startsWith('=')) assert.equal(editText(wb.sheet(si).cells.get(parseRef(x.ref).r).get(parseRef(x.ref).c)), x.formula2, `${x.ref} formula`);
    }
    for (let r = 0; r < 8; r++) for (let c = 3; c < 8; c++) {
      const ref = refName(r, c);
      if (!want.cells.some((x) => x.ref === ref)) assert.equal(shown(wb, si, ref), null, `${ref} empty`);
    }
    // The anchor as Excel saves it: a dynamic array formula (cm, t="array", its ref) when Excel's is, the same text.
    const ours = new TextDecoder().decode(wb.toBytes()) && wb.pkg.text(wb.sheets[si].path);
    const theirs = /<c r="D1"([^>]*)><f([^>]*)>([^<]*)<\/f>/.exec(want.xml);
    const mine = /<c r="D1"([^>]*)><f([^>]*)>([^<]*)<\/f>/.exec(ours);
    assert.equal(mine[3], theirs[3], 'the formula as the file has it');
    assert.equal(/cm="/.test(mine[1]), /cm="/.test(theirs[1]) || /LAMBDA/.test(k.formula) ? /cm="/.test(mine[1]) : false, 'cm');
    if (!/LAMBDA|blocked|nothing\)?$/.test(k.name)) assert.equal(/ref="([^"]+)"/.exec(mine[2])?.[1] ?? null, /ref="([^"]+)"/.exec(theirs[2])?.[1] ?? null, 'ref');
  });
});

test('a spill follows its source: grows, shrinks, is blocked by a typed cell and spills again when it is cleared', () => {
  const { wb, si, type } = sheet([['D1', '=SORT(A1:A6)'], ['F1', '=SUM(D1#)']]);
  const col = () => [0, 1, 2, 3, 4, 5, 6].map((r) => shown(wb, si, refName(r, 3)));
  assert.deepEqual(col(), [1, 3, 3, 5, 8, 8, null]);
  assert.equal(shown(wb, si, 'F1'), 28);
  type('A7', '2'); // (outside the source: nothing changes)
  type('D1', '=SORT(A1:A7)');
  assert.deepEqual(col(), [1, 2, 3, 3, 5, 8, 8]);
  assert.equal(shown(wb, si, 'F1'), 30);
  type('D3', 'x');
  assert.equal(shown(wb, si, 'D1'), '#SPILL!');
  assert.deepEqual(col().slice(1), [null, 'x', null, null, null, null]);
  type('D3', '');
  assert.deepEqual(col(), [1, 2, 3, 3, 5, 8, 8]);
  type('D1', '=FILTER(A1:A7,A1:A7>4)');
  assert.deepEqual(col(), [5, 8, 8, null, null, null, null]);
  assert.equal(shown(wb, si, 'F1'), 21);
  type('D1', '');
  assert.deepEqual(col(), [null, null, null, null, null, null, null]);
  assert.equal(shown(wb, si, 'F1'), '#REF!');
});

test('saved and opened again: the spill read back from the file, metadata.xml as Excel writes it', () => {
  const { wb } = sheet([['D1', '=SEQUENCE(3)']]);
  const again = Workbook.fromBytes(wb.toBytes());
  const si = again.sheetIndexOf('Dyn');
  assert.deepEqual(again.sheet(si).cells.get(0).get(3).spill, { r0: 0, c0: 3, r1: 2, c1: 3 });
  assert.match(again.pkg.text('xl/metadata.xml'), /<metadataType name="XLDAPR" minSupportedVersion="120000"[^>]*cellMeta="1"\/>.*<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"\/>.*<cellMetadata count="1"><bk><rc t="1" v="0"\/><\/bk><\/cellMetadata>/);
  assert.match(again.pkg.text('[Content_Types].xml'), /PartName="\/xl\/metadata.xml" ContentType="application\/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata\+xml"/);
  // (Typed in the opened copy, the spill grows from what the file said it filled.)
  again.apply(si, [{ r: 0, c: 3, formula: '_xlfn.SEQUENCE(5)' }]);
  assert.deepEqual([0, 1, 2, 3, 4].map((r) => again.sheet(si).cells.get(r).get(3).v), [1, 2, 3, 4, 5]);
});

test('formulas as typed and as the file has them: D1#, @, LET and LAMBDA names, SORT and FILTER prefixes', () => {
  assert.equal(normalizeFormula('SUM(d1#)'), 'SUM(_xlfn.ANCHORARRAY(D1))');
  assert.equal(normalizeFormula('@A1:A3'), 'A1:A3'); // (all @: saved as the old kind of formula, as Excel saves it)
  assert.equal(normalizeFormula('A2+@A1:A3'), 'A2+_xlfn.SINGLE(A1:A3)');
  assert.equal(normalizeFormula('let(x,A1,y,A2,x+y)'), '_xlfn.LET(_xlpm.x,A1,_xlpm.y,A2,_xlpm.x+_xlpm.y)');
  assert.equal(normalizeFormula('LAMBDA(x,x*2)(A1)'), '_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(A1)');
  assert.equal(normalizeFormula('sort(B1:B6,,-1)'), '_xlfn._xlws.SORT(B1:B6,,-1)');
  assert.equal(normalizeFormula('LET(a,SUM(A1,A2),a*2)'), '_xlfn.LET(_xlpm.a,SUM(A1,A2),_xlpm.a*2)');
  for (const f of ['SUM(D1#)', 'A2+@A1:A3', 'LET(x,A1,y,A2,x+y)', 'LAMBDA(x,x*2)(A1)', 'SORT(B1:B6,,-1)']) assert.equal(displayFormula(normalizeFormula(f)), f);
});

test('trim references as Excel 365 saves them (measured): _TRO_ in the file, the dots shown', () => {
  const cases = [
    ['SUM(A1:.E8)', 'SUM(_xlfn._TRO_TRAILING(A1:E8))'],
    ['A1.:E8', '_xlfn._TRO_LEADING(A1:E8)'],
    ['A1.:.E8+"a:.b"', '_xlfn._TRO_ALL(A1:E8)+"a:.b"'],
    ['ROWS(B:.B)', 'ROWS(_xlfn._TRO_TRAILING(B:B))'],
    ['ROWS(3:.6)', 'ROWS(_xlfn._TRO_TRAILING(3:6))'],
    ["SUM('My Data'!A1:.B4)", "SUM(_xlfn._TRO_TRAILING('My Data'!A1:B4))"],
  ];
  for (const [typed, file] of cases) {
    assert.equal(normalizeFormula(typed), file);
    assert.equal(displayFormula(file), typed);
  }
  assert.equal(normalizeFormula('trimrange(a1:e8)'), '_xlfn.TRIMRANGE(A1:E8)');
});

test('a defined LAMBDA called in another case is written as the name is, as Excel writes it (measured)', () => {
  // (Excel: =twice(5) and =TWICE(5)+sum(1) are Twice(5) and Twice(5)+SUM(1); both work out 10 and 11.)
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/formulas.xlsx')));
  const si = wb.sheetIndexOf('Calc');
  const typed = (text) => entryChanges(wb, si, 3000, 1, text).content.formula;
  assert.deepEqual([typed('=twice(5)'), typed('=TWICE(5)+sum(1)'), typed('=hyp(3,4)')], ['Twice(5)', 'Twice(5)+SUM(1)', 'Hyp(3,4)']);
  wb.apply(si, [entryChanges(wb, si, 3000, 1, '=twice(5)').content]);
  assert.equal(wb.sheet(si).cells.get(3000).get(1).v, 10);
});

test("a LAMBDA kept in a LET name and called is written with _xlpm. there too, as Excel's file has it", () => {
  // (Excel saved =LET(f,LAMBDA(x,x*2),f(3)) as _xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2),_xlpm.f(3)); this
  // editor wrote F(3), which Excel reads as a function it doesn't know.)
  const file = '_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2),_xlpm.f(3))';
  assert.equal(normalizeFormula('LET(f,LAMBDA(x,x*2),f(3))'), file);
  assert.equal(displayFormula(file), 'LET(f,LAMBDA(x,x*2),f(3))');
  assert.equal(normalizeFormula('LAMBDA(g,g(4))(LAMBDA(y,y+1))'), '_xlfn.LAMBDA(_xlpm.g,_xlpm.g(4))(_xlfn.LAMBDA(_xlpm.y,_xlpm.y+1))');
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'L');
  wb.setFormula(si, 'A1', file);
  assert.equal(wb.sheet(si).cells.get(0).get(0).v, 6);
});
