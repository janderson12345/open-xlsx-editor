import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseRef, refName } from '../src/ooxml/refs.js';
import { formatEach } from '../src/ooxml/formatops.js';
import { findAll, findNext, replaceChanges, pattern } from '../src/ooxml/findops.js';

// What Excel did on the same sheet over COM (Range.Find and FindNext, Cells.Replace; docs/FIND.md).
const range = (t) => { const p = parseRef(t); return { r0: p.r, c0: p.c, r1: p.r, c1: p.c, ar: p.r, ac: p.c }; };
/** Excel's probe sheet: A1 apple, B1 Apple pie, C1 pineapple; A2 0.5 as 0%, B2 =A2*2, C2 12; A3 =A1&"x", B3 a*b,
 *  C3 axxb; A4 =A1+A10, B4 45292 as m/d/yyyy. */
function setup() {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/formulas.xlsx')));
  const si = wb.addSheet(wb.sheets.length);
  const at = (ref, x) => ({ ...parseRef(ref), ...x });
  wb.apply(si, [at('A1', { value: 'apple' }), at('B1', { value: 'Apple pie' }), at('C1', { value: 'pineapple' }), at('A2', { value: 0.5 }),
    at('B2', { formula: 'A2*2' }), at('C2', { value: 12 }), at('A3', { formula: 'A1&"x"' }), at('B3', { value: 'a*b' }), at('C3', { value: 'axxb' }),
    at('A4', { formula: 'A1+A10' }), at('B4', { value: 45292 })]);
  wb.setStyles(si, formatEach(wb, si, range('A2'), { numFmt: '0%' }));
  wb.setStyles(si, formatEach(wb, si, range('B4'), { numFmt: 'm/d/yyyy' }));
  return { wb, si };
}
/** Excel's walk with Find and FindNext from A1: every match once, from the one after A1. */
const walk = (wb, si, what, opts = {}, start = 'A1') => {
  const out = [];
  let at = { si, ...parseRef(start) };
  for (let k = 0; k < 12; k++) {
    at = findNext(wb, at, what, opts);
    if (!at) break;
    const name = refName(at.r, at.c);
    if (out.includes(name)) break;
    out.push(name);
  }
  return out.join(',') || 'none';
};
const v = (wb, si, ref) => { const { r, c } = parseRef(ref); const cell = wb.sheet(si).cells.get(r)?.get(c); return cell?.f ? `=${cell.f}` : cell?.v ?? null; };
const run = (wb, steps) => { for (const st of steps) wb.apply(st.si, st.contents); };

test('find: from the cell after the active one, round to it; case, whole cells, wildcards', () => {
  const { wb, si } = setup();
  assert.equal(walk(wb, si, 'apple'), 'B1,C1,A1');
  assert.equal(walk(wb, si, 'apple', { byColumns: true }), 'B1,C1,A1');
  assert.equal(walk(wb, si, 'apple', { whole: true }), 'A1');
  assert.equal(walk(wb, si, 'Apple', { matchCase: true }), 'B1');
  assert.equal(walk(wb, si, 'apple', {}, 'B1'), 'C1,A1,B1');
  assert.equal(walk(wb, si, 'a*b', { whole: true }), 'B3,C3');
  assert.equal(walk(wb, si, 'a~*b', { whole: true }), 'B3');
  assert.equal(walk(wb, si, 'a?xb', { whole: true }), 'C3');
  assert.equal(walk(wb, si, 'p?e'), 'B1,C1,A1');
  assert.equal(walk(wb, si, '='), 'B2,A3,A4');
});

test('Look in: formulas as the formula bar has them, values as the cells show them', () => {
  const { wb, si } = setup();
  assert.equal(walk(wb, si, '50%'), 'A2');
  assert.equal(walk(wb, si, '50%', { lookIn: 'values' }), 'A2');
  assert.equal(walk(wb, si, '0.5'), 'none');
  assert.equal(walk(wb, si, 'A2'), 'B2');
  assert.equal(walk(wb, si, 'A2', { lookIn: 'values' }), 'none');
  assert.equal(walk(wb, si, '1', { lookIn: 'values' }), 'B2,C2,B4');
  assert.equal(walk(wb, si, '1'), 'C2,A3,A4,B4');
  assert.equal(walk(wb, si, '1/1/2024'), 'B4');
  assert.equal(walk(wb, si, '45292'), 'none');
});

test('replace: parts or whole cells, entered as if typed; formulas as plain text', () => {
  const all = (wb, si, what, opts) => findAll(wb, si, what, opts);
  let { wb, si } = setup();
  run(wb, replaceChanges(wb, all(wb, si, 'apple'), 'apple', 'pear'));
  assert.deepEqual(['A1', 'B1', 'C1', 'A3'].map((r) => v(wb, si, r)), ['pear', 'pear pie', 'pinepear', '=A1&"x"']);
  ({ wb, si } = setup());
  run(wb, replaceChanges(wb, all(wb, si, 'apple', { whole: true }), 'apple', 'pear', { whole: true }));
  assert.deepEqual(['A1', 'B1'].map((r) => v(wb, si, r)), ['pear', 'Apple pie']);
  ({ wb, si } = setup());
  run(wb, replaceChanges(wb, all(wb, si, 'A1'), 'A1', 'B9'));
  assert.deepEqual(['A3', 'A4'].map((r) => v(wb, si, r)), ['=B9&"x"', '=B9+B90']);
  ({ wb, si } = setup());
  const ch = replaceChanges(wb, all(wb, si, '1'), '1', '9');
  run(wb, ch);
  assert.deepEqual(['C2', 'A4', 'B3'].map((r) => v(wb, si, r)), [92, '=A9+A90', 'a*b']);
  assert.equal(v(wb, si, 'B4'), 45544); // (the date 1/1/2024 made 9/9/2024, its format kept, as Excel does)
  ({ wb, si } = setup());
  run(wb, replaceChanges(wb, all(wb, si, '50'), '50', '60'));
  assert.equal(v(wb, si, 'A2'), 0.6); // (50% made 60%)
  ({ wb, si } = setup());
  run(wb, replaceChanges(wb, all(wb, si, 'a*b'), 'a*b', 'Z'));
  assert.deepEqual(['B3', 'C3', 'A1', 'B1'].map((r) => v(wb, si, r)), ['Z', 'Z', 'apple', 'Apple pie']);
  ({ wb, si } = setup());
  run(wb, replaceChanges(wb, all(wb, si, 'p'), 'p', ''));
  assert.deepEqual(['A1', 'C1'].map((r) => v(wb, si, r)), ['ale', 'ineale']);
  ({ wb, si } = setup());
  run(wb, replaceChanges(wb, all(wb, si, 'apple', { whole: true }), 'apple', '', { whole: true }));
  assert.equal(v(wb, si, 'A1'), null);
  ({ wb, si } = setup());
  run(wb, replaceChanges(wb, all(wb, si, '12', { whole: true }), '12', '=1+1', { whole: true }));
  assert.equal(v(wb, si, 'C2'), '=1+1');
});

test('patterns and a whole workbook', () => {
  assert.equal(pattern(''), null);
  assert.ok(pattern('a.b').test('a.b') && !pattern('a.b').test('axb'));
  const { wb, si } = setup();
  const all = findAll(wb, si, 'apple', { within: 'workbook' });
  assert.ok(all.length >= 3 && all[0].si === si);
});

test('Look in Notes and Comments as Excel 365: notes alone, threaded comments alone, neither for Formulas', async () => {
  const { build } = await import('./steps.js');
  const { setNote } = await import('../src/ooxml/noteedit.js');
  const { addThread } = await import('../src/ooxml/threadedit.js');
  const { findAll: all } = await import('../src/ooxml/findops.js');
  const { refName } = await import('../src/ooxml/refs.js');
  const EXCEL = JSON.parse((await import('node:fs')).readFileSync('test/find2-excel.json', 'utf8'));
  // (Excel's cells, test/find2-excel.json from excel-find2.ps1: notes in A1, B2 and E5, a thread in C3, "apple" in D4.)
  const { wb, si } = build({ cells: { D4: 'apple' } });
  setNote(wb, si, 0, 0, 'Check the apple stock', '');
  setNote(wb, si, 1, 1, 'nothing here', '');
  setNote(wb, si, 4, 4, 'APPLES', '');
  addThread(wb, si, 2, 2, 'apple pie thread', 'Someone');
  const found = (lookIn) => all(wb, si, 'apple', { lookIn }).map((h) => refName(h.r, h.c));
  const of = (name) => EXCEL.find((x) => x.name === name).found;
  // (Excel's list starts after A1, so A1 comes last: the same cells here in search order.)
  assert.deepEqual(found('notes').sort(), [...of('notes: apple')].sort());
  assert.deepEqual(found('comments'), of('comments: apple'));
  assert.deepEqual(found('formulas'), of('formulas: apple, with notes about'));
});

test("Find and Replace with formats, as Excel's dialog does them (its window, find-format-excel.json)", async () => {
  const fs = await import('node:fs');
  const { build, area } = await import('./steps.js');
  const { findAll: all, replaceChanges } = await import('../src/ooxml/findops.js');
  const { refName } = await import('../src/ooxml/refs.js');
  const X = JSON.parse(fs.readFileSync('test/find-format-excel.json', 'utf8'));
  const { wb, si } = build({ cells: X.cells, formats: [...X.bold.map((a) => ({ range: a, bold: true })), ...X.filled.map((a) => ({ range: a, color: 65535 }))] });
  const opts = { findFormat: { bold: true }, replaceFormat: { fill: 'FFFFFF00' } };
  assert.deepEqual(all(wb, si, X.findAll.what, opts).map((h) => refName(h.r, h.c)), X.findAll.found, 'Find All, Bold alone');
  const run = (what, withText) => {
    const steps = replaceChanges(wb, all(wb, si, what, opts), what, withText, opts);
    for (const st of steps) { if (st.contents.length) wb.apply(st.si, st.contents); if (st.styles.length) wb.setStyles(st.si, st.styles); }
    return steps.count;
  };
  const after = () => Object.fromEntries(Object.keys(X.replaceAll.after).map((ref) => {
    const a = area(ref);
    const cell = wb.sheet(si).cells.get(a.r0)?.get(a.c0);
    const xf = wb.styles.cellXfs[cell?.s || 0];
    const fill = wb.styles.fills[xf.fillId];
    return [ref, [cell?.v, !!wb.styles.fonts[xf.fontId].b, fill?.pattern === 'solid' && fill.fg?.rgb?.toUpperCase() === 'FFFFFF00']];
  }));
  assert.equal(run(X.replaceAll.what, X.replaceAll.with), X.replaceAll.replacements, 'apple to fruit: two replaced');
  assert.deepEqual(after(), X.replaceAll.after);
  assert.equal(run(X.formatsOnly.what, X.formatsOnly.with), X.formatsOnly.replacements, 'formats only: three');
  assert.deepEqual(after(), X.formatsOnly.after);
});
