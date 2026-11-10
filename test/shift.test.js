import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseRef, refName, rangeName } from '../src/ooxml/refs.js';
import { formatEach } from '../src/ooxml/formatops.js';
import { mergeChange } from '../src/ooxml/mergeops.js';
import { readSheet } from '../src/ooxml/sheet.js';
import { parseInput } from '../src/ooxml/input.js';
import { shiftCells, shapeShift, ROWCOL_MESSAGES } from '../src/ooxml/rowcol.js';
import { SheetError } from '../src/ooxml/sheetops.js';

// Inserting and deleting cells, each case as Excel did it over COM (tools/excel-compare/excel-steps.ps1 on
// shift-cases.json, from shift-cases.mjs, giving test/shift-excel.json): the same sheet built here, the same step
// done, and every cell of A1:L26 (formula, value, fill, bold), the merged ranges, the next sheet's formulas and the
// defined names compared with Excel's.
const cases = JSON.parse(fs.readFileSync('tools/excel-compare/shift-cases.json', 'utf8'));
const excel = JSON.parse(fs.readFileSync('test/shift-excel.json', 'utf8').replace(/^﻿/, '')).cases;
const area = (text) => { const [a, b = a] = text.split(':').map(parseRef); return { r0: a.r, c0: a.c, r1: b.r, c1: b.c, ar: a.r, ac: a.c }; };
// (One reference Excel rewrites its own way, not copied: =SUM(B3:B3), B3 hidden in A3:B3, becomes =SUM(A3:A3) when
// cells are inserted at B3 and the merged range widens. Here it moves with B3: =SUM(D3:D3). docs/ROWCOL.md.)
const KNOWN = { 'merge across the edge, insert right': ['H12'] };
/** Excel's colour number (blue, green, red from the top byte down) as the file's ARGB. */
const argb = (n) => `FF${[n & 255, (n >> 8) & 255, (n >> 16) & 255].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase()}`;

function build(setup) {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/formulas.xlsx')));
  const si = wb.addSheet(wb.sheets.length);
  const s2 = wb.addSheet(wb.sheets.length);
  const name = wb.sheets[si].name;
  const own = (text) => text.replace(/Sheet1!/g, `${name}!`);
  const entry = (ref, text) => {
    const { r, c } = parseRef(ref);
    const p = parseInput(text);
    return p.formula != null ? { r, c, formula: p.formula } : { r, c, value: p.value };
  };
  wb.apply(si, Object.entries(setup.cells).map(([ref, text]) => entry(ref, text)));
  wb.apply(s2, Object.entries(setup.sheet2 || {}).map(([ref, text]) => entry(ref, own(text))));
  for (const [ref, text] of Object.entries(setup.arrays || {})) {
    // (An array formula, written as Excel writes one: in its first cell, with its range.)
    const a = area(ref);
    wb.apply(si, [{ r: a.r0, c: a.c0, formula: text.slice(1) }]);
    const { doc } = wb._part(si);
    const el = [...doc.getElementsByTagNameNS('*', 'c')].find((x) => x.getAttribute('r') === refName(a.r0, a.c0));
    const f = el.getElementsByTagNameNS('*', 'f')[0];
    f.setAttribute('t', 'array');
    f.setAttribute('ref', ref);
    wb._parsed.set(wb.sheets[si], { doc, model: readSheet(doc, wb.sharedStrings) });
  }
  for (const fmt of setup.formats || []) {
    const change = {};
    if (fmt.bold) change.font = { b: true };
    if (fmt.color != null) change.fill = { rgb: argb(fmt.color) };
    wb.setStyles(si, formatEach(wb, si, area(fmt.range), change));
  }
  // (Merged as Excel's Merge Cells merges: the first value kept, with its format.)
  for (const m of (setup.merges || []).map(area)) {
    const ch = mergeChange(wb, si, m, 'cells');
    wb.apply(si, ch.contents);
    wb.setMerges(si, ch.merges);
    wb.setStyles(si, ch.styles);
  }
  for (const [n, text] of Object.entries(setup.names || {})) {
    const el = wb.doc.createElementNS(wb.ns, 'definedName');
    el.setAttribute('name', n);
    el.textContent = own(text).slice(1);
    let names = wb.doc.getElementsByTagNameNS(wb.ns, 'definedNames')[0];
    if (!names) { names = wb.doc.createElementNS(wb.ns, 'definedNames'); wb.doc.documentElement.insertBefore(names, wb.doc.getElementsByTagNameNS(wb.ns, 'calcPr')[0] || null); }
    names.appendChild(el);
    wb.definedNames.push(Object.defineProperty({ name: n, sheet: null, text: el.textContent }, 'el', { value: el, writable: true }));
  }
  wb.calculate();
  return { wb, si, s2, name };
}

/** What a cell holds as excel-steps.ps1 writes it: { f, v, bold, color }, nothing for an empty unformatted cell. */
function look(wb, si, r, c) {
  const sheet = wb.sheet(si);
  const cell = sheet.cells.get(r)?.get(c);
  const out = {};
  if (!cell) return out;
  // (Each cell of an array formula reads as its formula, as Excel's Range.Formula gives it.)
  const array = sheet.arrays.find((a) => r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1);
  if (array) out.f = `=${sheet.cells.get(array.r0).get(array.c0).f}`;
  else if (cell.f != null && cell.f !== '') out.f = `=${cell.f}`;
  else if (cell.v != null) out.v = cell.v;
  const xf = wb.styles.cellXfs[cell.s || 0];
  if (wb.styles.fonts[xf.fontId].b) out.bold = true;
  const fill = wb.styles.fills[xf.fillId];
  if (fill.pattern === 'solid' && fill.fg?.rgb) out.color = fill.fg.rgb.toUpperCase();
  return out;
}
const theirs = (one) => {
  const out = {};
  if (one.f) out.f = one.f.replace(/^=@/, '=');
  else if (one.v != null) out.v = one.v;
  if (one.bold) out.bold = true;
  if (one.color != null) out.color = argb(one.color);
  return out;
};

/** The step of a case done here; its answer, as Excel's. */
function run(c, built) {
  const { wb, si } = built;
  for (const st of c.steps) {
    const sel = area(st.range);
    const axis = st.shift ? (st.shift === 'down' || st.shift === 'up' ? 'row' : 'col') : shapeShift(sel);
    shiftCells(wb, si, sel, st.do === 'insert' ? 1 : -1, axis, { unmerge: true });
  }
}

const saved = [];
function save(c, { wb, si, s2 }) {
  const file = `${process.env.SHIFT_SAVE}/${saved.length + 1}.xlsx`;
  fs.writeFileSync(file, wb.toBytes());
  saved.push({ name: c.name, file, sheet: wb.sheets[si].name, sheet2: wb.sheets[s2].name, read: c.read, cells2: Object.keys(c.setup.sheet2 || {}) });
  fs.writeFileSync(`${process.env.SHIFT_SAVE}/saved.json`, JSON.stringify(saved, null, 1));
}

for (const c of cases) {
  const want = excel.find((x) => x.name === c.name);
  test(`cells: ${c.name}, as Excel`, () => {
    const built = build(c.setup);
    const { wb, si, s2, name } = built;
    if (want.refused) {
      assert.throws(() => run(c, built), SheetError);
      return;
    }
    run(c, built);
    // (SHIFT_SAVE=folder: each workbook saved, for reading back in Excel with excel-readback.ps1.)
    if (process.env.SHIFT_SAVE) save(c, built);
    const read = area(c.read);
    const mine = {};
    for (let r = read.r0; r <= read.r1; r++) {
      for (let col = read.c0; col <= read.c1; col++) {
        const one = look(wb, si, r, col);
        if (Object.keys(one).length) mine[refName(r, col)] = one;
      }
    }
    const want2 = Object.fromEntries(Object.entries(want.cells).map(([k, v]) => [k, theirs(v)]).filter(([, v]) => Object.keys(v).length));
    for (const ref of KNOWN[c.name] || []) { delete mine[ref]; delete want2[ref]; }
    assert.deepEqual(mine, want2);
    assert.deepEqual(wb.sheet(si).merges.map(rangeName).sort(), [...(want.merges || [])].sort());
    for (const [ref, text] of Object.entries(want.sheet2 || {})) {
      const { r, c: col } = parseRef(ref);
      assert.equal(`=${wb.sheet(s2).cells.get(r)?.get(col)?.f}`, text.replace(/^=@/, '=').replace(/Sheet1!/g, `${name}!`), ref);
    }
    for (const [n, text] of Object.entries(want.names || {})) {
      if (!c.setup.names[n]) continue;
      assert.equal(wb.definedNames.find((d) => d.name === n).text, text.slice(1).replace(/Sheet1!/g, `${name}!`), n);
    }
  });
}

test('cells: Excel\'s question before a shift that splits merged cells, its refusals in its words', () => {
  const c = cases.find((x) => x.name === 'merge across the edge, insert down');
  const { wb, si } = build(c.setup);
  assert.throws(() => shiftCells(wb, si, area('B3:C4'), 1, 'row'), (e) => e.ask && e.message === ROWCOL_MESSAGES.unmerge);
  const full = build(cases.find((x) => x.name === 'filled last row, insert down').setup);
  assert.throws(() => shiftCells(full.wb, full.si, area('B3'), 1, 'row'), (e) => e.message === ROWCOL_MESSAGES.offSheet);
  const arr = build(cases.find((x) => x.name === 'array across the rows, insert right').setup);
  assert.throws(() => shiftCells(arr.wb, arr.si, area('B3:C4'), 1, 'col'), (e) => e.message === ROWCOL_MESSAGES.array);
});
