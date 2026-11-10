// What the tests of Excel's editing steps share (test/shift.test.js, test/paste.test.js): a case's sheet built here as
// tools/excel-compare/excel-steps.ps1 builds it in Excel, and its cells read the way that script reads Excel's.
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseRef, refName } from '../src/ooxml/refs.js';
import { formatEach } from '../src/ooxml/formatops.js';
import { mergeChange } from '../src/ooxml/mergeops.js';
import { readSheet } from '../src/ooxml/sheet.js';
import { parseInput } from '../src/ooxml/input.js';
import { formatCodeOf } from '../src/ooxml/styles.js';
import { colIndex } from '../src/ooxml/refs.js';

/** A case list and Excel's answers for it, by name. */
export function readCases(cases, answers) {
  const list = JSON.parse(fs.readFileSync(cases, 'utf8').replace(/^﻿/, ''));
  const excel = JSON.parse(fs.readFileSync(answers, 'utf8').replace(/^﻿/, '')).cases;
  return list.map((c) => ({ ...c, excel: excel.find((x) => x.name === c.name) }));
}
export const area = (text) => { const [a, b = a] = text.split(':').map(parseRef); return { r0: a.r, c0: a.c, r1: b.r, c1: b.c, ar: a.r, ac: a.c }; };
/** Excel's colour number (blue, green, red from the top byte down) as the file's ARGB. */
export const argb = (n) => `FF${[n & 255, (n >> 8) & 255, (n >> 16) & 255].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase()}`;
/** The error values as COM gives them (Value2 of an error cell). */
const COM_ERRORS = { '-2146826281': '#DIV/0!', '-2146826273': '#VALUE!', '-2146826265': '#REF!', '-2146826259': '#NAME?', '-2146826252': '#NUM!', '-2146826246': '#N/A', '-2146826288': '#NULL!' };

/** The case's sheet (and the next one, and its defined names) in a new workbook: { wb, si, s2, name }. Its formulas
 *  name Sheet1, read as the sheet made here. */
export function build(setup) {
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
    if (fmt.numFmt) change.numFmt = fmt.numFmt;
    if (fmt.border) change.border = { bottom: { style: 'thin' } };
    wb.setStyles(si, formatEach(wb, si, area(fmt.range), change));
  }
  // (Merged as Excel's Merge Cells merges: the first value kept, with its format.)
  for (const m of (setup.merges || []).map(area)) {
    const ch = mergeChange(wb, si, m, 'cells');
    wb.apply(si, ch.contents);
    wb.setMerges(si, ch.merges);
    wb.setStyles(si, ch.styles);
  }
  for (const [col, width] of Object.entries(setup.widths || {})) wb.setLines(si, 'col', colIndex(col), colIndex(col), { width });
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

/** A cell as excel-steps.ps1 writes it: { f, v, fmt, bold, color, border }, nothing for an empty unformatted cell.
 *  `results`: a formula's value too. */
export function look(wb, si, r, c, { results = false } = {}) {
  const sheet = wb.sheet(si);
  const cell = sheet.cells.get(r)?.get(c);
  const out = {};
  if (!cell) return out;
  // (Each cell of an array formula reads as its formula, as Excel's Range.Formula gives it.)
  const array = sheet.arrays.find((a) => r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1);
  if (array) out.f = `=${sheet.cells.get(array.r0).get(array.c0).f}`;
  else if (cell.f != null && cell.f !== '') out.f = `=${cell.f}`;
  if ((!out.f || results) && cell.v != null) out.v = cell.v;
  const code = formatCodeOf(wb.styles, cell.s || 0);
  if (code !== 'General') out.fmt = code;
  const xf = wb.styles.cellXfs[cell.s || 0];
  if (wb.styles.fonts[xf.fontId].b) out.bold = true;
  const fill = wb.styles.fills[xf.fillId];
  if (fill.pattern === 'solid' && fill.fg?.rgb) out.color = fill.fg.rgb.toUpperCase();
  if (wb.styles.borders[xf.borderId].bottom?.style) out.border = true;
  return out;
}
/** One of Excel's cells in look's terms. */
export function theirs(one, { results = false } = {}) {
  const out = {};
  if (one.f) out.f = one.f.replace(/^=@/, '=');
  if ((!one.f || results) && one.v != null) out.v = COM_ERRORS[String(one.v)] ?? one.v;
  if (one.fmt) out.fmt = one.fmt;
  if (one.bold) out.bold = true;
  if (one.color != null) out.color = argb(one.color);
  if (one.border) out.border = true;
  return out;
}
/** Every cell of `read` that has something, as look gives it. */
export function cellsIn(wb, si, read, opts) {
  const a = area(read);
  const out = {};
  for (let r = a.r0; r <= a.r1; r++) {
    for (let c = a.c0; c <= a.c1; c++) {
      const one = look(wb, si, r, c, opts);
      if (Object.keys(one).length) out[refName(r, c)] = one;
    }
  }
  return out;
}
/** Excel's cells of a case, as cellsIn gives them. */
export const excelCells = (answer, opts) => Object.fromEntries(Object.entries(answer.cells).map(([k, v]) => [k, theirs(v, opts)]).filter(([, v]) => Object.keys(v).length));

// (STEPS_SAVE=folder: each case's workbook saved there after its steps, with a list for
// tools/excel-compare/excel-readback.ps1 to read them back in Excel.)
const saved = [];
/** Save a case's workbook for reading back in Excel, when STEPS_SAVE names a folder. */
export function saveCase(c, { wb, si, s2 }) {
  const dir = process.env.STEPS_SAVE;
  if (!dir) return;
  const file = `${dir}/${saved.length + 1}.xlsx`;
  fs.writeFileSync(file, wb.toBytes());
  saved.push({ name: c.name, file, sheet: wb.sheets[si].name, sheet2: wb.sheets[s2].name, read: c.read, cells2: Object.keys(c.setup.sheet2 || {}) });
  fs.writeFileSync(`${dir}/saved.json`, JSON.stringify(saved, null, 1));
}
