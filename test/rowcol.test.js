import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { spawnSync } from 'node:child_process';
import { strFromU8, unzipSync } from 'fflate';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseRef, rangeName } from '../src/ooxml/refs.js';
import { formatEach } from '../src/ooxml/formatops.js';
import { shiftRefs } from '../src/calc/tokens.js';
import { ROWCOL_MESSAGES } from '../src/ooxml/rowcol.js';
import { parseXml } from '../src/ooxml/xml.js';

// What Excel did with the same sheet over COM (Rows(n).Insert, Range("2:4").Delete, Columns(n).Insert...; docs/ROWCOL.md).
const range = (text) => { const [a, b = a] = text.split(':').map(parseRef); return { r0: a.r, c0: a.c, r1: b.r, c1: b.c, ar: a.r, ac: a.c }; };
const dir = 'test/out/rowcol';
fs.mkdirSync(dir, { recursive: true });

/** Excel's probe: A1:A5 1 to 5; C1 =SUM(A1:A5), C2 =A5, C3 =A3, C4 =SUM(A3:A3), C5 =$A$4, C6 =SUM(A6:A9); the next
 *  sheet's A1 reading A4; names Four ($A$4) and Third ($A$3); E2:E4 merged; row 2 bold and 30pt; B2 filled. */
function setup() {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/formulas.xlsx')));
  const si = wb.addSheet(wb.sheets.length);
  const s2 = wb.addSheet(wb.sheets.length);
  const name = wb.sheets[si].name;
  wb.apply(si, [1, 2, 3, 4, 5].map((v, r) => ({ r, c: 0, value: v })).concat(
    ['SUM(A1:A5)', 'A5', 'A3', 'SUM(A3:A3)', '$A$4', 'SUM(A6:A9)'].map((formula, r) => ({ r, c: 2, formula }))));
  wb.apply(s2, [{ r: 0, c: 0, formula: `${name}!A4` }]);
  wb.setMerges(si, [range('E2:E4')]);
  wb.setStyles(si, formatEach(wb, si, range('A2:C2'), { font: { b: true } }));
  wb.setStyles(si, formatEach(wb, si, range('B2'), { fill: { rgb: 'FFFFFF00' } }));
  // (Row 2 formatted and 30pt high as a row, as Excel's Rows(2) is.)
  const row = [...wb._part(si).doc.getElementsByTagNameNS('*', 'row')].find((el) => el.getAttribute('r') === '2');
  row.setAttribute('s', String(wb.sheet(si).cells.get(1).get(0).s));
  row.setAttribute('customFormat', '1');
  row.setAttribute('ht', '30');
  row.setAttribute('customHeight', '1');
  for (const [n, ref] of [['Four', '$A$4'], ['Third', '$A$3']]) {
    const el = wb.doc.createElementNS(wb.ns, 'definedName');
    el.setAttribute('name', n);
    el.textContent = `${name}!${ref}`;
    let names = wb.doc.getElementsByTagNameNS(wb.ns, 'definedNames')[0];
    if (!names) { names = wb.doc.createElementNS(wb.ns, 'definedNames'); wb.doc.documentElement.insertBefore(names, wb.doc.getElementsByTagNameNS(wb.ns, 'calcPr')[0] || null); }
    names.appendChild(el);
    wb.definedNames.push(Object.defineProperty({ name: n, sheet: null, text: el.textContent }, 'el', { value: el, writable: true }));
  }
  wb.calculate();
  return { wb, si, s2, name };
}
const f = (wb, si, ref) => { const { r, c } = parseRef(ref); const cell = wb.sheet(si).cells.get(r)?.get(c); return cell?.f != null ? `=${cell.f}` : String(cell?.v ?? ''); };
const col = (wb, si) => ['C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7'].map((r) => f(wb, si, r)).join(' ');
const names = (wb) => wb.definedNames.filter((d) => d.name === 'Four' || d.name === 'Third').map((d) => `${d.name}=${d.text}`).join(' ');
const merges = (wb, si) => wb.sheet(si).merges.map(rangeName).join();

test('insert a row at 3: references below move, ranges across it grow, the row above\'s format copied', () => {
  const { wb, si, s2, name } = setup();
  wb.changeLines(si, 'row', 2, 1);
  assert.equal(col(wb, si), '=SUM(A1:A6) =A6  =A4 =SUM(A4:A4) =$A$5 =SUM(A7:A10)');
  assert.equal(f(wb, s2, 'A1'), `=${name}!A5`);
  assert.equal(names(wb), `Four=${name}!$A$5 Third=${name}!$A$4`);
  assert.equal(merges(wb, si), 'E2:E5');
  assert.deepEqual([f(wb, si, 'A3'), f(wb, si, 'A4')], ['', '3']);
  assert.equal(wb.sheet(si).rows.get(2).ht, 30);
  const s = wb.sheet(si).cells.get(2).get(1).s;
  assert.equal(wb.styles.fills[wb.styles.cellXfs[s].fillId].fg.rgb, 'FFFFFF00');
  assert.equal(wb.sheet(si).cells.get(0).get(2).v, 15);
});

test('insert a row at 1 and just below a range', () => {
  let { wb, si } = setup();
  wb.changeLines(si, 'row', 0, 1);
  assert.equal(col(wb, si), ' =SUM(A2:A6) =A6 =A4 =SUM(A4:A4) =$A$5 =SUM(A7:A10)');
  ({ wb, si } = setup());
  wb.changeLines(si, 'row', 5, 1);
  assert.equal(col(wb, si), '=SUM(A1:A5) =A5 =A3 =SUM(A3:A3) =$A$4  =SUM(A7:A10)');
});

test('delete rows: ranges shrink, references to deleted cells become #REF!', () => {
  let { wb, si, s2, name } = setup();
  wb.changeLines(si, 'row', 2, -1);
  assert.equal(col(wb, si), '=SUM(A1:A4) =A4 =SUM(#REF!) =$A$3 =SUM(A5:A8)  ');
  assert.equal(f(wb, s2, 'A1'), `=${name}!A3`);
  assert.equal(names(wb), `Four=${name}!$A$3 Third=${name}!#REF!`);
  assert.equal(merges(wb, si), 'E2:E3');
  ({ wb, si, s2, name } = setup());
  wb.changeLines(si, 'row', 1, -3);
  assert.equal(col(wb, si), '=SUM(A1:A2) =#REF! =SUM(A3:A6)    ');
  assert.equal(f(wb, s2, 'A1'), `=${name}!#REF!`);
  assert.equal(merges(wb, si), '');
  ({ wb, si } = setup());
  wb.changeLines(si, 'row', 0, -5);
  assert.equal(f(wb, si, 'C1'), '=SUM(A1:A4)');
});

test('columns: inserted ones take the left column\'s width and format; deleted ones make #REF!', () => {
  let { wb, si, s2, name } = setup();
  const doc = wb._part(si).doc;
  const cols = doc.createElementNS(wb.ns, 'cols');
  const c = doc.createElementNS(wb.ns, 'col');
  for (const [k, v] of [['min', '2'], ['max', '2'], ['width', '20.7109375'], ['customWidth', '1']]) c.setAttribute(k, v);
  cols.appendChild(c);
  doc.documentElement.insertBefore(cols, doc.getElementsByTagNameNS(wb.ns, 'sheetData')[0]);
  wb.changeLines(si, 'col', 2, 1);
  assert.equal(f(wb, si, 'D1'), '=SUM(A1:A5)');
  assert.equal(wb.sheet(si).cols.find((x) => x.min <= 2 && x.max >= 2)?.width, 20.7109375);
  const s = wb.sheet(si).cells.get(1).get(2).s;
  assert.equal(wb.styles.fills[wb.styles.cellXfs[s].fillId].fg.rgb, 'FFFFFF00');
  ({ wb, si } = setup());
  wb.changeLines(si, 'col', 1, 1);
  assert.deepEqual([f(wb, si, 'D1'), f(wb, si, 'D2'), merges(wb, si)], ['=SUM(A1:A5)', '=A5', 'F2:F4']);
  ({ wb, si, s2, name } = setup());
  wb.changeLines(si, 'col', 0, -1);
  assert.deepEqual([f(wb, si, 'B1'), f(wb, si, 'B2'), f(wb, s2, 'A1')], ['=SUM(#REF!)', '=#REF!', `=${name}!#REF!`]);
});

test('references: whole columns and rows, other sheets\' own references', () => {
  const o = { here: 'S', sheet: 'S', axis: 'row', at: 2, n: 1 };
  assert.equal(shiftRefs('SUM(A:A)+SUM(3:5)', o), 'SUM(A:A)+SUM(4:6)');
  assert.equal(shiftRefs('A5+T!A5', o), 'A6+T!A5');
  assert.equal(shiftRefs('A5', { ...o, here: 'T' }), 'A5');
  assert.equal(shiftRefs("'S'!A5", { ...o, here: null }), "'S'!A6");
});

test('a rule of the Excel 2010 form reading the sheet: its formula moved once, its xm:f kept', () => {
  const { wb, si, s2 } = setup();
  const name = wb.sheets[si].name;
  // (A list on the next sheet from this one's A1:A3, written as Excel writes a rule reading another sheet.)
  const doc = wb._part(s2).doc;
  const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
  const XM = 'http://schemas.microsoft.com/office/excel/2006/main';
  const ext = parseXml(`<extLst xmlns="${doc.documentElement.namespaceURI}"><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="${X14}"><x14:dataValidations count="1" xmlns:xm="${XM}">` +
    `<x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>${name}!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>B1</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>`).documentElement;
  doc.documentElement.appendChild(doc.importNode(ext, true));
  wb.changeLines(si, 'row', 0, 1);
  const f = doc.getElementsByTagNameNS(XM, 'f');
  assert.equal(f.length, 1);
  assert.equal(f[0].textContent, `${name}!$A$2:$A$4`);
});

test('refused as Excel refuses', () => {
  const { wb, si } = setup();
  wb.apply(si, [{ r: 1048575, c: 0, value: 1 }]);
  assert.throws(() => wb.changeLines(si, 'row', 2, 1), { message: ROWCOL_MESSAGES.offSheet });
});

test('undone through a snapshot; the saved workbook passes the structural check', () => {
  const { wb, si, s2 } = setup();
  const snap = wb.snapshot();
  wb.changeLines(si, 'row', 2, 2);
  wb.restore(snap);
  assert.equal(col(wb, si), '=SUM(A1:A5) =A5 =A3 =SUM(A3:A3) =$A$4 =SUM(A6:A9) ');
  assert.equal(f(wb, s2, 'A1').endsWith('!A4'), true);
  wb.changeLines(si, 'row', 2, 1);
  wb.changeLines(si, 'col', 1, -1);
  const bytes = wb.toBytes();
  fs.writeFileSync(`${dir}/changed.xlsx`, bytes);
  const out = spawnSync(process.execPath, ['tools/check-ooxml.mjs', `${dir}/changed.xlsx`], { encoding: 'utf8' });
  assert.equal(out.status, 0, out.stdout + out.stderr);
  assert.match(strFromU8(unzipSync(bytes)['xl/workbook.xml']), /<definedName name="Four">[^<]*\$A\$5<\/definedName>/);
});
