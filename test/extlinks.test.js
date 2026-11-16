import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseRef, refName } from '../src/ooxml/refs.js';
import { csvWorkbook } from '../src/ooxml/csv.js';
import { addSheet, renameSheet } from '../src/ooxml/sheetops.js';
import { entryChanges } from '../src/ooxml/entry.js';
import { defineName, editName, listNames } from '../src/ooxml/names.js';
import { snapshot, pasteLink } from '../src/ooxml/clipops.js';
import { readLinks, linkCopy, bookNames, bookIndexes, refreshLink, breakLinkSteps, updateLink } from '../src/ooxml/extlinks.js';

// Links to another workbook as Excel makes them (tools/excel-compare/excel-extlinks.ps1: Range.Copy in a source
// workbook, Worksheet.Paste with Link in another, names added that refer to the source, and the file saved; what it
// wrote reduced to test/extlinks-excel.json by extlinks-extract.mjs). Here the same two workbooks are built, the same
// links pasted and names defined, and the file's formulas, values, names and the link's part compared with Excel's.
const excel = JSON.parse(fs.readFileSync('test/extlinks-excel.json', 'utf8'));

const blank = (fileName) => {
  const wb = csvWorkbook(new Uint8Array(0), 'Sheet1.csv');
  wb.fileName = fileName;
  return wb;
};
const type = (wb, si, ref, text) => {
  const { r, c } = parseRef(ref);
  const ch = entryChanges(wb, si, r, c, text);
  wb.apply(si, [ch.content]);
  if (ch.style) wb.setStyles(si, [ch.style]);
};
const area = (text) => { const [a, b = a] = text.split(':').map(parseRef); return { r0: a.r, c0: a.c, r1: b.r, c1: b.c, ar: a.r, ac: a.c }; };

/** The source workbook as the probe made it: Sheet1 with numbers, text, TRUE, a formula and an error; Data Two; a name. */
function source() {
  const src = blank('oxesrc.xlsx');
  for (const [ref, text] of [['A1', '5'], ['B1', '2.5'], ['A2', 'text'], ['B2', '=A1*2'], ['C1', 'TRUE'], ['C2', '=1/0']]) type(src, 0, ref, text);
  addSheet(src, 1);
  renameSheet(src, 1, 'Data Two');
  type(src, 1, 'A1', '7');
  defineName(src, { name: 'SrcRate', refersTo: '=Sheet1!$B$1' });
  src.calculate();
  return src;
}

/** The workbook linking to it, the probe's steps done here. */
function linked(src = source()) {
  const dst = blank('oxedst.xlsx');
  dst.openBook = (name) => (name.toLowerCase() === 'oxesrc.xlsx' ? src : null);
  for (const [sheet, from, to] of [['Sheet1', 'A1:C2', 'D4'], ['Sheet1', 'B1', 'A1'], ['Data Two', 'A1', 'A10'], ['Sheet1', 'A1:A3', 'H1']]) {
    const clip = snapshot(src, src.sheetIndexOf(sheet), area(from));
    const book = linkCopy(dst, clip, src.fileName);
    for (const x of pasteLink(dst, 0, area(to), clip, { book })) {
      dst.apply(x.si, x.contents);
      if (x.styles.length) dst.setStyles(x.si, x.styles);
    }
  }
  defineName(dst, { name: 'ExtRate', refersTo: '=[oxesrc.xlsx]Sheet1!$A$1' });
  defineName(dst, { name: 'ExtBlock', refersTo: '=[oxesrc.xlsx]Sheet1!$A$1:$B$2' });
  defineName(dst, { name: 'LocalExt', refersTo: "='[oxesrc.xlsx]Data Two'!$A$1", sheet: 0 });
  defineName(dst, { name: 'ExtConst', refersTo: '=[oxesrc.xlsx]Sheet1!$A$1*2' });
  type(dst, 0, 'A12', '=SUM(ExtBlock)');
  // (Excel's own typing is =oxesrc.xlsx!SrcRate; here a book's name is typed in brackets: [oxesrc.xlsx]!SrcRate.)
  type(dst, 0, 'A13', '=[oxesrc.xlsx]!SrcRate');
  type(dst, 0, 'A14', '=[oxesrc.xlsx]Sheet1!$A$1+1');
  dst.calculate();
  return { src, dst };
}

/** A cell as the file has it: formula, value as written, type. */
function fileCell(wb, si, ref) {
  const { r, c } = parseRef(ref);
  const cell = wb.sheet(si).cells.get(r)?.get(c);
  const v = cell?.v;
  return { formula: cell?.f ?? null, value: v == null ? null : typeof v === 'boolean' ? (v ? '1' : '0') : String(v), t: cell.t === 's' ? 'str' : cell.t };
}

test("Paste Link to another workbook: the formulas and values Excel's file has", () => {
  const { dst } = linked();
  for (const [ref, want] of Object.entries(excel.cells)) assert.deepEqual(fileCell(dst, 0, ref), want, ref);
  // (And after saving and opening again, worked out from the values the link keeps.)
  const again = Workbook.fromBytes(dst.toBytes());
  again.resetCalc();
  again.calculate();
  for (const [ref, want] of Object.entries(excel.cells)) assert.deepEqual(fileCell(again, 0, ref), want, `${ref} reopened`);
});

test('Names that refer to another workbook: written as Excel writes them, shown with its file name', () => {
  const { dst } = linked();
  const ours = Object.fromEntries(dst.definedNames.map((d) => [d.name + (d.sheet != null ? `@${d.sheet}` : ''), d.text]));
  assert.deepEqual(ours, excel.names);
  const shown = Object.fromEntries(listNames(dst).map((n) => [n.name, n.refersTo]));
  assert.deepEqual(shown, { ExtBlock: '=[oxesrc.xlsx]Sheet1!$A$1:$B$2', ExtConst: '=[oxesrc.xlsx]Sheet1!$A$1*2', ExtRate: '=[oxesrc.xlsx]Sheet1!$A$1', LocalExt: "='[oxesrc.xlsx]Data Two'!$A$1" });
  // (Edited: the file's form again; the formula using it worked out from the link.)
  editName(dst, { name: 'ExtBlock' }, { refersTo: '=[oxesrc.xlsx]Sheet1!$A$1:$A$1' });
  assert.equal(dst.definedNames.find((d) => d.name === 'ExtBlock').text, '[1]Sheet1!$A$1:$A$1');
  dst.calculate();
  assert.equal(fileCell(dst, 0, 'A12').value, '5');
});

test("The link's part: Excel's sheets, names and kept values, and its file", () => {
  const { dst } = linked();
  const [link] = readLinks(dst);
  assert.equal(readLinks(dst).length, 1);
  // (Excel listed a third sheet, oxesrc, with refreshError: its own reading of =oxesrc.xlsx!SrcRate as a sheet's
  // name before it found the defined name. Not made here.)
  assert.deepEqual(link.sheetNames, excel.link.sheetNames.slice(0, 2));
  assert.deepEqual(link.names, excel.link.names.map((n) => ({ ...n, sheetId: null })));
  link.sheetNames.forEach((_, k) => {
    const want = excel.link.sheetData[k].cells;
    const got = [...link.cells[k]].map(([key, x]) => { const [r, c] = key.split(',').map(Number); return { r: refName(r, c), t: x.t, v: x.v }; })
      .sort((a, b) => parseRef(a.r).r - parseRef(b.r).r || parseRef(a.r).c - parseRef(b.r).c);
    assert.deepEqual(got, want, link.sheetNames[k]);
  });
  assert.equal(link.target, excel.link.target);
  fs.mkdirSync('test/out', { recursive: true });
  fs.writeFileSync('test/out/extlinks.xlsx', dst.toBytes());
});

test('A workbook with links, untouched, is saved byte for byte; a cell typed elsewhere leaves its links alone', () => {
  const bytes = linked().dst.toBytes();
  const wb = Workbook.fromBytes(bytes);
  const before = { ...wb.pkg.files };
  const same = Workbook.fromBytes(wb.toBytes());
  for (const [path, data] of Object.entries(before)) assert.deepEqual(same.pkg.files[path], data, path);
  type(wb, 0, 'Z9', '42');
  const after = Workbook.fromBytes(wb.toBytes());
  for (const path of ['xl/externalLinks/externalLink1.xml', 'xl/externalLinks/_rels/externalLink1.xml.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', '[Content_Types].xml']) {
    assert.deepEqual(after.pkg.files[path], before[path], path);
  }
});

test('Formulas shown with the file name and typed back to the link number', () => {
  const { dst } = linked();
  assert.equal(bookNames(dst, "SUM([1]Sheet1!A1:B2)+'[1]Data Two'!$A$1+[1]!SrcRate"), "SUM([oxesrc.xlsx]Sheet1!A1:B2)+'[oxesrc.xlsx]Data Two'!$A$1+[oxesrc.xlsx]!SrcRate");
  assert.equal(bookIndexes(dst, "SUM([oxesrc.xlsx]Sheet1!A1:B2)+'[OXESRC.xlsx]Data Two'!$A$1"), "SUM([1]Sheet1!A1:B2)+'[1]Data Two'!$A$1");
  // (A workbook not open here: a link of its own, #REF! until its values are read.)
  type(dst, 0, 'J1', '=[Other.xlsx]Sheet1!B2');
  dst.calculate();
  assert.equal(dst.sheet(0).cells.get(0).get(9).f, '[2]Sheet1!B2');
  assert.equal(String(dst.sheet(0).cells.get(0).get(9).v), '#REF!');
  const other = blank('Other.xlsx');
  type(other, 0, 'B2', '31');
  assert.ok(refreshLink(dst, 2, other));
  assert.equal(dst.sheet(0).cells.get(0).get(9).v, 31);
});

test('Update Values reads the source again; the values are kept until then', () => {
  const { src, dst } = linked();
  type(src, 0, 'A1', '50');
  src.calculate();
  dst.calculate();
  assert.equal(fileCell(dst, 0, 'D4').value, '5');
  assert.ok(refreshLink(dst, 1, src));
  assert.equal(fileCell(dst, 0, 'D4').value, '50');
  assert.equal(fileCell(dst, 0, 'E5').value, '100');
  assert.equal(fileCell(dst, 0, 'A12').value, '152.5');
  assert.equal(fileCell(dst, 0, 'A14').value, '51');
  // (Change Source: another file's name, the relationship's target.)
  updateLink(dst, 1, { target: 'renamed.xlsx' });
  assert.equal(readLinks(dst)[0].target, 'renamed.xlsx');
  assert.equal(bookNames(dst, '[1]Sheet1!A1'), '[renamed.xlsx]Sheet1!A1');
});

test("Break Link as Excel's: formulas using the link become their values; names and formulas through names stay", () => {
  const { dst } = linked();
  for (const st of breakLinkSteps(dst, 1)) dst.apply(st.si, st.contents);
  dst.calculate();
  for (const ref of ['D4', 'E4', 'F4', 'D5', 'E5', 'F5', 'A1', 'A10', 'H1', 'H2', 'A13', 'A14']) assert.equal(fileCell(dst, 0, ref).formula, null, ref);
  assert.equal(fileCell(dst, 0, 'A14').value, '6');
  assert.equal(fileCell(dst, 0, 'F5').value, '#DIV/0!');
  assert.equal(fileCell(dst, 0, 'A12').formula, 'SUM(ExtBlock)');
  assert.equal(dst.definedNames.find((d) => d.name === 'ExtRate').text, '[1]Sheet1!$A$1');
  assert.equal(readLinks(dst).length, 1);
});

test("This workbook's own Sheet1 renamed, or rows inserted in it, leaves references to the other workbook's Sheet1 alone", async () => {
  const { changeLines } = await import('../src/ooxml/rowcol.js');
  const { dst } = linked();
  changeLines(dst, 0, 'row', 0, 2);
  renameSheet(dst, 0, 'Mine');
  assert.equal(fileCell(dst, 0, 'D6').formula, '[1]Sheet1!A1');
  assert.equal(fileCell(dst, 0, 'A3').formula, '[1]Sheet1!$B$1');
  assert.equal(dst.definedNames.find((d) => d.name === 'ExtBlock').text, '[1]Sheet1!$A$1:$B$2');
  dst.calculate();
  assert.equal(fileCell(dst, 0, 'D6').value, '5');
});
