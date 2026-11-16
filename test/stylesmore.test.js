// Delete, Duplicate and Merge Styles as Excel does them (src/ooxml/cellstylemanage.js): each case done here from the
// styles Excel's new workbook starts with, compared with what Excel saved for the same over COM
// (test/stylesmore-excel.json, tools/excel-compare/excel-stylesmore.ps1): every style entry, the cellStyles, and the
// cells' xfs, each with the font, fill, border and number format its ids point to and the style it follows (the lists'
// order aside: Excel numbers cellStyleXfs anew when it saves, Normal first, then by name, hidden ones last).
import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { serializeXml } from '../src/ooxml/xml.js';
import { newCellStyle, modifyCellStyle, styleInfo } from '../src/ooxml/cellstyledefine.js';
import { customCellStyles, deletedBuiltins } from '../src/ooxml/cellstyleedit.js';
import { deleteCellStyle, duplicateCellStyle, duplicateName, mergeCellStyles, sameNameStyles } from '../src/ooxml/cellstylemanage.js';

const excel = JSON.parse(fs.readFileSync('test/stylesmore-excel.json', 'utf8').trim());
const FONT = '<font><sz val="11"/><color theme="1"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font>';
// (Excel 365's new workbook's styles, as its saved files show them.)
const START = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
  + `<fonts count="1">${FONT}</fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>`
  + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
  + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
  + '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
  + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

const items = (xml, tag) => [...(xml || '').replace(new RegExp(`^<${tag}s[^>]*>`), '').matchAll(new RegExp(`<${tag}(?: [^>]*?)?(?:/>|>[\\s\\S]*?</${tag}>)`, 'g'))].map((x) => x[0]);
const pick = (text, tag) => new RegExp(`<${tag}\\b[\\s\\S]*?</${tag}>`).exec(text)?.[0] || null;
const cellStylesOf = (parts) => items(parts.cellStyles, 'cellStyle');
/** The style name of style entry `i` (its cellStyle's), or '' for an entry no cellStyle names. */
const styleName = (parts, i) => cellStylesOf(parts).map((cs) => [/ name="([^"]*)"/.exec(cs)[1], Number(/ xfId="(\d+)"/.exec(cs)[1])]).find(([, x]) => x === i)?.[0] ?? '';

/** An xf as compared: its attributes but the ids (in any order), its children, what the ids point to, and the style
 *  it follows. */
function resolve(xf, parts) {
  const id = (k) => Number(new RegExp(` ${k}="(\\d+)"`).exec(xf)?.[1] || 0);
  const n = id('numFmtId');
  const code = new RegExp(`<numFmt numFmtId="${n}" formatCode="([^"]*)"`).exec(parts.numFmts || '')?.[1];
  return {
    xf: xf.replace(/ (numFmtId|fontId|fillId|borderId|xfId)="\d+"/g, '').replace(/^<xf((?: [\w]+="[^"]*")*)/, (m, a) => `<xf${a.split(' ').sort().join(' ')}`),
    numFmt: code ?? n,
    font: items(parts.fonts, 'font')[id('fontId')],
    fill: items(parts.fills, 'fill')[id('fillId')],
    border: items(parts.borders, 'border')[id('borderId')],
    style: / xfId="/.test(xf) ? styleName(parts, id('xfId')) : undefined,
  };
}
/** A saved (or our) styles part's look: every style entry (by its style's name), the cellStyles (their order and
 *  attributes but xfId and Excel's xr:uid), and cell formats `ss`. */
function looks(parts, ss) {
  const xfs = items(parts.cellStyleXfs, 'xf').map((xf, i) => ({ name: styleName(parts, i), ...resolve(xf, parts) }));
  return {
    styles: xfs.sort((a, b) => (a.name < b.name ? -1 : a.name > b.name ? 1 : 0)),
    cellStyles: cellStylesOf(parts).map((cs) => cs.replace(/ xr:uid="[^"]*"/g, '').replace(/ xfId="\d+"/g, '')),
    cells: ss.map((s) => resolve(items(parts.cellXfs, 'xf')[s], parts)),
  };
}
const cellS = (sheetData, ref) => Number(new RegExp(`<c r="${ref}"(?: s="(\\d+)")?`).exec(sheetData)?.[1] || 0);
const theirs = (key, refs) => looks(excel[key], refs.map((ref) => cellS(excel[key].sheetData, ref)));
const ourParts = (wb) => {
  const text = serializeXml(wb.stylesDoc);
  return Object.fromEntries(['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles'].map((t) => [t, pick(text, t)]));
};
const ours = (c, refs) => looks(ourParts(c.wb), refs.map(c.s));

function start() {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  wb.pkg.setText(wb.stylesPath, START);
  const fresh = Workbook.fromBytes(wb.toBytes());
  const si = fresh.addSheet(fresh.sheets.length);
  const at = (ref) => ({ r: Number(ref.slice(1)) - 1, c: ref.charCodeAt(0) - 65 });
  const set = (ref, change) => {
    const { r, c } = at(ref);
    const s = fresh.sheet(si).cells.get(r)?.get(c)?.s || 0;
    fresh.setStyles(si, [{ r, c, s: fresh.styleWith(s, change) }]);
  };
  const s = (ref) => { const { r, c } = at(ref); return fresh.sheet(si).cells.get(r)?.get(c)?.s || 0; };
  return { wb: fresh, si, set, s };
}
/** Excel's case: B2 bold red 14pt, yellow, a thin bottom border, 0.00, centred; Mine made from it, on D4 and D5 (D5
 *  then italic of its own). */
function mine(c) {
  for (const ch of [{ font: { b: true } }, { font: { color: { rgb: 'FFFF0000' } } }, { font: { size: 14 } }, { fill: { rgb: 'FFFFFF00' } },
    { border: { bottom: { style: 'thin' } } }, { numFmt: '0.00' }, { align: { horizontal: 'center' } }]) c.set('B2', ch);
  newCellStyle(c.wb, c.s('B2'), 'Mine');
  for (const ref of ['D4', 'D5']) c.set(ref, { cellStyle: 'Mine' });
  c.set('D5', { font: { i: true } });
}
/** A workbook saved and opened again: what another editor on the page would hold. */
const reopened = (c) => Workbook.fromBytes(c.wb.toBytes());

test('Delete a style of the workbook\'s own: its entry and cellStyle gone, its cells back to Normal but for their own font', () => {
  const c = start();
  mine(c);
  const count = c.wb.styles.cellXfs.length;
  const out = deleteCellStyle(c.wb, 'Mine');
  assert.equal(out.formats.length, 2);
  assert.equal(c.wb.styles.cellXfs.length, count, 'the cells\' formats changed in place');
  assert.deepEqual(ours(c, ['B2', 'D4', 'D5']), theirs('delete Mine', ['B2', 'D4', 'D5']));
  assert.deepEqual(customCellStyles(c.wb), []);
  assert.equal(styleInfo(c.wb, 'Mine'), null);
});

test('Delete a style no cell uses', () => {
  const c = start();
  c.set('B2', { font: { b: true } });
  newCellStyle(c.wb, c.s('B2'), 'Mine');
  deleteCellStyle(c.wb, 'Mine');
  assert.deepEqual(ours(c, ['B2']), theirs('delete unused Mine', ['B2']));
});

test('Delete a built-in style: kept hidden, its cells back to Normal, the entries after it still pointed at', () => {
  const c = start();
  for (const ref of ['B2', 'B3']) c.set(ref, { cellStyle: 'Good' });
  c.set('B3', { font: { color: { rgb: 'FFFF0000' } } });
  c.set('B4', { cellStyle: 'Title' });
  deleteCellStyle(c.wb, 'Good');
  assert.deepEqual(ours(c, ['B2', 'B3', 'B4']), theirs('delete Good', ['B2', 'B3', 'B4']));
  assert.deepEqual([...deletedBuiltins(c.wb)], [26]);
  // (Excel refuses to give it to a cell again: Range.Style = "Good" fails over COM.)
  const s = c.s('C9');
  c.set('C9', { cellStyle: 'Good' });
  assert.equal(c.s('C9'), s);
  assert.match(excel['delete Good then apply'].notes.error, /Style/);
  assert.match(deleteCellStyle(c.wb, 'Good').error, /no style named Good/);
});

test('Delete a built-in style the workbook hasn\'t got: added hidden, as Excel adds it', () => {
  const c = start();
  deleteCellStyle(c.wb, 'Good');
  assert.deepEqual(ours(c, []), theirs('delete Good unused', []));
});

test('Normal can\'t be deleted', () => {
  const c = start();
  assert.match(deleteCellStyle(c.wb, 'Normal').error, /Normal style can't be deleted/);
  assert.match(excel['delete Normal'].notes.error, /Unable to get the Delete property/);
});

test('Duplicate a style of the workbook\'s own: "Mine 2", its parts as Mine\'s, a part it leaves out Normal\'s', () => {
  const c = start();
  mine(c);
  modifyCellStyle(c.wb, 'Mine', null, { ...styleInfo(c.wb, 'Mine').includes, border: false });
  assert.equal(duplicateName(c.wb, 'Mine'), 'Mine 2');
  assert.ok(!duplicateCellStyle(c.wb, 'Mine', 'Mine 2').error);
  assert.deepEqual(ours(c, ['B2', 'D4', 'D5']), theirs('duplicate Mine', ['B2', 'D4', 'D5']));
  assert.deepEqual(customCellStyles(c.wb), excel['duplicate Mine'].notes.names);
  assert.equal(duplicateName(c.wb, 'Mine'), 'Mine 3');
  assert.match(duplicateCellStyle(c.wb, 'Mine', 'mine 2').error, /already exists/);
});

test('Duplicate a built-in style the workbook hasn\'t got: "Good 2" of its own, Good itself not added', () => {
  const c = start();
  duplicateCellStyle(c.wb, 'Good', duplicateName(c.wb, 'Good'));
  assert.deepEqual(ours(c, []), theirs('duplicate Good', []));
});

test('Duplicate with the Style box\'s Format...: the copy changed, the style it came from not', () => {
  const c = start();
  mine(c);
  duplicateCellStyle(c.wb, 'Mine', 'Mine 2', null, { font: { b: false } });
  const font = (n) => c.wb.styles.fonts[styleInfo(c.wb, n).xf.fontId];
  assert.equal(!!font('Mine').b, true);
  assert.equal(!!font('Mine 2').b, false);
});

test('Merge Styles from a workbook with styles of other names: each added, a number format of its own given a new id', () => {
  const from = start();
  from.set('A1', { font: { b: true } });
  from.set('A1', { numFmt: '0.0" kg"' });
  from.set('A1', { fill: { rgb: 'FF92D050' } });
  newCellStyle(from.wb, from.s('A1'), 'Other');
  from.set('A2', { font: { i: true } });
  newCellStyle(from.wb, from.s('A2'), 'Third', { number: false, border: false, fill: false, alignment: false, protection: false });
  const c = start();
  c.set('A1', { numFmt: '0.000' });
  const other = reopened(from);
  assert.deepEqual(sameNameStyles(c.wb, other), [], 'no question: Normal the same in both');
  const out = mergeCellStyles(c.wb, other);
  assert.deepEqual(out.added, ['Other', 'Third']);
  assert.deepEqual(ours(c, ['A1']), theirs('merge new names', ['A1']));
  assert.deepEqual(customCellStyles(c.wb), excel['merge new names'].notes.names);
});

test('Merge Styles with styles of the same names, answered Yes: put in place of this workbook\'s, its cells with them', () => {
  const from = start();
  for (const ch of [{ font: { i: true } }, { font: { size: 12 } }, { fill: { rgb: 'FF92D050' } }, { border: { bottom: { style: 'dashed' } } }, { numFmt: '0%' }, { align: { horizontal: 'left' } }]) from.set('A1', ch);
  newCellStyle(from.wb, from.s('A1'), 'Mine');
  modifyCellStyle(from.wb, 'Good', { font: { b: true } });
  from.set('A2', { font: { b: true } });
  newCellStyle(from.wb, from.s('A2'), 'Other');
  const c = start();
  mine(c);
  c.set('B6', { cellStyle: 'Good' });
  const other = reopened(from);
  assert.deepEqual(sameNameStyles(c.wb, other).sort(), ['Good', 'Mine']);
  const out = mergeCellStyles(c.wb, other, true);
  assert.deepEqual(out.replaced.sort(), ['Good', 'Mine']);
  assert.deepEqual(ours(c, ['B2', 'D4', 'D5', 'B6']), theirs('merge same names', ['B2', 'D4', 'D5', 'B6']));
});

test('Merge Styles answered No: the same names kept as they were, the others added', () => {
  const from = start();
  from.set('A1', { font: { i: true } });
  newCellStyle(from.wb, from.s('A1'), 'Mine');
  newCellStyle(from.wb, from.s('A1'), 'Other');
  const c = start();
  mine(c);
  const before = ours(c, ['D4', 'D5']).cells;
  const out = mergeCellStyles(c.wb, reopened(from), false);
  assert.deepEqual(out, { added: ['Other'], replaced: [], formats: [] });
  assert.deepEqual(ours(c, ['D4', 'D5']).cells, before);
});

test('Merge Styles\' built-in styles: Normal and Title from the other workbook, Bad changed there added, all customBuiltin', () => {
  const from = start();
  from.set('A1', { cellStyle: 'Title' });
  // (Bad first: the editor adds a built-in style at Normal's size, where Excel's Bad, changed after Normal, stayed 11pt.)
  modifyCellStyle(from.wb, 'Bad', { font: { i: true } });
  modifyCellStyle(from.wb, 'Normal', { font: { size: 14 } });
  const c = start();
  c.set('B2', { cellStyle: 'Title' });
  modifyCellStyle(c.wb, 'Title', { font: { b: true } });
  mergeCellStyles(c.wb, reopened(from));
  // (C3 has no format of its own: Normal's, at 14pt.)
  assert.deepEqual(ours(c, ['B2', 'C3']), theirs('merge built-ins', ['B2', 'C3']));
  assert.equal(c.wb.styles.normal.size, 14);
});

test('Merge Styles from a workbook into itself, or one with no styles part, changes nothing', () => {
  const c = start();
  mine(c);
  assert.deepEqual(mergeCellStyles(c.wb, c.wb), { added: [], replaced: [], formats: [] });
  assert.deepEqual(mergeCellStyles(c.wb, null), { added: [], replaced: [], formats: [] });
});
