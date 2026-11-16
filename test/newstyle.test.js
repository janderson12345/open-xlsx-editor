// New Cell Style and Modify as Excel makes them (src/ooxml/cellstyledefine.js): each case done here from the styles
// Excel's new workbook starts with, compared with what Excel saved for the same over COM (test/newstyle-excel.json,
// tools/excel-compare/excel-newstyle.ps1): the style's cellStyleXfs entry and cellStyle, and the cells' xfs, each with
// the font, fill, border and number format its ids point to (the lists' order aside).
import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { serializeXml } from '../src/ooxml/xml.js';
import { newCellStyle, modifyCellStyle, styleInfo, styleNameRefusal } from '../src/ooxml/cellstyledefine.js';

const excel = JSON.parse(fs.readFileSync('test/newstyle-excel.json', 'utf8').trim());
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

/** An xf as compared: its attributes but the ids, its children, and what the ids point to. */
function resolve(xf, parts) {
  const id = (k) => Number(new RegExp(` ${k}="(\\d+)"`).exec(xf)?.[1] || 0);
  return {
    // (Its attributes in any order: the cell's own format here was made a change at a time, Excel's at once.)
    xf: xf.replace(/ (fontId|fillId|borderId|xfId)="\d+"/g, '').replace(/^<xf((?: [\w]+="[^"]*")*)/, (m, a) => `<xf${a.split(' ').sort().join(' ')}`),
    font: items(parts.fonts, 'font')[id('fontId')],
    fill: items(parts.fills, 'fill')[id('fillId')],
    border: items(parts.borders, 'border')[id('borderId')],
  };
}
/** A saved (or our) styles part's look of the style named `name` and of cell formats `ss`; the cellStyles. */
function looks(parts, name, ss) {
  const cs = new RegExp(`<cellStyle name="${name}" xfId="(\\d+)"`).exec(parts.cellStyles);
  return {
    style: resolve(items(parts.cellStyleXfs, 'xf')[Number(cs[1])], parts),
    cells: ss.map((s) => resolve(items(parts.cellXfs, 'xf')[s], parts)),
    cellStyles: parts.cellStyles.replace(/ xr:uid="[^"]*"/g, '').replace(/ xfId="\d+"/g, '').replace(/ count="\d+"/, ''),
  };
}
const cellS = (sheetData, ref) => Number(new RegExp(`<c r="${ref}"(?: s="(\\d+)")?`).exec(sheetData)?.[1] || 0);

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
/** B2 as Excel's case formatted it: bold red 14pt, yellow, a thin bottom border, 0.00, centred. */
function example(c) {
  for (const ch of [{ font: { b: true } }, { font: { color: { rgb: 'FFFF0000' } } }, { font: { size: 14 } }, { fill: { rgb: 'FFFFFF00' } },
    { border: { bottom: { style: 'thin' } } }, { numFmt: '0.00' }, { align: { horizontal: 'center' } }]) c.set('B2', ch);
}
const ourParts = (wb) => {
  const text = serializeXml(wb.stylesDoc);
  return Object.fromEntries(['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles'].map((t) => [t, pick(text, t)]));
};
const theirs = (key, name, refs) => looks(excel[key], name, refs.map((ref) => cellS(excel[key].sheetData, ref)));

test('New Cell Style by example: the style takes the cell\'s parts, a cellStyle in name order, the cell left as it was', () => {
  const c = start();
  example(c);
  const before = c.s('B2');
  assert.ok(!newCellStyle(c.wb, before, 'Mine').error);
  assert.deepEqual(looks(ourParts(c.wb), 'Mine', [c.s('B2')]), theirs('new from B2', 'Mine', ['B2']));
  assert.equal(c.s('B2'), before);
  assert.equal(excel['new from B2'].notes.b2Style, 'Normal');
});

test('New Cell Style with its border and alignment left out: applyBorder="0" applyAlignment="0"', () => {
  const c = start();
  example(c);
  newCellStyle(c.wb, c.s('B2'), 'Mine', { border: false, alignment: false });
  assert.deepEqual(looks(ourParts(c.wb), 'Mine', [c.s('B2')]), theirs('new from B2 without border and alignment', 'Mine', ['B2']));
});

test('Modify: the style changed in place, and the cells using it with it, but for a part a cell holds as its own', () => {
  const c = start();
  example(c);
  newCellStyle(c.wb, c.s('B2'), 'Mine');
  for (const ref of ['D4', 'D5']) c.set(ref, { cellStyle: 'Mine' });
  c.set('D5', { font: { i: true } });
  assert.deepEqual(looks(ourParts(c.wb), 'Mine', ['B2', 'D4', 'D5'].map(c.s)), theirs('new applied to D4 and D5', 'Mine', ['B2', 'D4', 'D5']));
  const count = c.wb.styles.cellXfs.length;
  const out = modifyCellStyle(c.wb, 'Mine', { font: { b: false }, fill: { rgb: 'FF92D050' }, numFmt: '0%', align: { horizontal: 'left' } });
  assert.deepEqual(out.formats.length, 2);
  assert.equal(c.wb.styles.cellXfs.length, count, 'the cells\' formats changed in place, none added');
  assert.deepEqual(looks(ourParts(c.wb), 'Mine', ['B2', 'D4', 'D5'].map(c.s)), theirs('modified after D4 and D5', 'Mine', ['B2', 'D4', 'D5']));
});

test('Modify a built-in style: marked customBuiltin, a part it leaves out not given to its cells', () => {
  const c = start();
  for (const ref of ['B2', 'B3']) c.set(ref, { cellStyle: 'Good' });
  c.set('B3', { font: { color: { rgb: 'FFFF0000' } } });
  assert.equal(styleInfo(c.wb, 'Good').includes.border, false);
  modifyCellStyle(c.wb, 'Good', { font: { b: true }, borders: { outline: { left: { style: 'thin' } } } });
  assert.deepEqual(looks(ourParts(c.wb), 'Good', ['B2', 'B3'].map(c.s)), theirs('Good modified', 'Good', ['B2', 'B3']));
});

test('Modify Normal: every cell of no format of its own with it, a cell\'s own bold font kept', () => {
  const c = start();
  c.set('C3', { font: { b: true } });
  modifyCellStyle(c.wb, 'Normal', { font: { size: 14 } });
  const ours = looks(ourParts(c.wb), 'Normal', [0, c.s('C3')]);
  const want = theirs('Normal modified', 'Normal', ['B2', 'C3']);
  assert.deepEqual(ours, want);
  assert.equal(c.wb.styles.normal.size, 14);
});

test('a new style\'s name: refused when empty or taken (a built-in one too, in any case), a style renamed in Modify', () => {
  const c = start();
  assert.match(styleNameRefusal(c.wb, ' '), /Type a name/);
  assert.match(styleNameRefusal(c.wb, 'good'), /already exists/);
  newCellStyle(c.wb, 0, 'Mine');
  assert.match(newCellStyle(c.wb, 0, 'MINE').error, /already exists/);
  assert.ok(!modifyCellStyle(c.wb, 'Mine', null, null, 'Ours').error);
  assert.ok(styleInfo(c.wb, 'Ours') && !styleInfo(c.wb, 'Mine'));
  assert.match(modifyCellStyle(c.wb, 'Good', null, null, 'Fine').error, /can't be renamed/);
});
