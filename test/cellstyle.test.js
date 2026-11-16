// Cell styles applied as Excel applies them (src/ooxml/cellstyleedit.js): each of Excel's 47 built-in styles applied
// here to a cell of a workbook without it, and Good over a bold red cell, Heading 1 over Good, compared with what Excel
// saved for the same (test/cellstyles-excel.json, tools/excel-compare/excel-cellstyles.ps1): the style's cellStyleXfs
// entry and cellStyle, and the cell's xf, each with the font, fill, border and number format its ids point to (the
// font's typeface aside: Excel's new workbook is in Aptos Narrow, ours in the fixture's).
import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { serializeXml } from '../src/ooxml/xml.js';
import { BUILTIN_CELL_STYLES } from '../src/engine/cellstyles.js';

const excel = JSON.parse(fs.readFileSync('test/cellstyles-excel.json', 'utf8').trim());
const items = (xml, tag) => [...(xml || '').replace(new RegExp(`^<${tag}s[^>]*>`), '').matchAll(new RegExp(`<${tag}(?: [^>]*)?(?:/>|>[\\s\\S]*?</${tag}>)`, 'g'))].map((x) => x[0]);
const plainFont = (x) => (x || '').replace(/<name val="[^"]*"\/><family val="\d+"\/>(<charset val="\d+"\/>)?(<scheme val="minor"\/>)?/, '<minor/>').replace(/<name val="[^"]*"\/><family val="\d+"\/><scheme val="major"\/>/, '<major/>');
const NO_FILLS = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
const NO_BORDER = '<border><left/><right/><top/><bottom/><diagonal/></border>';

/** An xf as compared: its attributes but the ids, and what the ids point to. */
function resolve(xf, parts) {
  const id = (k) => Number(new RegExp(` ${k}="(\\d+)"`).exec(xf)?.[1] || 0);
  const fill = parts.fills[id('fillId')];
  const border = parts.borders[id('borderId')];
  const numFmtId = id('numFmtId');
  return {
    attrs: xf.replace(/ (numFmtId|fontId|fillId|borderId|xfId)="\d+"/g, ''),
    font: plainFont(parts.fonts[id('fontId')]),
    fill: NO_FILLS.includes(fill) ? 'none' : fill,
    border: border === NO_BORDER ? 'none' : border,
    numFmt: numFmtId ? (new RegExp(`<numFmt numFmtId="${numFmtId}" formatCode="([^"]*)"/>`).exec(parts.numFmts || '')?.[1] ?? numFmtId) : 0,
  };
}

/** What a saved styles part says of cell B2's xf and the style it names. */
function looks(parts, s, name) {
  const lists = { fonts: items(parts.fonts, 'font'), fills: items(parts.fills, 'fill'), borders: items(parts.borders, 'border'), numFmts: parts.numFmts };
  const xf = items(parts.cellXfs, 'xf')[s];
  // (Every character a regular expression gives meaning to escaped, backslashes too: CodeQL js/incomplete-sanitization.)
  const esc = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const cs = new RegExp(`<cellStyle name="${esc}" xfId="(\\d+)"( builtinId="\\d+")?/>`).exec(parts.cellStyles);
  const styleXf = cs ? items(parts.cellStyleXfs, 'xf')[Number(cs[1])] : null;
  return { cell: resolve(xf, lists), style: styleXf ? resolve(styleXf, lists) : null, builtin: cs?.[2] || null };
}

function ours(steps) {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  // (A cell of no format of its own, as Excel's B2.)
  wb.setStyles(0, [{ r: 1, c: 1, s: 0 }]);
  for (const st of steps) {
    const s = wb.sheet(0).cells.get(1)?.get(1)?.s || 0;
    wb.setStyles(0, [{ r: 1, c: 1, s: wb.styleWith(s, st) }]);
  }
  const text = serializeXml(wb.stylesDoc);
  const pick = (tag) => new RegExp(`<${tag}\\b[\\s\\S]*?</${tag}>`).exec(text)?.[0] || null;
  return { parts: { numFmts: pick('numFmts'), fonts: pick('fonts'), fills: pick('fills'), borders: pick('borders'), cellStyleXfs: pick('cellStyleXfs'), cellXfs: pick('cellXfs'), cellStyles: pick('cellStyles') }, s: wb.sheet(0).cells.get(1).get(1).s };
}

test('the 47 built-in cell styles measured', () => {
  assert.equal(BUILTIN_CELL_STYLES.length, 47);
  assert.deepEqual(excel.styles.map((x) => x.name).sort(), BUILTIN_CELL_STYLES.map((x) => x.name).sort());
});

test('each built-in cell style applied as Excel applies it, in a workbook without it', () => {
  const diffs = [];
  for (const { name } of excel.styles) {
    const saved = excel.saved[name];
    const s = Number(/ s="(\d+)"/.exec(saved.cell)?.[1] || 0);
    const theirs = looks(saved, s, name);
    const mine = ours([{ cellStyle: name }]);
    const got = looks(mine.parts, mine.s, name);
    if (JSON.stringify(got) !== JSON.stringify(theirs)) diffs.push(`${name}: ours ${JSON.stringify(got)}\n   Excel ${JSON.stringify(theirs)}`);
  }
  assert.deepEqual(diffs, []);
});

test('a style over a format of its own: Good over a bold red cell, Heading 1 over Good', () => {
  for (const [key, steps, name] of [['Good over bold red', [{ font: { b: true, color: { rgb: 'FFFF0000' } } }, { cellStyle: 'Good' }], 'Good'], ['Good then Heading 1', [{ cellStyle: 'Good' }, { cellStyle: 'Heading 1' }], 'Heading 1']]) {
    const saved = excel.saved[key];
    const theirs = looks(saved, Number(/ s="(\d+)"/.exec(saved.cell)?.[1] || 0), name);
    const mine = ours(steps);
    assert.deepEqual(looks(mine.parts, mine.s, name), theirs, key);
  }
});
