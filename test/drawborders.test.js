import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { spawnSync } from 'node:child_process';
import { Workbook } from '../src/ooxml/workbook.js';
import { xfOf } from '../src/ooxml/styles.js';
import { drawLines } from '../src/ooxml/formatops.js';
import { recentColors, addRecentColor, MAX_RECENT } from '../src/ooxml/recentcolors.js';
import { rgbOf } from '../src/editor/commands-borders.js';
import { borderDrawing } from '../src/editor/drawborders.js';

// Draw Borders and More Colors written as Excel writes them: each scene of tools/excel-compare/excel-drawborders.ps1
// (Excel's saved borders, cell by cell, in test/drawborders-excel.json) drawn here on basic.xlsx and compared with
// Excel's text for text. Draw Border is a mouse gesture COM can't make; its outline, grid and erasing are compared
// with the ribbon's Outside Borders, All Borders and No Border run over COM (ExecuteMso) on the same cells.
const excel = JSON.parse(fs.readFileSync('test/drawborders-excel.json', 'utf8'));
const open = () => Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
const auto = { indexed: 64 };
const cell = (ref) => ({ r: Number(ref.slice(1)) - 1, c: ref.charCodeAt(0) - 65 });
const area = (a, b) => { const x = cell(a); const y = cell(b); return { r0: x.r, c0: x.c, r1: y.r, c1: y.c }; };
/** Give cells' sides lines first, as each scene's neighbours had: [[ref, { side: line }]]. */
const setUp = (wb, list) => wb.setStyles(0, list.map(([ref, border]) => ({ ...cell(ref), s: wb.styleWith(wb.sheet(0).cells.get(cell(ref).r)?.get(cell(ref).c)?.s || 0, { border }) })));
/** The border entry of every cell of a scene, as written. */
const written = (wb, refs) => {
  const xml = strFromU8(unzipSync(wb.toBytes())['xl/styles.xml']);
  const body = xml.match(/<borders[^>]*>([\s\S]*?)<\/borders>/)[1];
  const list = body.match(/<border(?:\s[^>]*?)?(?:\/>|>[\s\S]*?<\/border>)/g);
  return Object.fromEntries(refs.map((ref) => {
    const { r, c } = cell(ref);
    return [ref, list[xfOf(wb.styles, wb.sheet(0).cells.get(r)?.get(c)?.s || 0).borderId]];
  }));
};
const compare = (wb, scene) => {
  const want = excel.scenes[scene];
  const got = written(wb, Object.keys(want));
  return Object.keys(want).filter((ref) => got[ref] !== want[ref]).map((ref) => `${scene} ${ref}: Excel ${want[ref]}, here ${got[ref]}`);
};
const draw = (wb, range, mode, opts) => wb.setStyles(0, drawLines(wb, 0, range, mode, opts));

test('Draw Border outlines the range dragged, the neighbours\' facing lines taken away (Excel: Outside Borders on H2:J4)', () => {
  const wb = open();
  setUp(wb, [['H1', { bottom: { style: 'double', color: auto } }], ['K3', { left: { style: 'medium', color: { rgb: 'FFFF0000' } } }], ['G3', { right: { style: 'thin', color: auto } }]]);
  draw(wb, area('H2', 'J4'), 'draw');
  assert.deepEqual(compare(wb, 'msoOutline'), []);
});

test('Draw Border Grid draws every line in the range (Excel: All Borders on H8:J10)', () => {
  const wb = open();
  setUp(wb, [['K9', { left: { style: 'medium', color: { rgb: 'FFFF0000' } } }], ['H7', { bottom: { style: 'thin', color: auto } }]]);
  draw(wb, area('H8', 'J10'), 'grid');
  assert.deepEqual(compare(wb, 'msoGrid'), []);
});

test('Erase Border takes every line off, the diagonal and the neighbours\' facing lines too (Excel: No Border on N14:P16)', () => {
  const wb = open();
  const thin = { style: 'thin', color: auto };
  const cells = [];
  for (let r = 13; r <= 15; r++) for (let c = 13; c <= 15; c++) cells.push([`${String.fromCharCode(65 + c)}${r + 1}`, { left: thin, right: thin, top: thin, bottom: thin }]);
  setUp(wb, cells);
  setUp(wb, [['O15', { diagonalDown: true, diagonal: thin }], ['M15', { right: { style: 'medium', color: auto } }], ['Q15', { left: { style: 'medium', color: auto } }],
    ['N13', { bottom: { style: 'double', color: auto } }]]);
  draw(wb, area('N14', 'P16'), 'erase');
  assert.deepEqual(compare(wb, 'msoErase'), []);
});

test('beside one cell a neighbour\'s same line stays; beside a range of several it goes (Excel: H24, H20:I20)', () => {
  const wb = open();
  const thin = { style: 'thin', color: auto };
  setUp(wb, [['H19', { bottom: thin }], ['I19', { bottom: { style: 'medium', color: auto } }], ['G24', { right: thin }], ['I24', { left: thin }]]);
  draw(wb, area('H20', 'I20'), 'draw', { side: null });
  draw(wb, area('H24', 'H24'), 'draw');
  // (Top Border on H20:I20 is the top half of an outline: here the outline, so H20:I20's own other sides differ.)
  const got = written(wb, ['H19', 'I19', 'G24', 'H24', 'I24']);
  const want = excel.scenes.msoTop;
  for (const ref of ['H19', 'I19', 'G24', 'H24', 'I24']) assert.equal(got[ref], want[ref], ref);
});

test('a gridline clicked: just that edge drawn or erased, with the line chosen', () => {
  const wb = open();
  const dashed = { style: 'dashed', color: { rgb: 'FF123456' } };
  draw(wb, area('C3', 'C3'), 'draw', { side: 'bottom', line: dashed });
  const b = wb.styles.borders[xfOf(wb.styles, wb.sheet(0).cells.get(2).get(2).s).borderId];
  assert.deepEqual([b.bottom.style, b.bottom.color.rgb, b.top.style ?? null, b.left.style ?? null, b.right.style ?? null], ['dashed', 'FF123456', null, null, null]);
  // (Erasing the line from the cell below takes the line over it, which is C3's.)
  draw(wb, area('C4', 'C4'), 'erase', { side: 'top' });
  const after = wb.styles.borders[xfOf(wb.styles, wb.sheet(0).cells.get(2).get(2).s).borderId];
  assert.equal(after.bottom.style ?? null, null);
});

test('the thirteen line styles and a custom colour are written as Excel writes them', () => {
  const STYLES = ['hair', 'dotted', 'dashDotDot', 'dashDot', 'dashed', 'thin', 'mediumDashDotDot', 'slantDashDot', 'mediumDashDot', 'mediumDashed', 'medium', 'thick', 'double'];
  const wb = open();
  STYLES.forEach((style, k) => draw(wb, area(`A${20 + k}`, `A${20 + k}`), 'draw', { side: 'bottom', line: { style } }));
  const got = written(wb, STYLES.map((_, k) => `A${20 + k}`));
  const bottom = (x) => x.match(/<bottom[\s\S]*?(?:\/>|<\/bottom>)/)[0];
  for (let k = 0; k < STYLES.length; k++) assert.equal(bottom(got[`A${20 + k}`]), bottom(excel.scenes.styles[`A${20 + k}`]), STYLES[k]);
  // More Colors' RGB 18, 52, 86 as a line, a fill and a text colour (C20, C21, C22).
  const rgb = rgbOf('#123456');
  assert.equal(rgb, 'FF123456');
  draw(wb, area('C20', 'C20'), 'draw', { side: 'bottom', line: { style: 'thin', color: { rgb } } });
  assert.equal(written(wb, ['C20']).C20, excel.scenes.custom.C20.border);
  const s = wb.styleWith(0, { fill: { rgb } });
  const f = wb.styleWith(0, { font: { color: { rgb } } });
  wb.setStyles(0, [{ r: 20, c: 2, s }, { r: 21, c: 2, s: f }]);
  const xml = strFromU8(unzipSync(wb.toBytes())['xl/styles.xml']);
  assert.ok(xml.includes(excel.scenes.custom.C21.fill), 'the fill as Excel wrote it');
  assert.match(excel.scenes.custom.C22.font, /<color rgb="FF123456"\/>/);
  assert.match(xml.match(/<fonts[\s\S]*?<\/fonts>/)[0], /<font>(?:(?!<\/font>)[\s\S])*<color rgb="FF123456"\/>/);
});

test('a colour typed: six hex digits with or without # or FF, anything else refused', () => {
  assert.deepEqual(['123456', '#abcdef', 'FF00FF00', ' #0a0B0c '].map(rgbOf), ['FF123456', 'FFABCDEF', 'FF00FF00', 'FF0A0B0C']);
  assert.deepEqual(['12345', '#GG0000', '80123456', '', null].map(rgbOf), [null, null, null, null, null]);
});

test('Recent Colors: kept in styles.xml as Excel kept them, latest first, once each, at most ten', () => {
  const wb = open();
  assert.deepEqual(recentColors(wb), []);
  addRecentColor(wb, 'FFABCDEF');
  addRecentColor(wb, 'ff123456');
  const xml = strFromU8(unzipSync(wb.toBytes())['xl/styles.xml']);
  assert.equal(xml.match(/<colors>[\s\S]*?<\/colors>/)[0], excel.mruAfterSave);
  // (After dxfs and tableStyles, before extLst, by the schema's order.)
  assert.ok(!/<\/colors>[\s\S]*<(dxfs|tableStyles)[ >]/.test(xml));
  const again = Workbook.fromBytes(wb.toBytes());
  assert.deepEqual(recentColors(again), [{ rgb: 'FF123456' }, { rgb: 'FFABCDEF' }]);
  addRecentColor(again, 'FFABCDEF');
  assert.deepEqual(recentColors(again).map((c) => c.rgb), ['FFABCDEF', 'FF123456']);
  for (let k = 0; k < 12; k++) addRecentColor(again, `FF0000${k.toString(16).padStart(2, '0')}`);
  assert.equal(recentColors(again).length, MAX_RECENT);
  assert.equal(recentColors(again)[0].rgb, 'FF00000B');
  fs.mkdirSync('test/out/formats', { recursive: true });
  fs.writeFileSync('test/out/formats/drawborders.xlsx', again.toBytes());
  const out = spawnSync(process.execPath, ['tools/check-ooxml.mjs', 'test/out/formats/drawborders.xlsx'], { encoding: 'utf8' });
  assert.equal(out.status, 0, out.stdout + out.stderr);
});

test('the three modes as Excel\'s buttons: one at a time, pressed again it ends, a selection ends it', () => {
  const classes = new Set();
  const state = { onSelect: new Set(), ribbon: null, readOnly: false, workbook: null };
  const ran = [];
  const d = borderDrawing({ root: { classList: { toggle: (k, on) => (on ? classes.add(k) : classes.delete(k)) } }, state, editing: () => ({ run: (...a) => { ran.push(a); return true; } }) });
  const NAME = { BorderDrawLine: 'draw', BorderDrawGrid: 'grid', BorderErase: 'erase' };
  const MSO = { draw: 'BorderDrawLine', grid: 'BorderDrawGrid', erase: 'BorderErase', null: '' };
  for (const [step, want] of Object.entries(excel.modes)) {
    const what = step.replace(/^\d+: /, '');
    if (what === 'start') { /* nothing on */ } else if (what === 'select') for (const f of state.onSelect) f();
    else if (what === 'value') { /* a value typed leaves the mode */ } else d.act('drawBorderMode', NAME[what]);
    assert.equal(MSO[d.act('drawing')], want, step);
  }
  d.act('drawBorderMode', 'erase');
  assert.ok(classes.has('oxe-erasing-borders'));
  d.act('drawn', { r0: 1, c0: 1, r1: 2, c1: 2 }, null);
  assert.deepEqual(ran.at(-1), ['drawBorders', 'erase', { r0: 1, c0: 1, r1: 2, c1: 2 }, { side: null, line: { style: 'thin' } }]);
  assert.equal(d.act('stopDrawing'), true);
  assert.equal(d.act('drawing'), null);
  // (A line style or colour picked with no mode on turns Draw Border on.)
  d.act('borderLineStyle', 'mediumDashed');
  d.act('borderLineColor', { rgb: 'FF00B050' });
  assert.equal(d.act('drawing'), 'draw');
  d.act('drawn', { r0: 0, c0: 0, r1: 0, c1: 0 }, 'left');
  assert.deepEqual(ran.at(-1)[3], { side: 'left', line: { style: 'mediumDashed', color: { rgb: 'FF00B050' } } });
  assert.equal(d.act('borderLineStyle', 'wavy'), false);
});
