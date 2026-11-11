import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { Workbook } from '../src/ooxml/workbook.js';
import { xfOf } from '../src/ooxml/styles.js';
import { spawnSync } from 'node:child_process';

// What the Format Cells box sets, written as Excel writes it: each case of tools/excel-compare/excel-formatcells.ps1
// (Excel's saved styles.xml in test/formatcells-excel.json) made here on basic.xlsx's A1, its cell format, font, fill
// and border compared with Excel's text for text.
const excel = JSON.parse(fs.readFileSync('test/formatcells-excel.json', 'utf8'));
const red = { rgb: 'FFFF0000' };
const auto = { indexed: 64 };
const CHANGES = {
  rotate45: { align: { rotation: 45 } },
  rotateMinus90: { align: { rotation: 180 } },
  rotate90: { align: { rotation: 90 } },
  vertical: { align: { rotation: 255 } },
  shrink: { align: { shrink: true } },
  hFill: { align: { horizontal: 'fill' } },
  hJustify: { align: { horizontal: 'justify' } },
  hDistributed: { align: { horizontal: 'distributed' } },
  hCenterAcross: { align: { horizontal: 'centerContinuous' } },
  vJustify: { align: { vertical: 'justify' } },
  vDistributed: { align: { vertical: 'distributed' } },
  vTop: { align: { vertical: 'top' } },
  readingRtl: { align: { readingOrder: 2 } },
  gradientLinear: { fill: { gradient: { degree: 90, stops: [{ position: 0, color: { rgb: 'FFFFFFFF' } }, { position: 1, color: { rgb: 'FF0070C0' } }] } } },
  gradientLinear45: { fill: { gradient: { degree: 45, stops: [{ position: 0, color: red }, { position: 0.5, color: { rgb: 'FFFFFFFF' } }, { position: 1, color: red }] } } },
  gradientPath: { fill: { gradient: { type: 'path', left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, stops: [{ position: 0, color: { rgb: 'FFFFFFFF' } }, { position: 1, color: { rgb: 'FFFFC000' } }] } } },
  gradientTheme: { fill: { gradient: { degree: 0, stops: [{ position: 0, color: { theme: 0 } }, { position: 1, color: { theme: 4 } }] } } },
  patternGray50: { fill: { pattern: 'mediumGray' } },
  patternDown: { fill: { pattern: 'darkDown', fg: red, bg: { rgb: 'FFFFFF00' } } },
  patternGray125: { fill: { pattern: 'gray125', fg: { rgb: 'FF0000FF' } } },
  solidPatternColor: { fill: { rgb: 'FFFFFF00' } },
  borderRedDashed: { border: { bottom: { style: 'dashed', color: red } } },
  borderDiagDown: { border: { diagonalDown: true, diagonal: { style: 'thin', color: auto } } },
  borderDiagUp: { border: { diagonalUp: true, diagonal: { style: 'thin', color: auto } } },
  borderDiagBoth: { border: { diagonalUp: true, diagonalDown: true, diagonal: { style: 'thin', color: auto } } },
  borderLeftHair: { border: { left: { style: 'hair', color: auto } } },
  borderTopThick: { border: { top: { style: 'thick', color: { theme: 4 } } } },
  borderSlant: { border: { right: { style: 'slantDashDot', color: auto } } },
  fontSuper: { font: { vertAlign: 'superscript' } },
  fontSub: { font: { vertAlign: 'subscript' } },
  fontDoubleU: { font: { u: 'double' } },
  fontAccountingU: { font: { u: 'singleAccounting' } },
  fontDoubleAccU: { font: { u: 'doubleAccounting' } },
  unlocked: { protection: { locked: false } },
  hidden: { protection: { hidden: true } },
  unlockedHidden: { protection: { locked: false, hidden: true } },
  // (Excel took these one setting at a time; here each is one change, then a second.)
  rotateShrinkIndent: [{ align: { horizontal: 'left', indent: 2, shrink: true } }, { align: { rotation: 30 } }],
  rotateThenIndent: [{ align: { rotation: 30 } }, { align: { horizontal: 'left', indent: 2 } }],
  indentThenRotate: [{ align: { horizontal: 'left', indent: 2 } }, { align: { rotation: 30 } }],
  wrapThenShrink: [{ align: { wrap: true } }, { align: { shrink: true } }],
  shrinkThenWrap: [{ align: { shrink: true } }, { align: { wrap: true } }],
  fillThenRotate: [{ align: { horizontal: 'fill' } }, { align: { rotation: 45 } }],
  centerRotate: [{ align: { horizontal: 'center' } }, { align: { rotation: 45 } }],
};
const open = () => Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
/** The n-th entry of a list of styles.xml, as text. */
const entries = (xml, list, item) => {
  const body = xml.match(new RegExp(`<${list}[^>]*>([\\s\\S]*?)</${list}>`))[1];
  return body.match(new RegExp(`<${item}(?:\\s[^>]*?)?(?:/>|>[\\s\\S]*?</${item}>)`, 'g'));
};
/** An xf without its indexes into the lists and its number format, which differ between the two workbooks. */
const bare = (xf) => xf.replace(/ (numFmtId|fontId|fillId|borderId)="\d+"/g, '');
/** A font's change from the workbook's own (basic.xlsx's font is Calibri, Excel's new workbook's Aptos Narrow). */
const fontChange = (font, base) => font.replace(base.replace(/^<font>|<\/font>$/g, ''), '|');

test('each setting of the Format Cells box is written as Excel writes it (test/formatcells-excel.json)', () => {
  const wrong = [];
  const named = new Set(excel.cases.map((c) => c.name));
  assert.deepEqual(Object.keys(CHANGES).filter((k) => !named.has(k)), []);
  const excelBase = excel.cases.find((c) => c.name === 'rotate45').font;
  for (const want of excel.cases) {
    const wb = open();
    let s = 0;
    for (const change of [CHANGES[want.name]].flat()) s = wb.styleWith(s, change);
    wb.setStyles(0, [{ r: 0, c: 0, s }]);
    const xml = strFromU8(unzipSync(wb.toBytes())['xl/styles.xml']);
    const xf = xfOf(wb.styles, s);
    const fonts = entries(xml, 'fonts', 'font');
    const got = {
      xf: bare(entries(xml, 'cellXfs', 'xf')[s]),
      font: fontChange(fonts[xf.fontId], fonts[0]),
      fill: entries(xml, 'fills', 'fill')[xf.fillId],
      border: entries(xml, 'borders', 'border')[xf.borderId],
    };
    const exp = { xf: bare(want.xf), font: fontChange(want.font, excelBase), fill: want.fill, border: want.border };
    for (const k of Object.keys(exp)) if (got[k] !== exp[k]) wrong.push(`${want.name} ${k}: Excel ${exp[k]}, here ${got[k]}`);
  }
  assert.equal(excel.cases.length, Object.keys(CHANGES).length);
  assert.deepEqual(wrong, []);
});

test('read back: rotation, shrink, reading order, protection, gradients, patterns, diagonals and superscript', () => {
  const wb = open();
  const s = wb.styleWith(0, { align: { rotation: 255, shrink: true, readingOrder: 2 }, protection: { locked: false, hidden: true }, font: { vertAlign: 'superscript' },
    border: { diagonalUp: true, diagonal: { style: 'thin', color: auto } }, fill: CHANGES.gradientPath.fill });
  wb.setStyles(0, [{ r: 0, c: 0, s }]);
  const again = Workbook.fromBytes(wb.toBytes());
  const xf = xfOf(again.styles, again.sheet(0).cells.get(0).get(0).s);
  assert.deepEqual([xf.align.rotation, xf.align.shrink, xf.align.readingOrder], [255, true, 2]);
  assert.deepEqual(xf.protection, { locked: false, hidden: true });
  assert.equal(again.styles.fonts[xf.fontId].vertAlign, 'superscript');
  const b = again.styles.borders[xf.borderId];
  assert.deepEqual([b.diagonalUp, b.diagonalDown, b.diagonal.style], [true, false, 'thin']);
  const f = again.styles.fills[xf.fillId];
  assert.equal(f.pattern, 'gradient');
  assert.deepEqual([f.gradient.type, f.gradient.left, f.gradient.stops.length, f.gradient.stops[1].color.rgb], ['path', 0.5, 2, 'FFFFC000']);
  // (A cell format that says nothing is locked; turning both back on leaves no protection element.)
  assert.deepEqual(xfOf(again.styles, 0).protection, { locked: true, hidden: false });
  const back = again.styleWith(xfOf(again.styles, 0) && again.sheet(0).cells.get(0).get(0).s, { protection: { locked: true, hidden: false } });
  again.setStyles(0, [{ r: 0, c: 0, s: back }]);
  const xml = strFromU8(unzipSync(again.toBytes())['xl/styles.xml']);
  assert.doesNotMatch(entries(xml, 'cellXfs', 'xf')[back], /<protection/);
  fs.mkdirSync('test/out/formats', { recursive: true });
  fs.writeFileSync('test/out/formats/formatcells.xlsx', again.toBytes());
  const out = spawnSync(process.execPath, ['tools/check-ooxml.mjs', 'test/out/formats/formatcells.xlsx'], { encoding: 'utf8' });
  assert.equal(out.status, 0, out.stdout + out.stderr);
});

test('the Border tab over a range: the outline on its edge, the inside lines between its cells (Excel: D2:F4)', async () => {
  const { formatCells } = await import('../src/ooxml/formatops.js');
  const wb = open();
  const si = 0;
  const thin = { style: 'thin', color: auto };
  const sel = { r0: 1, c0: 3, r1: 3, c1: 5, ar: 1, ac: 3 };
  wb.setStyles(si, formatCells(wb, si, sel, { borders: { outline: { left: thin, right: thin, top: thin, bottom: thin }, inside: { horizontal: { style: 'dashed', color: auto }, vertical: { style: 'dashed', color: auto } } } }));
  const sides = (xml) => ['left', 'right', 'top', 'bottom'].map((k) => (xml.match(new RegExp(`<${k} style="([a-zA-Z]+)"`)) || ['', '-'])[1]).join(' ');
  const wrong = [];
  for (const [ref, want] of Object.entries(excel.range)) {
    const r = Number(ref.slice(1)) - 1;
    const c = ref.charCodeAt(0) - 65;
    if (r < 1 || r > 3 || c < 3 || c > 5) continue; // (The neighbours: see docs/FORMATS.md.)
    const b = wb.styles.borders[xfOf(wb.styles, wb.sheet(si).cells.get(r)?.get(c)?.s || 0).borderId];
    const got = ['left', 'right', 'top', 'bottom'].map((k) => b[k].style || '-').join(' ');
    if (got !== sides(want)) wrong.push(`${ref}: Excel ${sides(want)}, here ${got}`);
  }
  assert.deepEqual(wrong, []);
});

test('rows for stacked text: a line a letter and 1.5pt (Excel: ab 31.5, Stack 76.5 in Calibri 11)', async () => {
  const { turnedPt } = await import('../src/editor/rowheights.js');
  assert.deepEqual([turnedPt('ab', 255, '', 15), turnedPt('Stack', 255, '', 15), turnedPt('Hello world text', 255, '', 15)], [31.5, 76.5, 241.5]);
});
