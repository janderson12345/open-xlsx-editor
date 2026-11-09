import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { Workbook } from '../src/ooxml/workbook.js';
import { xfOf, formatCodeOf } from '../src/ooxml/styles.js';
import { parseRef } from '../src/ooxml/refs.js';
import { toggleFont, stepFont, stepSize, toggleWrap, stepIndent, setBorders, stepDecimals, formatEach } from '../src/ooxml/formatops.js';
import { spawnSync } from 'node:child_process';

// What Excel's own Home buttons did over COM (CommandBars.ExecuteMso on a new workbook; docs/FORMATS.md) is what
// these expect.
const open = (name = 'basic.xlsx') => Workbook.fromBytes(new Uint8Array(fs.readFileSync(`test/fixtures/${name}`)));
const range = (text, active = null) => {
  const [a, b = a] = text.split(':').map(parseRef);
  const act = active ? parseRef(active) : a;
  return { r0: a.r, c0: a.c, r1: b.r, c1: b.c, ar: act.r, ac: act.c };
};
const at = (wb, ref, si = 0) => { const { r, c } = parseRef(ref); return wb.sheet(si).cells.get(r)?.get(c)?.s || 0; };
const fontAt = (wb, ref) => wb.styles.fonts[xfOf(wb.styles, at(wb, ref)).fontId];
const sides = (wb, ref) => {
  const b = wb.styles.borders[xfOf(wb.styles, at(wb, ref)).borderId];
  return ['left', 'top', 'bottom', 'right'].filter((k) => b[k].style).map((k) => `${k[0].toUpperCase()}:${b[k].style}`).join(' ');
};
const run = (wb, cells) => wb.setStyles(0, cells);
const dir = 'test/out/formats';
fs.mkdirSync(dir, { recursive: true });
/** The structural checker's verdict on saved bytes (tools/check-ooxml.mjs): '' when it finds nothing. */
const check = (name, bytes) => {
  fs.writeFileSync(`${dir}/${name}.xlsx`, bytes);
  const out = spawnSync(process.execPath, ['tools/check-ooxml.mjs', `${dir}/${name}.xlsx`], { encoding: 'utf8' });
  return out.status === 0 ? '' : out.stdout + out.stderr;
};

test('toggles follow the active cell: on for all unless it has it', () => {
  const wb = open();
  run(wb, toggleFont(wb, 0, range('E5:F5', 'F5'), 'b'));
  assert.deepEqual([fontAt(wb, 'E5').b, fontAt(wb, 'F5').b], [true, true]);
  run(wb, toggleFont(wb, 0, range('E5:G5', 'E5'), 'b')); // (E5 bold: all off, G5 too)
  assert.deepEqual([fontAt(wb, 'E5').b, fontAt(wb, 'F5').b, fontAt(wb, 'G5').b], [false, false, false]);
  run(wb, toggleFont(wb, 0, range('E6'), 'u'));
  assert.equal(fontAt(wb, 'E6').u, 'single');
  run(wb, toggleWrap(wb, 0, range('E7')));
  assert.equal(xfOf(wb.styles, at(wb, 'E7')).align.wrap, true);
  run(wb, toggleWrap(wb, 0, range('E7')));
  assert.equal(xfOf(wb.styles, at(wb, 'E7')).align.wrap, false);
});

test('a format already there is used again; a new one goes at the end', () => {
  const wb = open();
  const before = wb.styles.cellXfs.length;
  run(wb, toggleFont(wb, 0, range('H1:H3'), 'i'));
  assert.equal(wb.styles.cellXfs.length, before + 1);
  assert.equal(at(wb, 'H1'), before);
  run(wb, toggleFont(wb, 0, range('H4'), 'i'));
  assert.equal(at(wb, 'H4'), before);
});

test('Grow and Shrink Font step through Excel\'s sizes from the active cell', () => {
  const seq = (from, dir, n) => { const out = []; let s = from; for (let k = 0; k < n; k++) out.push(s = stepSize(s, dir)); return out; };
  assert.deepEqual(seq(11, 1, 14), [12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72, 72, 72]);
  assert.deepEqual(seq(11, -1, 6), [10, 9, 8, 8, 8, 8]);
  assert.equal(stepSize(13, 1), 14);
  const wb = open();
  run(wb, formatEach(wb, 0, range('J2'), { font: { size: 20 } }));
  run(wb, stepFont(wb, 0, range('J1:J2', 'J1'), 1)); // (Excel: 11 and 20 both become 12)
  assert.deepEqual([fontAt(wb, 'J1').size, fontAt(wb, 'J2').size], [12, 12]);
});

test('indenting makes a cell left-aligned unless it is right-aligned', () => {
  const wb = open();
  run(wb, formatEach(wb, 0, range('K2'), { align: { horizontal: 'center' } }));
  run(wb, formatEach(wb, 0, range('K3'), { align: { horizontal: 'right' } }));
  run(wb, stepIndent(wb, 0, range('K1:K3'), 1));
  const a = (ref) => xfOf(wb.styles, at(wb, ref)).align;
  assert.deepEqual([a('K1'), a('K2'), a('K3')].map((x) => `${x.horizontal} ${x.indent}`), ['left 1', 'left 1', 'right 1']);
  run(wb, stepIndent(wb, 0, range('K1'), -1));
  assert.equal(`${a('K1').horizontal} ${a('K1').indent}`, 'left 0');
});

test('border presets on B2:C3, as Excel stores them', () => {
  const cases = {
    bottom: ['', '', 'B:thin', 'B:thin'],
    outside: ['L:thin T:thin', 'T:thin R:thin', 'L:thin B:thin', 'B:thin R:thin'],
    all: Array(4).fill('L:thin T:thin B:thin R:thin'),
    thickOutside: ['L:medium T:medium', 'T:medium R:medium', 'L:medium B:medium', 'B:medium R:medium'],
    doubleBottom: ['', '', 'B:double', 'B:double'],
    topThickBottom: ['T:thin', 'T:thin', 'B:medium', 'B:medium'],
  };
  for (const [preset, want] of Object.entries(cases)) {
    const wb = open('formulas.xlsx');
    run(wb, setBorders(wb, 0, range('B2:C3'), preset));
    assert.deepEqual(['B2', 'C2', 'B3', 'C3'].map((r) => sides(wb, r)), want, preset);
  }
});

test('a border takes the neighbour\'s facing side away unless it is the same line', () => {
  const cases = [
    // [neighbour, its side, its line, preset on B2, the neighbour after]
    ['B1', 'bottom', 'medium', 'top', ''],
    ['B1', 'bottom', 'thin', 'top', 'B:thin'],
    ['A2', 'right', 'thick', 'left', ''],
    ['B1', 'bottom', 'medium', 'none', ''],
    ['C2', 'left', 'thin', 'none', ''],
  ];
  for (const [ref, side, style, preset, want] of cases) {
    const wb = open('formulas.xlsx');
    run(wb, formatEach(wb, 0, range(ref), { border: { [side]: { style } } }));
    run(wb, setBorders(wb, 0, range('B2'), preset));
    assert.equal(sides(wb, ref), want, `${preset} after ${ref} ${side} ${style}`);
  }
});

test('Increase and Decrease Decimal, as Excel changes the format', () => {
  assert.equal(stepDecimals('General', 1, '1.5'), '0.00');
  assert.equal(stepDecimals('0.00', 1, null), '0.000');
  assert.equal(stepDecimals('General', -1, '2.123457'), '0.00000');
  assert.equal(stepDecimals('0.00%', -1), '0.0%');
  assert.equal(stepDecimals('0.0%', -1), '0%');
  assert.equal(stepDecimals('#,##0.00 "kg";[Red]-#,##0.00', 1), '#,##0.000 "kg";[Red]-#,##0.000');
  assert.equal(stepDecimals('0', 1), '0.0');
  assert.equal(stepDecimals('m/d/yyyy', 1), null);
  assert.equal(stepDecimals('General', 1, 'text'), null);
});

test('Percent, Comma and Currency are Excel\'s built-in cell styles', () => {
  const wb = open();
  run(wb, formatEach(wb, 0, range('L1'), { style: 'Percent' }));
  run(wb, formatEach(wb, 0, range('L2'), { style: 'Comma' }));
  run(wb, formatEach(wb, 0, range('L3'), { style: 'Currency' }));
  run(wb, formatEach(wb, 0, range('L4'), { style: 'Percent' }));
  assert.equal(formatCodeOf(wb.styles, at(wb, 'L1')), '0%');
  assert.equal(formatCodeOf(wb.styles, at(wb, 'L2')), '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)');
  assert.equal(at(wb, 'L4'), at(wb, 'L1'));
  const files = unzipSync(wb.toBytes());
  const xml = strFromU8(files['xl/styles.xml']);
  for (const [name, id] of [['Percent', 5], ['Comma', 3], ['Currency', 4]]) assert.match(xml, new RegExp(`<cellStyle name="${name}" xfId="\\d+" builtinId="${id}"/>`));
  assert.match(xml, /<numFmt numFmtId="44" formatCode="_\(&quot;\$&quot;\* #,##0\.00_\)/);
  assert.equal(xml.match(/builtinId="5"/g).length, 1);
});

test('a formatted workbook saves and reads back; one without a styles part gets one', () => {
  const wb = open();
  run(wb, toggleFont(wb, 0, range('M1'), 'b'));
  run(wb, formatEach(wb, 0, range('M2'), { fill: { rgb: 'FFFFFF00' }, font: { color: { rgb: 'FFFF0000' } }, numFmt: '0.000' }));
  const bytes = wb.toBytes();
  assert.equal(check('formatted', bytes), '');
  const back = Workbook.fromBytes(bytes);
  assert.equal(fontAt(back, 'M1').b, true);
  assert.equal(back.styles.fills[xfOf(back.styles, at(back, 'M2')).fillId].fg.rgb, 'FFFFFF00');
  assert.equal(formatCodeOf(back.styles, at(back, 'M2')), '0.000');

  const files = unzipSync(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  delete files['xl/styles.xml'];
  files['xl/_rels/workbook.xml.rels'] = strToU8(strFromU8(files['xl/_rels/workbook.xml.rels']).replace(/<Relationship [^>]*styles"[^>]*\/>/, ''));
  files['[Content_Types].xml'] = strToU8(strFromU8(files['[Content_Types].xml']).replace(/<Override [^>]*styles\.xml"[^>]*\/>/, ''));
  const bare = Workbook.fromBytes(zipSync(files));
  bare.setStyles(0, toggleFont(bare, 0, range('A1'), 'b'));
  const saved = bare.toBytes();
  assert.equal(check('no-styles-part', saved), '');
  assert.equal(fontAt(Workbook.fromBytes(saved), 'A1').b, true);
});
