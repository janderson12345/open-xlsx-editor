import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseRef, refName } from '../src/ooxml/refs.js';
import { formatCodeOf } from '../src/ooxml/styles.js';
import { resolveColor } from '../src/ooxml/theme.js';
import { pasteHtml, cssColor, styleChange } from '../src/ooxml/cliphtml.js';
import { readHtmlTable, cssUnescape, parseDecls } from '../src/ooxml/htmltable.js';
import { pasteText } from '../src/ooxml/clipops.js';

// HTML from other programs pasted as Excel pastes it: each case of tools/excel-compare/clip-html-cases.json (made-up
// tables in Excel's, Google Sheets', a browser's and Word's forms) was put on the clipboard and pasted at B2 by a
// hidden Excel (excel-clip-paste.ps1, giving test/clip-html-excel.json). Here the same HTML is pasted at B2 and each
// cell compared with Excel's: value, number format, font, fill, borders, alignment, wrapping, indent, merged range,
// and the column widths. Paste Values (PasteSpecial as text) reads the text alone.
const cases = JSON.parse(fs.readFileSync('tools/excel-compare/clip-html-cases.json', 'utf8'));
const excel = JSON.parse(fs.readFileSync('test/clip-html-excel.json', 'utf8').replace(new RegExp('^' + String.fromCharCode(0xfeff)), ''));

const COM_ERRORS = { '-2146826281': '#DIV/0!' };
const H = { 1: 'general', '-4131': 'left', '-4108': 'center', '-4152': 'right', 7: 'centerContinuous', '-4130': 'justify', 5: 'fill' };
const V = { '-4160': 'top', '-4108': 'center', '-4107': 'bottom', '-4130': 'justify' };
/** COM's line style and weight as the file's border style. */
const sideStyle = ({ style, weight }) => (style === -4119 ? 'double' : style === -4118 ? 'dotted' : style === -4115 ? (weight === -4138 ? 'mediumDashed' : 'dashed')
  : weight === 1 ? 'hair' : weight === 2 ? 'thin' : weight === -4138 ? 'medium' : 'thick');
/** A number format as Excel reads it back (NumberFormat leaves out the quotes and backslashes the file has). */
const plainCode = (code) => code.replace(/\\(.)/g, '$1').replace(/"([^"]*)"/g, '$1');

function fresh() {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/formulas.xlsx')));
  const si = wb.addSheet(wb.sheets.length);
  return { wb, si };
}

function apply(wb, steps) {
  for (const x of steps) {
    wb.apply(x.si, x.contents);
    wb.setMerges(x.si, x.merges);
    if (x.styles.length) wb.setStyles(x.si, x.styles);
  }
}

/** Our cell as excel-clip-paste.ps1 reads Excel's. */
function ours(wb, si, r, c) {
  const sheet = wb.sheet(si);
  const cell = sheet.cells.get(r)?.get(c);
  const xf = wb.styles.cellXfs[cell?.s || 0];
  const font = wb.styles.fonts[xf.fontId];
  const fill = wb.styles.fills[xf.fillId];
  const hex = (color, auto) => (color && !color.auto && color.indexed !== 64 ? (resolveColor(color, wb.theme)?.replace('#', '') || auto) : auto)?.toUpperCase() ?? null;
  const merge = sheet.merges.find((m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1);
  // (A side as COM reads it: this cell's, or the neighbour's facing one.)
  const sideOf = (rr, cc, name) => {
    const x = wb.styles.cellXfs[sheet.cells.get(rr)?.get(cc)?.s || 0];
    const b = wb.styles.borders[x.borderId][name];
    return b?.style ? { style: b.style, color: hex(b.color, '000000') } : null;
  };
  const border = {
    left: sideOf(r, c, 'left') || (c > 0 ? sideOf(r, c - 1, 'right') : null),
    top: sideOf(r, c, 'top') || (r > 0 ? sideOf(r - 1, c, 'bottom') : null),
    bottom: sideOf(r, c, 'bottom') || sideOf(r + 1, c, 'top'),
    right: sideOf(r, c, 'right') || sideOf(r, c + 1, 'left'),
  };
  const fontOf = (f) => ({ name: f.name === 'Calibri' ? 'Aptos Narrow' : f.name, size: f.size, bold: f.b, italic: f.i, underline: f.u === 'single' ? 2 : f.u ? 5 : -4142, strike: f.strike, color: hex(f.color, '000000') });
  const out = {
    value: cell?.v ?? null,
    numFmt: plainCode(formatCodeOf(wb.styles, cell?.s || 0)),
    font: fontOf(font),
    fill: fill.pattern === 'solid' ? hex(fill.fg, null) : null,
    h: xf.align.horizontal, v: xf.align.vertical, wrap: xf.align.wrap, indent: xf.align.indent,
    merge: merge ? `${refName(merge.r0, merge.c0)}:${refName(merge.r1, merge.c1)}` : null,
    border,
  };
  if (cell?.runs) {
    // (Rich text: each run's font, the cell's where it says nothing; a property the runs differ in reads 'mixed', as
    // COM reads it.)
    out.runs = cell.runs.map((x) => ({ text: x.text, font: fontOf(x.font ? Object.fromEntries(Object.entries({ ...font, ...x.font }).map(([k, v]) => [k, v ?? font[k]])) : font) }));
    for (const k of Object.keys(out.font)) if (new Set(out.runs.map((x) => x.font[k])).size > 1) out.font[k] = 'mixed';
  }
  return out;
}

/** Excel's cell in the same terms. */
function theirs(x) {
  const border = Object.fromEntries(Object.entries(x.border).map(([k, s]) => [k, s ? { style: sideStyle(s), color: s.color } : null]));
  const value = typeof x.value === 'number' && COM_ERRORS[x.value] ? COM_ERRORS[x.value] : x.value === '' ? null : x.value;
  return {
    value,
    numFmt: plainCode(x.numFmt),
    font: { ...x.font },
    fill: x.fill,
    h: H[x.h], v: V[x.v], wrap: x.wrap, indent: x.indent, merge: x.merge, border,
    ...(x.runs ? { runs: x.runs } : {}),
  };
}

for (const c of cases) {
  test(`HTML pasted as Excel pastes it: ${c.name}`, () => {
    const { wb, si } = fresh();
    const at = { r0: 1, c0: 1, r1: 1, c1: 1 };
    const steps = c.values ? pasteText(wb, si, at, c.text) : pasteHtml(wb, si, at, c.html);
    apply(wb, steps);
    const e = excel.find((x) => x.name === c.name);
    for (const [ref, cell] of Object.entries(e.cells)) {
      const { r, c: col } = parseRef(ref);
      const want = theirs(cell);
      const got = ours(wb, si, r, col);
      // (Excel's font colour for the automatic one reads 000000, as black does.)
      assert.deepEqual(got, want, `${c.name} ${ref}`);
    }
    // (Paste leaves the columns' widths as they were, in Excel as here.)
    const first = parseRef(c.read.split(':')[0]).c;
    e.widths.forEach((w, k) => assert.equal(wb.sheet(si).cols.find((x) => x.min <= first + k && x.max >= first + k && x.width != null)?.width ?? 8.43, w));
  });
}

test('Keep Source Column Widths: the table\'s col widths in characters', () => {
  const { wb, si } = fresh();
  const [x] = pasteHtml(wb, si, { r0: 1, c0: 1, r1: 1, c1: 1 }, cases[0].html, { widths: true });
  // (145px and 89px wide columns: 20 and 12 characters, as Excel set them in the copy, px = chars * 7 + 5.)
  assert.deepEqual(x.widths.slice(0, 3).map((w) => [w.c, w.width]), [[1, 20], [2, 9.71484375], [3, 12]]);
});

test('CSS pieces: colours, escapes, declarations', () => {
  assert.equal(cssColor('#fc0'), 'FFCC00');
  assert.equal(cssColor('rgb(221, 235, 247)'), 'DDEBF7');
  assert.equal(cssColor('rgba(0, 0, 0, 0)'), null);
  assert.equal(cssColor('windowtext'), 'auto');
  assert.equal(cssUnescape('\\0022$\\0022\\#\\,\\#\\#0\\.00'), '"$"#,##0.00');
  assert.deepEqual(parseDecls("font-family:'A;B', serif; color : red"), { 'font-family': "'A;B', serif", color: 'red' });
  assert.deepEqual(styleChange({ 'border-bottom': '2px dashed #f00', 'white-space': 'nowrap', 'vertical-align': 'bottom' }), { border: { bottom: { style: 'mediumDashed', color: { rgb: 'FFFF0000' } } } });
});

test('HTML without a table is one cell; unclosed cells and rows are closed', () => {
  assert.deepEqual(readHtmlTable('<p><b>Just</b> text</p>').rows.map((r) => r.map((x) => x.text)), [['Just text']]);
  const t = readHtmlTable('<table><tr><td>a<td>b<tr><td>c &amp; d<td>&#x45;</table>');
  assert.deepEqual(t.rows.map((r) => r.map((x) => x.text)), [['a', 'b'], ['c & d', 'E']]);
});

test('Long or odd HTML is read in linear time', () => {
  const big = `<table>${'<tr><td style="border:1px solid rgb(1,2,3)">x</td></tr>'.repeat(5000)}</table>`;
  const t0 = Date.now();
  assert.equal(readHtmlTable(big).rows.length, 5000);
  readHtmlTable(`<td a='${'"'.repeat(20000)}`);
  readHtmlTable('<'.repeat(40000));
  readHtmlTable('<a'.repeat(40000));
  // (Nested past what recursion could walk: read as siblings past 200 deep.)
  assert.equal(readHtmlTable(`<table><tr><td>${'<b>'.repeat(20000)}x`).rows[0][0].text, 'x');
  assert.ok(Date.now() - t0 < 3000);
});
