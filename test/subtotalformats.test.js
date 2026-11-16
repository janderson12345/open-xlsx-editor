import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { serializeXml } from '../src/ooxml/xml.js';
import { subtotal } from '../src/ooxml/subtotal.js';

// Subtotal over formatted cells: test/outline-excel.json holds what Excel wrote (tools/excel-compare/excel-outline.ps1,
// "subtotal formatted ..."; Range.Subtotal over a list with number formats, a fill, italic, a bold 30pt row, bold
// headings). The same list here, the same Subtotal, and every cell's and row's format compared with Excel's by what it
// is: its number format, bold, italic and fill colour, and the row's own format and height.
const excel = JSON.parse(fs.readFileSync('test/outline-excel.json', 'utf8'));
const byName = new Map(excel.cases.map((c) => [c.name, c]));
const LIST = [['Region', 'Item', 'Units', 'Price'], ['East', 'pear', 4, 1.5], ['East', 'fig', 2, 3], ['North', 'kiwi', 7, 0.5], ['West', 'plum', 1, 2.25], ['West', 'pear', 5, 1.5], ['West', 'lime', 3, 0.75]];

function list() {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'List');
  wb.apply(si, LIST.flatMap((row, r) => row.map((value, c) => ({ r, c, value }))));
  return { wb, si };
}
const range = (r0, r1, c0, c1, s) => { const out = []; for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) out.push({ r, c, s }); return out; };
/** The probe's Formats: Units 0.00, Price "$"#,##0.00, Item italic, Region filled yellow, the headings bold. */
function formats(wb, si) {
  wb.setStyles(si, [...range(1, 6, 2, 2, wb.styleWith(0, { numFmt: '0.00' })), ...range(1, 6, 3, 3, wb.styleWith(0, { numFmt: '"$"#,##0.00' })),
    ...range(1, 6, 1, 1, wb.styleWith(0, { font: { i: true } })), ...range(1, 6, 0, 0, wb.styleWith(0, { fill: { pattern: 'solid', fg: { rgb: 'FFFFFF00' } } })),
    ...range(0, 0, 0, 3, wb.styleWith(0, { font: { b: true } }))]);
}
/** A whole row given a format, as Rows(n).Font does: the row's own and its cells'. */
function rowFormat(wb, si, r, change) {
  const s = wb.styleWith(0, change);
  wb.setStyles(si, range(r, r, 0, 3, s));
  wb.setLineStyles(si, { rows: [{ r, s }] });
}
const ST = (below) => (wb, si) => subtotal(wb, si, { r0: 0, c0: 0, r1: 6, c1: 3 }, { by: 0, fn: 'sum', cols: [2, 3], below });
const STEPS = {
  'subtotal formatted cells': [formats, ST(true)],
  'subtotal formatted cells above': [formats, ST(false)],
  'subtotal formatted row': [(wb, si) => { rowFormat(wb, si, 2, { font: { b: true } }); wb.setLines(si, 'row', 2, 2, { height: 30 }); wb.setStyles(si, [{ r: 4, c: 2, s: wb.styleWith(0, { fill: { pattern: 'solid', fg: { rgb: 'FFFF0000' } } }) }]); },
    (wb, si) => subtotal(wb, si, { r0: 0, c0: 0, r1: 6, c1: 3 }, { by: 0, fn: 'sum', cols: [2] })],
  'subtotal formatted row above': [(wb, si) => { rowFormat(wb, si, 0, { font: { b: true } }); rowFormat(wb, si, 1, { font: { i: true } }); },
    (wb, si) => subtotal(wb, si, { r0: 0, c0: 0, r1: 6, c1: 3 }, { by: 0, fn: 'sum', cols: [2], below: false })],
};

const attrsOf = (tag) => Object.fromEntries([...tag.matchAll(/ ([\w:]+)="([^"]*)"/g)].map((m) => [m[1], m[2]]));
const items = (xml, tag) => [...(xml || '').matchAll(new RegExp(`<${tag}\\b[^>]*?(?:/>|>[\\s\\S]*?</${tag}>)`, 'g'))].map((m) => m[0]);
const BUILTIN = { 0: 'General', 2: '0.00' };
/** Format s described: number format, bold, italic, fill colour; '' for the plain one. */
function describer({ numFmts, fonts, fills, cellXfs }) {
  const codes = Object.fromEntries(items(numFmts, 'numFmt').map(attrsOf).map((a) => [a.numFmtId, a.formatCode.replace(/&quot;/g, '"')]));
  const fontList = items(fonts.replace(/^<fonts[^>]*>/, ''), 'font');
  const fillList = items(fills.replace(/^<fills[^>]*>/, ''), 'fill');
  const xfs = items(cellXfs.replace(/^<cellXfs[^>]*>/, ''), 'xf').map(attrsOf);
  return (s) => {
    const xf = xfs[Number(s || 0)];
    const font = fontList[Number(xf.fontId || 0)] || '';
    const fill = (fillList[Number(xf.fillId || 0)] || '').match(/<fgColor rgb="(\w+)"/);
    const parts = [codes[xf.numFmtId] ?? BUILTIN[xf.numFmtId || 0] ?? xf.numFmtId, /<b\/>/.test(font) ? 'bold' : '', /<i\/>/.test(font) ? 'italic' : '', fill ? fill[1] : ''];
    const text = parts.filter((x) => x && x !== 'General').join(' ');
    return text;
  };
}
function sheetFormats(sheetData, describe) {
  const out = {};
  for (const m of sheetData.matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
    const a = attrsOf(m[1]);
    if (a.customFormat === '1' && describe(a.s)) out[`row ${a.r}`] = describe(a.s);
    if (a.ht) out[`row ${a.r} height`] = a.ht;
    for (const c of (m[2] || '').matchAll(/<c\b([^>]*?)(?:\/>|>)/g)) { const ca = attrsOf(c[1]); if (describe(ca.s)) out[ca.r] = describe(ca.s); }
  }
  return out;
}

for (const [name, [fill, steps]] of Object.entries(STEPS)) {
  test(`Subtotal's formats as Excel gives them: ${name}`, () => {
    const want = byName.get(name);
    assert.ok(want, `no measurement for ${name}`);
    const { wb, si } = list();
    fill(wb, si);
    steps(wb, si);
    const stylesXml = serializeXml(wb.stylesDoc);
    const pick = (re) => (stylesXml.match(re) || [''])[0];
    const ours = describer({ numFmts: pick(/<numFmts\b[\s\S]*?<\/numFmts>/), fonts: pick(/<fonts\b[\s\S]*?<\/fonts>/), fills: pick(/<fills\b[\s\S]*?<\/fills>/), cellXfs: pick(/<cellXfs\b[\s\S]*?<\/cellXfs>/) });
    const expected = sheetFormats(want.sheetData, describer(want));
    assert.ok(Object.keys(expected).length > 8, 'Excel formatted cells');
    assert.deepEqual(sheetFormats((serializeXml(wb._part(si).doc).match(/<sheetData>[\s\S]*?<\/sheetData>/) || [''])[0], ours), expected);
  });
}
