import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { serializeXml } from '../src/ooxml/xml.js';
import { groupLines, setSummary, autoOutline } from '../src/ooxml/outline.js';
import { applyOutlineStyles } from '../src/ooxml/outlinestyles.js';
import { subtotal } from '../src/ooxml/subtotal.js';

// The outline's automatic styles: test/outline-excel.json holds what Excel wrote (tools/excel-compare/excel-outline.ps1:
// Outline.AutomaticStyles, Range.ApplyOutlineStyles, AutoOutline and Subtotal with the styles on, Group with them on).
// The same steps here, and each cell's and row's format compared with Excel's by what it is: the cell style it is based
// on (RowLevel_1, ColLevel_1...: name, builtinId, iLevel), its font's bold and italic, and the row's own format; with
// sheetPr's outlinePr.
const excel = JSON.parse(fs.readFileSync('test/outline-excel.json', 'utf8'));
const byName = new Map(excel.cases.map((c) => [c.name, c]));

function book(fill) {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'Styles');
  fill(wb, si);
  return { wb, si };
}
/** Excel's probe cells: numbers down A1:A12 and across B1:H1. */
const numbers = (wb, si) => wb.apply(si, [...Array.from({ length: 12 }, (_, r) => ({ r, c: 0, value: (r + 1) * 10 })), ...Array.from({ length: 7 }, (_, k) => ({ r: 0, c: k + 1, value: k + 2 }))]);
/** 3 rows of numbers in A and B, SUMs under them in row 4 and beside them in C. */
const sums = (wb, si) => {
  const cells = [];
  for (let r = 0; r < 3; r++) for (let c = 0; c < 2; c++) cells.push({ r, c, value: r + c + 2 });
  cells.push({ r: 3, c: 0, formula: 'SUM(A1:A3)' }, { r: 3, c: 1, formula: 'SUM(B1:B3)' });
  for (let r = 0; r < 4; r++) cells.push({ r, c: 2, formula: `SUM(A${r + 1}:B${r + 1})` });
  wb.apply(si, cells);
};
const list = (wb, si) => wb.apply(si, [['Region', 'Item', 'Units', 'Price'], ['East', 'pear', 4, 1.5], ['East', 'fig', 2, 3], ['North', 'kiwi', 7, 0.5], ['West', 'plum', 1, 2.25], ['West', 'pear', 5, 1.5], ['West', 'lime', 3, 0.75]]
  .flatMap((row, r) => row.map((value, c) => ({ r, c, value }))));
const g = (wb, si, axis, a, b) => groupLines(wb, si, axis, a - 1, b - 1);
const AREA = { r0: 0, c0: 0, r1: 11, c1: 7 };
const STEPS = {
  'automatic styles on': [numbers, (wb, si) => setSummary(wb, si, { styles: true })],
  'automatic styles on then off': [numbers, (wb, si) => { setSummary(wb, si, { styles: true }); setSummary(wb, si, { styles: false }); }],
  'group with automatic styles': [numbers, (wb, si) => { setSummary(wb, si, { styles: true }); g(wb, si, 'row', 2, 4); }],
  'apply styles two levels': [numbers, (wb, si) => { g(wb, si, 'row', 2, 7); g(wb, si, 'row', 3, 4); g(wb, si, 'col', 2, 4); applyOutlineStyles(wb, si, AREA); }],
  'apply styles three levels': [numbers, (wb, si) => { g(wb, si, 'row', 2, 9); g(wb, si, 'row', 3, 7); g(wb, si, 'row', 4, 5); applyOutlineStyles(wb, si, AREA); }],
  'apply styles to one cell': [numbers, (wb, si) => { g(wb, si, 'row', 2, 4); applyOutlineStyles(wb, si, { r0: 4, c0: 0, r1: 4, c1: 0 }); }],
  'apply styles summary above': [numbers, (wb, si) => { setSummary(wb, si, { below: false }); g(wb, si, 'row', 3, 5); applyOutlineStyles(wb, si, AREA); }],
  'auto outline with automatic styles': [sums, (wb, si) => { setSummary(wb, si, { styles: true }); autoOutline(wb, si); }],
  'subtotal with automatic styles': [list, (wb, si) => { setSummary(wb, si, { styles: true }); subtotal(wb, si, { r0: 0, c0: 0, r1: 6, c1: 3 }, { by: 0, fn: 'sum', cols: [2] }); }],
};

const attrsOf = (tag) => Object.fromEntries([...tag.matchAll(/ ([\w:]+)="([^"]*)"/g)].map((m) => [m[1], m[2]]));
const items = (xml, tag) => [...(xml || '').matchAll(new RegExp(`<${tag}\\b[^>]*?(?:/>|>[\\s\\S]*?</${tag}>)`, 'g'))].map((m) => m[0]);
/** Format s described: its style (name builtinId iLevel) and font's bold and italic; Normal and plain as ''. */
function describer({ fonts, cellXfs, cellStyles }) {
  const fontList = items(fonts.replace(/^<fonts[^>]*>/, ''), 'font');
  const xfs = items(cellXfs.replace(/^<cellXfs[^>]*>/, ''), 'xf').map(attrsOf);
  const styles = new Map(items(cellStyles, 'cellStyle').map(attrsOf).map((a) => [a.xfId, `${a.name} ${a.builtinId} ${a.iLevel ?? ''}`.trim()]));
  return (s) => {
    const xf = xfs[Number(s || 0)];
    const font = fontList[Number(xf.fontId || 0)] || '';
    const look = `${/<b\/>/.test(font) ? 'b' : ''}${/<i\/>/.test(font) ? 'i' : ''}`;
    const style = styles.get(xf.xfId || '0') || '';
    return style === 'Normal 0' && !look ? '' : `${style} ${look}`.trim();
  };
}
/** The sheet's formats: { 'A5': ..., 'row 5': ... } for every cell and row with a format other than Normal. */
function formats(sheetData, describe) {
  const out = {};
  for (const m of sheetData.matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
    const a = attrsOf(m[1]);
    if (a.customFormat === '1' && describe(a.s)) out[`row ${a.r}`] = describe(a.s);
    for (const c of (m[2] || '').matchAll(/<c\b([^>]*?)(?:\/>|>)/g)) { const ca = attrsOf(c[1]); if (describe(ca.s)) out[ca.r] = describe(ca.s); }
  }
  return out;
}

for (const [name, [fill, steps]] of Object.entries(STEPS)) {
  test(`automatic styles as Excel applies them: ${name}`, () => {
    const want = byName.get(name);
    assert.ok(want, `no measurement for ${name}`);
    const { wb, si } = book(fill);
    steps(wb, si);
    const sheetXml = serializeXml(wb._part(si).doc);
    const stylesXml = serializeXml(wb.stylesDoc);
    const pick = (re) => (stylesXml.match(re) || [''])[0];
    const ours = describer({ fonts: pick(/<fonts\b[\s\S]*?<\/fonts>/), cellXfs: pick(/<cellXfs\b[\s\S]*?<\/cellXfs>/), cellStyles: pick(/<cellStyles\b[\s\S]*?<\/cellStyles>/) });
    const expected = formats(want.sheetData, describer(want));
    if (!/automatic styles on/.test(name)) assert.ok(Object.keys(expected).length, 'Excel styled cells');
    assert.deepEqual(formats((sheetXml.match(/<sheetData>[\s\S]*?<\/sheetData>/) || [''])[0], ours), expected);
    const pr = (x) => attrsOf((x.match(/<outlinePr\b[^>]*\/>/) || [''])[0]);
    assert.deepEqual(pr(sheetXml), pr(want.sheetPr));
  });
}

test('every automatic styles case of the probe is checked', () => {
  assert.deepEqual(excel.cases.map((c) => c.name).filter((n) => /styles/.test(n) && !STEPS[n]), []);
});
