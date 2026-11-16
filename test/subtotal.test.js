import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { serializeXml } from '../src/ooxml/xml.js';
import { subtotal, removeSubtotals } from '../src/ooxml/subtotal.js';
import { groupLines } from '../src/ooxml/outline.js';
import { currentRegion } from '../src/ooxml/special.js';

// Excel's Subtotal and Remove All: test/outline-excel.json holds what Excel wrote (tools/excel-compare/excel-outline.ps1,
// Range.Subtotal and RemoveSubtotal on a made-up list Region, Item, Units, Price, saved and read back). The same here,
// and every cell of A1:D25 compared with Excel's (its formula, its value, numbers to 1e-9, whether it is bold), with
// each row's outline level, hidden and collapsed, sheetFormatPr's outlineLevelRow, outlinePr and the row breaks.
const excel = JSON.parse(fs.readFileSync('test/outline-excel.json', 'utf8'));
const byName = new Map(excel.cases.map((c) => [c.name, c]));
const LIST = [['Region', 'Item', 'Units', 'Price'], ['East', 'pear', 4, 1.5], ['East', 'fig', 2, 3], ['North', 'kiwi', 7, 0.5], ['West', 'plum', 1, 2.25], ['West', 'pear', 5, 1.5], ['West', 'lime', 3, 0.75]];

function list() {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'List');
  wb.apply(si, LIST.flatMap((row, r) => row.map((value, c) => ({ r, c, value }))));
  return { wb, si };
}
const A = (r1) => ({ r0: 0, c0: 0, r1: r1 - 1, c1: 3 });
const ST = (wb, si, r1, by, fn, cols, replace, pageBreaks, below) => subtotal(wb, si, A(r1), { by, fn, cols, replace, pageBreaks, below });
const STEPS = {
  'subtotal sum units': (wb, si) => ST(wb, si, 7, 0, 'sum', [2], true, false, true),
  'subtotal count two columns': (wb, si) => ST(wb, si, 7, 0, 'count', [2, 3], true, false, true),
  'subtotal average above': (wb, si) => ST(wb, si, 7, 0, 'average', [2], true, false, false),
  'subtotal page breaks': (wb, si) => ST(wb, si, 7, 0, 'sum', [2], true, true, true),
  'subtotal max then add min': (wb, si) => { ST(wb, si, 7, 0, 'max', [2], true, false, true); ST(wb, si, 11, 0, 'min', [2], false, false, true); },
  'subtotal replace': (wb, si) => { ST(wb, si, 7, 0, 'max', [2], true, false, true); ST(wb, si, 11, 0, 'sum', [3], true, false, true); },
  'subtotal at item': (wb, si) => ST(wb, si, 7, 1, 'sum', [2], true, false, true),
  'subtotal from one cell': (wb, si) => subtotal(wb, si, currentRegion(wb.sheet(si), 2, 1), { by: 0, fn: 'sum', cols: [2] }),
  'subtotal functions': (wb, si) => { ST(wb, si, 7, 0, 'product', [2], true, false, true); ST(wb, si, 11, 0, 'countNums', [3], false, false, true); },
  'subtotal stdev var': (wb, si) => { ST(wb, si, 7, 0, 'stdDev', [2], true, false, true); ST(wb, si, 11, 0, 'stdDevp', [2], false, false, true); ST(wb, si, 15, 0, 'var', [2], false, false, true); ST(wb, si, 19, 0, 'varp', [2], false, false, true); },
  'remove subtotals': (wb, si) => { ST(wb, si, 7, 0, 'sum', [2], true, true, true); removeSubtotals(wb, si, A(11)); },
};

const attrsOf = (tag) => Object.fromEntries([...tag.matchAll(/ ([\w:]+)="([^"]*)"/g)].map((m) => [m[1], m[2]]));
/** Excel's saved sheet as rows: { r: { attrs, cells: { ref: { f, v, bold } } } }, its strings resolved. */
function excelRows(c) {
  const boldXf = new Set([...c.cellXfs.matchAll(/<xf\b[^>]*>/g)].map((m, k) => [k, attrsOf(m[0])]).filter(([, a]) => a.fontId === '1').map(([k]) => String(k)));
  const out = {};
  for (const m of c.sheetData.matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)) {
    const a = attrsOf(m[1]);
    const cells = {};
    for (const x of m[2].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
      const ca = attrsOf(x[1]);
      const f = (x[2].match(/<f>([^<]*)<\/f>/) || [])[1] ?? null;
      const raw = (x[2].match(/<v>([^<]*)<\/v>/) || [])[1];
      const v = ca.t === 's' ? c.strings[Number(raw)] : ca.t === 'e' || ca.t === 'str' ? raw : Number(raw);
      cells[ca.r] = { f, v, bold: boldXf.has(ca.s || '0') };
    }
    out[a.r] = { level: a.outlineLevel || null, hidden: a.hidden || null, collapsed: a.collapsed || null, cells };
  }
  return out;
}
function ourRows(wb, si) {
  const sheet = wb.sheet(si);
  const out = {};
  for (const r of new Set([...sheet.cells.keys(), ...sheet.rows.keys()])) {
    const info = sheet.rows.get(r) || {};
    const cells = {};
    for (const [c, cell] of sheet.cells.get(r) || []) {
      const v = cell.t === 'e' ? cell.v : cell.v;
      cells[`${'ABCD'[c]}${r + 1}`] = { f: cell.f ?? null, v, bold: !!wb.styles.fonts[wb.styles.cellXfs[cell.s || 0].fontId]?.b };
    }
    out[String(r + 1)] = { level: info.level ? String(info.level) : null, hidden: info.hidden ? '1' : null, collapsed: info.collapsed ? '1' : null, cells };
  }
  return out;
}
const near = (o) => JSON.parse(JSON.stringify(o, (k, v) => (typeof v === 'number' ? Number(v.toPrecision(12)) : v)));

for (const [name, steps] of Object.entries(STEPS)) {
  test(`Subtotal as Excel does it: ${name}`, () => {
    const want = byName.get(name);
    const { wb, si } = list();
    steps(wb, si);
    wb.calculate();
    assert.deepEqual(near(ourRows(wb, si)), near(excelRows(want)));
    const xml = serializeXml(wb._part(si).doc);
    const level = (s) => attrsOf((s.match(/<sheetFormatPr\b[^>]*>/) || [''])[0]).outlineLevelRow;
    assert.equal(level(xml), level(want.sheetFormatPr), 'outlineLevelRow');
    assert.deepEqual(attrsOf((xml.match(/<outlinePr\b[^>]*\/>/) || [''])[0]), attrsOf((want.sheetPr.match(/<outlinePr\b[^>]*\/>/) || [''])[0]));
    const breaks = (s) => [...s.matchAll(/<brk\b([^>]*)\/>/g)].map((m) => attrsOf(m[1]));
    assert.deepEqual(breaks((xml.match(/<rowBreaks\b[\s\S]*?<\/rowBreaks>/) || [''])[0]), breaks(want.rowBreaks));
  });
}

test('Subtotal over a list with no rows under its headings does nothing', () => {
  const { wb, si } = list();
  assert.equal(subtotal(wb, si, { r0: 0, c0: 0, r1: 0, c1: 3 }, { by: 0, cols: [2] }), null);
});

test('Subtotal on rows grouped already goes a level deeper', () => {
  const { wb, si } = list();
  groupLines(wb, si, 'row', 1, 6);
  subtotal(wb, si, A(7), { by: 0, fn: 'sum', cols: [2], replace: false });
  assert.equal(wb.sheet(si).rows.get(1).level, 2);
  assert.equal(wb.sheet(si).rows.get(3).level, 1);
});
