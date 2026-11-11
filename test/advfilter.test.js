import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { parseRange } from '../src/ooxml/refs.js';
import { shownText } from '../src/ooxml/filterops.js';

// Data > Advanced as Excel does it: test/advfilter-excel.json holds 7 cases of Range.AdvancedFilter on the same made-up
// list (tools/excel-compare/excel-advfilter.ps1): the rows hidden or the cells copied, and the names Excel wrote. The
// same list and criteria here, and the same Advanced Filter.
const excel = JSON.parse(fs.readFileSync('test/advfilter-excel.json', 'utf8')).cases;
const LIST = [['Name', 'Qty', 'Team'], ['apple', 1, 'red'], ['pear', 2, 'blue'], ['Apple', 3, 'red'], ['fig', 1, 'green'], ['plum', 2.5, 'red'], ['kiwi', 10, 'blue'], ['pear', 2, 'blue'], ['avocado', 7, 'red']];
const value = (t) => (t !== '' && Number.isFinite(Number(t)) ? Number(t) : t);
function build(c) {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  wb.renameSheet(0, 'Sheet1');
  const sheet = wb.sheet(0);
  for (const r of [...sheet.cells.keys()]) for (const k of [...sheet.cells.get(r).keys()]) wb.setCell(0, { r, c: k }, null);
  wb.setStyles(0, [...sheet.cells].flatMap(([r, line]) => [...line.keys()].map((k) => ({ r, c: k, s: 0 }))));
  LIST.forEach((row, r) => row.forEach((v, k) => wb.setCell(0, { r, c: k }, v)));
  // (A criteria cell ="=pear" is the text =pear.)
  const rows = (c.criteria || []).map((line) => line.split('|'));
  rows.forEach((row, r) => row.forEach((t, k) => { if (t !== '') wb.setCell(0, { r, c: 5 + k }, /^="(.*)"$/.test(t) ? t.slice(2, -1) : value(t)); }));
  for (const [r, info] of wb.sheet(0).rows) if (info.hidden) wb.setLines(0, 'row', r, r, { hidden: false });
  return { wb, criteria: rows.length ? { r0: 0, c0: 5, r1: rows.length - 1, c1: 5 + rows[0].length - 1 } : null };
}

test("Advanced Filter: the rows hidden or copied, and Excel's names (7 cases)", () => {
  const wrong = [];
  for (const c of excel) {
    const { wb, criteria } = build(c);
    const refused = wb.advancedFilter(0, { list: parseRange('A1:C9'), criteria, copyTo: c.action === 2 ? { r: 0, c: 8 } : null, unique: c.unique });
    if (refused) { wrong.push(`${c.name}: ${refused}`); continue; }
    const hidden = [...wb.sheet(0).rows].filter(([, row]) => row.hidden).map(([r]) => r + 1).sort((a, b) => a - b);
    if (JSON.stringify(hidden) !== JSON.stringify(c.hidden)) wrong.push(`${c.name}: Excel hides ${c.hidden}, here ${hidden}`);
    const copied = [];
    for (let r = 0; r < 9; r++) { const line = [8, 9, 10].map((k) => shownText(wb, wb.sheet(0).cells.get(r)?.get(k))).join('|'); if (line !== '||') copied.push(line); }
    if (JSON.stringify(copied) !== JSON.stringify(c.copied)) wrong.push(`${c.name}: Excel copies ${c.copied}, here ${copied}`);
    const names = wb.definedNames.filter((d) => d.name.startsWith('_xlnm.') && d.sheet === 0).map((d) => `${d.name}${d.el.getAttribute('hidden') ? ' hidden' : ''}=${d.text}`).sort();
    const want = c.names.map((n) => { const m = /name="([^"]+)"[^>]*?(hidden="1")?>([^<]*)</.exec(n); return `${m[1]}${m[2] ? ' hidden' : ''}=${m[3]}`; }).sort();
    if (JSON.stringify(names) !== JSON.stringify(want)) wrong.push(`${c.name}: Excel's names ${want}, here ${names}`);
  }
  assert.equal(excel.length, 7);
  assert.deepEqual(wrong, []);
});

test('Advanced Filter refuses criteria headed by a name the list hasn\'t', () => {
  const { wb } = build({ criteria: ['Colour', 'red'] });
  assert.match(wb.advancedFilter(0, { list: parseRange('A1:C9'), criteria: { r0: 0, c0: 5, r1: 1, c1: 5 } }), /headings/);
});
