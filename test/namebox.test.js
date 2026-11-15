import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { selectionName, goToTarget, goToNames } from '../src/ooxml/goto.js';
import { defineName } from '../src/ooxml/names.js';
import { setPrintArea } from '../src/ooxml/pagesetupedit.js';
import { parseRange, rangeName } from '../src/ooxml/refs.js';

// The name a selection has, against what Excel reported for the same names and selection (Selection.Name in a hidden
// Excel, tools/excel-compare/excel-namebox.ps1; COM can't read the Name Box itself).
const measured = JSON.parse(fs.readFileSync('test/namebox-excel.json', 'utf8'));
const open = (file) => Workbook.fromBytes(new Uint8Array(fs.readFileSync(`test/fixtures/${file}`)));

/** A workbook of two sheets, read as Excel's Sheet1 and Sheet2, with a case's names and print area. */
function bookOf(c) {
  const wb = open('general.xlsx');
  wb.addSheet(1);
  const named = (text) => text.replace(/Sheet1!/g, `${wb.sheets[0].name}!`).replace(/Sheet2!/g, `${wb.sheets[1].name}!`);
  for (const n of c.names || []) {
    // (A name of ranges, as a file from Excel has it: the New Name box here doesn't take commas at the top level.)
    const [first, ...more] = named(n.refersTo).split(',');
    defineName(wb, { name: n.name, refersTo: first, sheet: n.sheet, hidden: n.hidden });
    if (more.length) {
      const d = wb.definedNames.find((x) => x.name === n.name && x.sheet === (n.sheet ?? null));
      d.text = [first, ...more].join(',').replace(/^=/, '');
      d.el.textContent = d.text;
    }
  }
  if (c.print) setPrintArea(wb, 0, [parseRange(c.print.replace(/\$/g, ''))]);
  return wb;
}

/** Excel's answer with this workbook's sheet names ("Sheet2!Theirs"), the shown sheet's own without its sheet. */
const expected = (wb, c) => {
  if (c.excel == null) return null;
  const [sheet, name] = c.excel.includes('!') ? c.excel.split('!') : [null, c.excel];
  if (sheet == null || sheet === `Sheet${c.on + 1}`) return name;
  return `${wb.sheets[Number(sheet.slice(5)) - 1].name}!${name}`;
};

test('the selection\'s name, as Excel reports it for the same names (Selection.Name, measured)', () => {
  let same = 0;
  for (const c of measured.cases.filter((x) => !x.table)) {
    const wb = bookOf(c);
    const areas = c.select.split(',').map((ref) => goToTarget(wb, c.on, ref).sel);
    assert.equal(selectionName(wb, c.on, areas), expected(wb, c), c.id);
    same++;
  }
  assert.equal(same, 31, 'every case without a table');
});

test('a table is no name: its rows have none, a name over them has its own', () => {
  const wb = open('formulas.xlsx');
  const data = wb.sheetIndexOf('Data');
  const of = (ref) => selectionName(wb, data, [parseRange(ref)]);
  assert.equal(of('A2:E9'), null, "Fruit's data rows");
  assert.equal(of('A1:E9'), null, 'the whole of Fruit');
  assert.equal(of('B2:B9'), null, 'a column of it');
  defineName(wb, { name: 'Body', refersTo: '=Data!$A$2:$E$9' });
  assert.equal(of('A2:E9'), 'Body');
  for (const c of measured.cases.filter((x) => x.table)) assert.equal(c.excel, c.names.length ? 'Body' : null, c.id);
});

test('going to a name selects its ranges as Application.Goto did, the name the selection\'s then (measured)', () => {
  const wb = bookOf({ names: [{ name: 'Rate', refersTo: '=Sheet1!$B$2:$C$4' }, { name: 'Multi', refersTo: '=Sheet1!$C$3:$D$4,Sheet1!$A$1' }] });
  for (const g of measured.goto) {
    const t = goToTarget(wb, 1, g.name);
    assert.equal(t.si, 0, g.name);
    const areas = [...(t.sel.more || []), t.sel];
    assert.equal(areas.map((a) => rangeName(a)).join(','), g.selection.replace(/\$/g, ''), g.name);
    assert.equal(rangeName({ r0: t.sel.ar, c0: t.sel.ac, r1: t.sel.ar, c1: t.sel.ac }), g.active.replace(/\$/g, ''), g.name);
    assert.equal(selectionName(wb, t.si, areas), g.excel, g.name);
  }
  assert.ok(goToNames(wb, 0).includes('Multi'), 'a name of ranges of one sheet in the list to go to');
});
