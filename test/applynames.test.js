import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { applyNamesTo } from '../src/ooxml/applynames.js';
import { parseRange } from '../src/ooxml/refs.js';

const area = (ref) => { const a = parseRange(ref); return { r0: a.r0, c0: a.c0, r1: a.r1, c1: a.c1 }; };
const at = (ref) => { const a = area(ref); return { r: a.r0, c: a.c0 }; };

// Excel's Range.ApplyNames on each formula by itself in H8 (a hidden Excel, docs/NAMES.md), with Qty =$B$2,
// Price =$C$2, Rate =$E$1, Amounts =$B$2:$B$5, Local1 =$F$1 on this sheet (Remote =Data!$A$1 is another sheet's).
const NAMES = [['Qty', 'B2'], ['Price', 'C2'], ['Rate', 'E1'], ['Amounts', 'B2:B5'], ['Local1', 'F1']].map(([name, ref]) => ({ name, area: area(ref) }));
const EXCEL = {
  ignore: [
    ['B2*C2', 'Qty*Price'], ['$B$2*$C$2', 'Qty*Price'], ['B$2*$C2', 'Qty*Price'], ['SUM(B2:B5)', 'SUM(Amounts)'],
    ['SUM($B$2:$B$5)', 'SUM(Amounts)'], ['SUM(B2:B4)', 'SUM(Qty:B4)'], ['SUM($B$2:$B$4)', null], ['SUM(B$2:B4)', 'SUM(Qty:B4)'],
    ['SUM($B2:$B4)', null], ['SUM(B3:B5)', null], ['SUM(A1:B2)', 'SUM(A1:Qty)'], ['B2+1', 'Qty+1'], ['Sheet1!B2*2', null],
    ['Sheet1!$B$2*2', null], ['Data!$A$1+1', null], ['Data!A1', null], ['B2&"B2"', 'Qty&"B2"'], ['E1*2', 'Rate*2'],
    ['B2:B5', 'Amounts'], ['0.07*B2', '0.07*Qty'], ['$E$1+B2', 'Rate+Qty'], ['$F$1', 'Local1'], ['B3', null],
    ['SUM(B2,C2)', 'SUM(Qty,Price)'], ['-B2', '-Qty'], ['B2%', 'Qty%'], ['Qty*2', null], ['SUM(B2:C2)', 'SUM(Qty:Price)'],
    ['B2 B2:B5', 'Qty Amounts'],
  ],
  exact: [
    ['B2*C2', null], ['$B$2*$C$2', 'Qty*Price'], ['B$2*$C2', null], ['SUM(B2:B5)', null], ['SUM($B$2:$B$5)', 'SUM(Amounts)'],
    ['SUM(B2:B4)', null], ['SUM(A1:B2)', null], ['$E$1+B2', 'Rate+B2'], ['$F$1', 'Local1'], ['B2:B5', null],
  ],
};

test('references replaced by their names, as Excel\'s ApplyNames (Ignore Relative/Absolute on, then off)', () => {
  // (H8: no row or column name meets row 8, so Use row and column names makes no difference here, as in Excel.)
  for (const [mode, cases] of Object.entries(EXCEL)) {
    const wrong = cases.filter(([f, want]) => applyNamesTo(f, at('H8'), NAMES, { ignore: mode === 'ignore' }) !== want)
      .map(([f, want]) => `${mode} ${f}: Excel ${want}, here ${applyNamesTo(f, at('H8'), NAMES, { ignore: mode === 'ignore' })}`);
    assert.deepEqual(wrong, []);
  }
  // (In row 2, Use row and column names gave B2 the column's name: =B2*C2 in H2 became =Amounts*Price.)
  assert.equal(applyNamesTo('B2*C2', at('H2'), NAMES), 'Amounts*Price');
  assert.equal(applyNamesTo('B2*C2', at('H2'), NAMES, { rowCol: false }), 'Qty*Price');
});

test('row and column names, as Excel\'s on a grid of them', () => {
  // (Jan C2:C4, Feb D2:D4; North C2:D2, South C3:D3, East C4:D4.)
  const grid = [['Jan', 'C2:C4'], ['Feb', 'D2:D4'], ['North', 'C2:D2'], ['South', 'C3:D3'], ['East', 'C4:D4']].map(([name, ref]) => ({ name, area: area(ref) }));
  const cases = [
    ['H2', 'C2', 'Jan', 'North Jan', 'Jan'], ['H3', 'C2', 'North Jan', 'North Jan', 'Jan North'], ['H2', 'D2', 'Feb', 'North Feb', 'Feb'],
    ['H3', 'D3', 'Feb', 'South Feb', 'Feb'], ['C6', 'C3', 'South', 'South Jan', 'South'], ['D6', 'C3', 'South Jan', 'South Jan', 'Jan South'],
    ['C6', 'D4', 'East Feb', 'East Feb', 'Feb East'], ['H9', 'C3', 'South Jan', 'South Jan', 'Jan South'], ['H9', '$C$3', 'South Jan', 'South Jan', 'Jan South'],
    ['H9', 'D4*2', 'East Feb*2', 'East Feb*2', 'Feb East*2'], ['H3', 'C3+D3', 'Jan+Feb', 'South Jan+South Feb', 'Jan+Feb'],
    ['C7', 'C2+C3', 'North+South', 'North Jan+South Jan', 'North+South'], ['H9', 'SUM(C2:C4)', 'SUM(Jan)', 'SUM(Jan)', 'SUM(Jan)'],
    ['H9', 'SUM(C2:D2)', 'SUM(North)', 'SUM(North)', 'SUM(North)'], ['H9', 'C5', null, null, null], ['H2', 'E2', null, null, null],
  ];
  const wrong = [];
  for (const [cell, f, byDefault, omitNothing, colFirst] of cases) {
    const got = [applyNamesTo(f, at(cell), grid), applyNamesTo(f, at(cell), grid, { omitCol: false, omitRow: false }), applyNamesTo(f, at(cell), grid, { colFirst: true })];
    if (got.join('|') !== [byDefault, omitNothing, colFirst].join('|')) wrong.push(`${cell} =${f}: Excel ${[byDefault, omitNothing, colFirst].join(', ')}, here ${got.join(', ')}`);
  }
  assert.deepEqual(wrong, []);
});

test('over a sheet: the formulas the selection holds rewritten, their values the same; one cell is the whole sheet', async () => {
  const { build } = await import('./steps.js');
  const { defineName, deleteName } = await import('../src/ooxml/names.js');
  const { applicableNames, applyNamesSteps } = await import('../src/ooxml/applynames.js');
  const { wb, si } = build({ cells: { B2: '4', C2: '5', E1: '0.07', B3: '1', B4: '1', B5: '1', H2: '=B2*C2', H3: '=SUM(B2:B5)*E1', H5: '="B2"&B2' } });
  for (const d of [...wb.definedNames]) deleteName(wb, { name: d.name, sheet: d.sheet ?? null });
  const here = wb.sheets[si].name;
  wb.apply(si, [{ r: 3, c: 7, formula: `'${here}'!B2` }]);
  for (const [n, ref] of [['Qty', '$B$2'], ['Price', '$C$2'], ['Rate', '$E$1'], ['Amounts', '$B$2:$B$5'], ['Rel', 'B3'], ['Tax', null]]) {
    defineName(wb, { name: n, refersTo: ref ? `='${here}'!${ref}` : '=0.07' });
  }
  const names = applicableNames(wb, si);
  assert.deepEqual(names.map((n) => n.name).sort(), ['Amounts', 'Price', 'Qty', 'Rate']);
  const f = (r) => wb.sheet(si).cells.get(r).get(7);
  const before = [1, 2, 3, 4].map((r) => f(r).v);
  const sheetRef = f(3).f;
  const got = applyNamesSteps(wb, si, { r0: 0, c0: 0, r1: 0, c1: 0 }, names);
  wb.apply(si, got.steps[0].contents);
  // (H2 in row 2: Amounts is B's column there, as Excel's Use row and column names gave it; H4, with its sheet's
  // name, stays.)
  assert.deepEqual([1, 2, 3, 4].map((r) => f(r).f), ['Amounts*Price', 'SUM(Amounts)*Rate', sheetRef, '"B2"&Qty']);
  // (Values as before but H2's: Amounts*Price spills into H3, #SPILL!, as Excel's did after its ApplyNames.)
  assert.deepEqual([1, 2, 3, 4].map((r) => f(r).v), ['#SPILL!', ...before.slice(1)]);
  assert.equal(got.count, 3);
  assert.equal(applyNamesSteps(wb, si, { r0: 0, c0: 0, r1: 0, c1: 0 }, names).none, true);
});
