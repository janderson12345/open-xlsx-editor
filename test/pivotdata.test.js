import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { Workbook } from '../src/ooxml/workbook.js';
import { createPivot } from '../src/ooxml/pivotnew.js';

// (Excel 365 over COM, on PivotTables it made from this data at H1 (Region and Item on rows, Year on columns, Sum of
// Sales) and H20 (Item on rows; Sum of Sales, Count of Qty, Average of Sales): GETPIVOTDATA's answers. The same
// PivotTables made here give the same.)
const DATA = [['Region', 'Item', 'Year', 'Sales', 'Qty'], ['East', 'Apple', 2023, 10, 1], ['East', 'Pear', 2023, 20, 2], ['West', 'Apple', 2024, 30, 3],
  ['West', 'Plum', 2023, 40, 4], ['East', 'Apple', 2024, 5, 5], ['North', 'Pear', 2024, 7, null]];
const EXCEL = [
  ["GETPIVOTDATA(\"Sales\",H1)", 112],
  ["GETPIVOTDATA(\"Sum of Sales\",H1)", 112],
  ["GETPIVOTDATA(\"sum of sales\",H3)", 112],
  ["GETPIVOTDATA(\" Sales\",H1)", "#REF!"],
  ["GETPIVOTDATA(\"Sales\",H1,\"Region\",\"East\")", 35],
  ["GETPIVOTDATA(\"Sales\",H1,\"region\",\"EAST\")", 35],
  ["GETPIVOTDATA(\"Sales\",H1,\"Region\",\"East\",\"Item\",\"Apple\")", 15],
  ["GETPIVOTDATA(\"Sales\",H1,\"Item\",\"Apple\")", "#REF!"],
  ["GETPIVOTDATA(\"Sales\",H1,\"Year\",2023)", 70],
  ["GETPIVOTDATA(\"Sales\",H1,\"Year\",\"2023\")", 70],
  ["GETPIVOTDATA(\"Sales\",H1,\"Year\",2023,\"Region\",\"West\")", 40],
  ["GETPIVOTDATA(\"Sales\",H1,\"Region\",\"South\")", "#REF!"],
  ["GETPIVOTDATA(\"Sales\",H1,\"Qty\",1)", "#REF!"],
  ["GETPIVOTDATA(\"Qty\",H1)", "#REF!"],
  ["GETPIVOTDATA(\"Sales\",A1)", "#REF!"],
  ["GETPIVOTDATA(\"Sales\",H1,\"Region\")", "#REF!"],
  ["GETPIVOTDATA(\"Sales\",H1,\"Region\",\"East\",\"Item\",\"Plum\")", "#REF!"],
  ["GETPIVOTDATA(\"Sales\",H1,\"Year\",2025)", "#REF!"],
  ["GETPIVOTDATA(\"Sales\",H1:K3)", 112],
  ["GETPIVOTDATA(\"Sales\",\"H1\")", "#REF!"],
  ["GETPIVOTDATA(\"Qty\",H20)", 5],
  ["GETPIVOTDATA(\"Count of Qty\",H20)", 5],
  ["GETPIVOTDATA(\"Count of Qty\",H20,\"Item\",\"Pear\")", 1],
  ["GETPIVOTDATA(\"Average of Sales\",H20,\"Item\",\"Apple\")", 15],
  ["GETPIVOTDATA(\"Sum of Sales\",H20,\"Item\",\"Plum\")", 40],
  ["GETPIVOTDATA(\"Sales\",H20,\"Item\",\"Apple\")", 45],
  ["GETPIVOTDATA(\"Year\",H20)", "#REF!"],
  ["GETPIVOTDATA(H1,H1)", 112],
  ["GETPIVOTDATA(\"Sales\",H1,\"Item\",\"Apple\",\"Item\",\"Pear\")", "#REF!"],
  ["GETPIVOTDATA(\"Sales\",H1,\"Region\",\"East\",\"Item\",\"Pear\",\"Year\",2024)", 0],
  ["GETPIVOTDATA(\"Sales\",H1,\"Region\",\"North\",\"Year\",2023)", 0],
  ["GETPIVOTDATA(\"Sales\",H1,\"Region\",\"East\",\"Year\",2024)", 5],
  ["GETPIVOTDATA(\"Sales\",H1,\"Item\",\"Apple\",\"Region\",\"East\")", 15],
  ["GETPIVOTDATA(\"Sales\",H1,\"Region\",\"east \")", "#REF!"],
  ["GETPIVOTDATA(\"Sales\",H1,\"Year\",2023.0)", 70],
  ["GETPIVOTDATA(\"Sales\",H1,\"Year\",TRUE)", "#REF!"],
  ["GETPIVOTDATA(\"Average of Sales\",H20)", 18.6666666666667],
  ["GETPIVOTDATA(\"Sales\",H20)", 112],
  ["GETPIVOTDATA(\"Average\",H20)", "#REF!"],
  ["GETPIVOTDATA(\"Sales\",H1,\"Region\",\"East\",\"Year\",2024,\"Item\",\"Apple\")", 5],
  ["GETPIVOTDATA(\"SALES\",H1)", 112],
  ["GETPIVOTDATA(\"Sales\",H1,\"Region\",1)", "#REF!"],
  ["GETPIVOTDATA(\"Sales\",H1,\"Region\",\"East\",\"Region\",\"East\")", 35],
  ["GETPIVOTDATA(\"Sum of Sales\",H20,\"Item\",\"Pear\")", 27],
  ["GETPIVOTDATA(\"Sales\",H1,\"Item\",\"Pear\",\"Region\",\"North\")", 7],
];

test('GETPIVOTDATA: 45 asks of two PivotTables, as Excel answers them', () => {
  const wb = Workbook.fromBytes(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')));
  const si = wb.addSheet(wb.sheets.length, 'Src');
  wb.apply(si, DATA.flatMap((row, r) => row.map((value, c) => ({ r, c, value }))).filter((x) => x.value != null));
  const src = { si, r0: 0, c0: 0, r1: DATA.length - 1, c1: 4 };
  for (const [at, layout] of [[{ si, r: 0, c: 7 }, { rows: [0, 1], cols: [2], values: [{ field: 3, fn: 'sum' }] }],
    [{ si, r: 19, c: 7 }, { rows: [1], cols: [], values: [{ field: 3, fn: 'sum' }, { field: 4, fn: 'count' }, { field: 3, fn: 'average' }] }]]) {
    const out = createPivot(wb, src, at, layout);
    assert.ok(!out.error, out.error);
    for (const ch of out.changes) wb.apply(ch.si, ch.contents);
  }
  for (const [f, want] of EXCEL) {
    const v = wb.calc.formulaAt(si, 59, 0, f);
    const got = v && v.code ? v.code : v;
    if (typeof want === 'number') assert.ok(Math.abs(got - want) < 1e-12, `${f}: ${got}`);
    else assert.equal(got, want, f);
  }
});
