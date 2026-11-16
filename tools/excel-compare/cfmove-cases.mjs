// Inserting, deleting and sorting on a sheet with conditional formats of both forms, for excel-cfmoves.ps1: data bars
// on B2:B10, 3 stars on C2:C10, =D2>Data!$A$1 on D2:D10, greater than 5 on E2:E10 (made-up numbers); one step each.
// usage: node tools/excel-compare/cfmove-cases.mjs > tools/excel-compare/cfmove-cases.json
const step = (name, s) => ({ name, step: s });
const cases = [
  step('start', { do: 'none', range: 'A1' }),
  step('insert row 1', { do: 'insert', range: '1:1' }),
  step('insert rows 5:6 (inside the formats)', { do: 'insert', range: '5:6' }),
  step('insert row 11 (just below the formats)', { do: 'insert', range: '11:11' }),
  step('delete rows 2:3 (their first rows)', { do: 'delete', range: '2:3' }),
  step('delete rows 2:10 (all of them)', { do: 'delete', range: '2:10' }),
  step('insert column A', { do: 'insert', range: 'A:A' }),
  step('insert column C (between the bars and the stars)', { do: 'insert', range: 'C:C' }),
  step('insert column F (just after the last format)', { do: 'insert', range: 'F:F' }),
  step('delete column B (the bars)', { do: 'delete', range: 'B:B' }),
  step('delete column C (the stars)', { do: 'delete', range: 'C:C' }),
  step('delete columns B:E (all of them)', { do: 'delete', range: 'B:E' }),
  step('insert C5 shifting down (inside the stars)', { do: 'insert', range: 'C5', shift: 'down' }),
  step('insert B11:E11 shifting down (under all of them)', { do: 'insert', range: 'B11:E11', shift: 'down' }),
  step('delete C5:C6 shifting up', { do: 'delete', range: 'C5:C6', shift: 'up' }),
  step('delete C5 shifting left (a star moved onto the formula rule)', { do: 'delete', range: 'C5', shift: 'left' }),
  step('delete D5 shifting left (a formula cell moved onto the stars)', { do: 'delete', range: 'D5', shift: 'left' }),
  step('insert D5:D6 shifting right (inside the formula rule)', { do: 'insert', range: 'D5:D6', shift: 'right' }),
  step('Data: insert row 1', { do: 'insert', range: '1:1', sheet: 'Data' }),
  step('Data: delete row 1', { do: 'delete', range: '1:1', sheet: 'Data' }),
  step('sort B2:E10 by B, largest first', { do: 'sort', range: 'B2:E10', key: 'B2', desc: true }),
];
console.log(JSON.stringify({ cases }, null, 1));
