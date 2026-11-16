// Inserting, deleting and sorting on a sheet with sparklines, data validation rules (one reading another sheet), Circle
// Invalid Data's ovals, notes and a threaded comment, for excel-refmoves.ps1: the sheet built the same each time
// (made-up numbers and letters), one step done, and where each thing went read back.
// usage: node tools/excel-compare/refmove-cases.mjs > tools/excel-compare/refmove-cases.json
const values = {};
for (let r = 2; r <= 7; r++) {
  if (r === 6) continue;
  for (let c = 2; c <= 13; c++) values[`${String.fromCharCode(64 + c)}${r}`] = ((r * 7 + c * 3) % 11) - 3;
}
// (Invalid on purpose: B12 not in the list, C12 and C14 over I10, D11 and D14 not over C, E12 not in Data's list.)
const column = (col, list) => list.forEach((v, i) => { values[`${col}${10 + i}`] = v; });
column('B', ['a', 'b', 'x', 'c', 'a', 'b']);
column('C', [3, 1, 6, 2, 5, 4]);
column('D', [4, 0, 7, 3, 1, 5]);
column('E', ['p', 'q', 'zz', 'r', 'p', 'q']);
column('H', ['a', 'b', 'c']);
values.I10 = 4;
const sheet = {
  values,
  sparklines: [
    { at: 'O2:O5', data: 'Sheet1!B2:M5', type: 1 },
    { at: 'O7', data: 'Sheet1!B7:M7', type: 2 },
    { at: 'Q2', data: 'Data!B2:M2', type: 1 },
  ],
  // (Validation.Add: type 3 list, 1 whole number, 7 custom; operator 1 between.)
  validations: [
    { at: 'B10:B15', type: 3, op: 1, formula1: '=$H$10:$H$12' },
    { at: 'C10:C15', type: 1, op: 1, formula1: '1', formula2: '=$I$10' },
    { at: 'D10:D15', type: 7, op: 1, formula1: '=D10>C10' },
    { at: 'E10:E15', type: 3, op: 1, formula1: '=Data!$A$1:$A$3' },
  ],
  notes: ['B12', 'F20'],
  threads: ['C13'],
};
const data = { values: { A1: 'p', A2: 'q', A3: 'r', ...Object.fromEntries(Array.from({ length: 12 }, (_, i) => [`${String.fromCharCode(66 + i)}2`, (i * 5) % 7])) } };
const step = (name, s) => ({ name, step: s });
const cases = [
  step('start', { do: 'none', range: 'A1' }),
  step('insert row 1', { do: 'insert', range: '1:1' }),
  step('insert rows 3:4 (inside the sparkline group)', { do: 'insert', range: '3:4' }),
  step('delete row 3 (a sparkline and its data)', { do: 'delete', range: '3:3' }),
  step('delete rows 2:5 (all of the group)', { do: 'delete', range: '2:5' }),
  step('insert column C (inside the data)', { do: 'insert', range: 'C:C' }),
  step('delete column C (part of the data, a rule)', { do: 'delete', range: 'C:C' }),
  step('delete columns B:M (all of the data)', { do: 'delete', range: 'B:M' }),
  step('delete column O (the group\'s cells)', { do: 'delete', range: 'O:O' }),
  step('insert column A', { do: 'insert', range: 'A:A' }),
  step('insert rows 11:12 (inside the rules)', { do: 'insert', range: '11:12' }),
  step('delete row 12 (invalid cells, a note)', { do: 'delete', range: '12:12' }),
  step('delete rows 10:12 (the rules\' first rows)', { do: 'delete', range: '10:12' }),
  step('delete rows 10:15 (all of the rules)', { do: 'delete', range: '10:15' }),
  step('delete column H (a list\'s source)', { do: 'delete', range: 'H:H' }),
  step('insert B11:E11 shifting down', { do: 'insert', range: 'B11:E11', shift: 'down' }),
  step('insert C11 shifting down (part of a rule)', { do: 'insert', range: 'C11', shift: 'down' }),
  step('delete C12 shifting left', { do: 'delete', range: 'C12', shift: 'left' }),
  step('delete C12:C13 shifting up', { do: 'delete', range: 'C12:C13', shift: 'up' }),
  step('insert O3 shifting right (a sparkline\'s cell)', { do: 'insert', range: 'O3', shift: 'right' }),
  step('insert B3:M3 shifting down (the data, not the sparklines)', { do: 'insert', range: 'B3:M3', shift: 'down' }),
  step('Data: insert row 1', { do: 'insert', range: '1:1', sheet: 'Data' }),
  step('Data: insert column A', { do: 'insert', range: 'A:A', sheet: 'Data' }),
  step('Data: delete rows 1:3', { do: 'delete', range: '1:3', sheet: 'Data' }),
  step('sort B10:E15 by C, largest first', { do: 'sort', range: 'B10:E15', key: 'C10', desc: true }),
  step('sort B10:I15 by B', { do: 'sort', range: 'B10:I15', key: 'B10' }),
  step('sort B2:O5 by B (the sparklines\' cells in it)', { do: 'sort', range: 'B2:O5', key: 'B2' }),
  step('sort B2:M5 by B (their data only)', { do: 'sort', range: 'B2:M5', key: 'B2' }),
  step('sort B10:F15 left to right by row 10', { do: 'sort', range: 'B10:F15', key: 'B10', leftToRight: true }),
  // (Whether new lines and cells take the sparklines of the line or cell before them, as they take its format.)
  step('insert column P (after the group\'s column)', { do: 'insert', range: 'P:P' }),
  step('insert row 6 (under the group)', { do: 'insert', range: '6:6' }),
  step('insert row 8 (under the column sparkline)', { do: 'insert', range: '8:8' }),
  step('insert O3 shifting down (a sparkline above)', { do: 'insert', range: 'O3', shift: 'down' }),
  step('insert P2:P3 shifting right (sparklines to the left)', { do: 'insert', range: 'P2:P3', shift: 'right' }),
  step('delete O3 shifting up (a sparkline)', { do: 'delete', range: 'O3', shift: 'up' }),
  // (Whether new lines and cells take the rules of the line or cell before them too.)
  step('insert row 16 (under the rules)', { do: 'insert', range: '16:16' }),
  step('insert column F (after the last rule)', { do: 'insert', range: 'F:F' }),
  step('insert B16:E16 shifting down (under the rules)', { do: 'insert', range: 'B16:E16', shift: 'down' }),
  step('insert B16:C16 shifting down (under two of the rules)', { do: 'insert', range: 'B16:C16', shift: 'down' }),
  step('insert F10:F15 shifting right (after the rules)', { do: 'insert', range: 'F10:F15', shift: 'right' }),
  step('insert E12:E13 shifting right (inside a rule)', { do: 'insert', range: 'E12:E13', shift: 'right' }),
  step('delete C12:E12 shifting left (rules deleted and moved)', { do: 'delete', range: 'C12:E12', shift: 'left' }),
  step('delete B12:C12 shifting left (one rule deleted, one moved)', { do: 'delete', range: 'B12:C12', shift: 'left' }),
];
console.log(JSON.stringify({ sheet, data, cases }, null, 1));
