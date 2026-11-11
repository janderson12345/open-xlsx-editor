// Inserting and deleting rows, columns and cells on a sheet with a table, a chart, a shape and notes, for
// excel-anchors.ps1: the sheet is built the same each time (made-up stock list), one step done, and where each thing
// ended up read back (or Excel's refusal).
// usage: node tools/excel-compare/anchor-cases.mjs [--no-table] > tools/excel-compare/anchor-cases.json
//   --no-table: the same sheet without the table (the data as plain cells), for notes and drawings alone
const sheet = {
  data: [['Item', 'Qty', 'Price'], ...['Bolt', 'Nut', 'Screw', 'Washer', 'Rivet', 'Pin', 'Clip', 'Hook', 'Ring', 'Peg'].map((name, i) => [name, (i * 7) % 13 + 2, ((i * 5) % 9) + 0.5])],
  at: 'B2', // the data, and the table "Stock" on it (B2:D12)
  chart: { range: 'F3:K14', series: 'C3:C12' }, // a clustered column chart of Qty, sized to these cells
  shape: 'M5:O8', // a rectangle sized to these cells
  notes: ['B6', 'F20', 'A1', 'D12'],
  table: !process.argv.includes('--no-table'),
};
const step = (name, s) => ({ name, step: s });
const cases = [
  step('insert row 1', { do: 'insert', range: '1:1' }),
  step('insert rows 5:6 (inside the table and chart)', { do: 'insert', range: '5:6' }),
  step('insert rows 15:17 (below the table and chart)', { do: 'insert', range: '15:17' }),
  step('delete row 1', { do: 'delete', range: '1:1' }),
  step('delete rows 5:6 (table rows, a note)', { do: 'delete', range: '5:6' }),
  step('delete row 12 (the table\'s last row)', { do: 'delete', range: '12:12' }),
  step('delete rows 3:14 (the table\'s data, all of the chart)', { do: 'delete', range: '3:14' }),
  step('delete rows 2:12 (all of the table)', { do: 'delete', range: '2:12' }),
  step('delete rows 6:7 (inside the shape)', { do: 'delete', range: '6:7' }),
  step('insert column A', { do: 'insert', range: 'A:A' }),
  step('insert columns C:D (inside the table)', { do: 'insert', range: 'C:D' }),
  step('delete column C (the table\'s Qty)', { do: 'delete', range: 'C:C' }),
  step('delete columns F:K (all of the chart)', { do: 'delete', range: 'F:K' }),
  step('delete columns B:D (all of the table)', { do: 'delete', range: 'B:D' }),
  step('delete column N (inside the shape)', { do: 'delete', range: 'N:N' }),
  step('insert B5:D5 shifting down (the table\'s width)', { do: 'insert', range: 'B5:D5', shift: 'down' }),
  step('insert C5 shifting down (part of the table)', { do: 'insert', range: 'C5', shift: 'down' }),
  step('delete B5:D6 shifting up (the table\'s width)', { do: 'delete', range: 'B5:D6', shift: 'up' }),
  step('insert F5 shifting down (under the chart)', { do: 'insert', range: 'F5', shift: 'down' }),
  step('insert A6 shifting right (the note at B6)', { do: 'insert', range: 'A6', shift: 'right' }),
];
console.log(JSON.stringify({ sheet, cases }, null, 1));
