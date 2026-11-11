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
// With the table, more about tables: its edges, its header row, part of it, a formula reading a column of it.
if (sheet.table) {
  sheet.formula = { at: 'A20', text: '=SUM(Stock[Qty])+COUNTA(Stock[Price])' };
  cases.push(
    step('insert row 2 (above the header)', { do: 'insert', range: '2:2' }),
    step('insert row 13 (just below the table)', { do: 'insert', range: '13:13' }),
    step('insert row 12 (above the last row)', { do: 'insert', range: '12:12' }),
    step('delete row 2 (the header row)', { do: 'delete', range: '2:2' }),
    step('delete rows 1:3 (the header and a row)', { do: 'delete', range: '1:3' }),
    step('delete rows 3:12 (all of the data)', { do: 'delete', range: '3:12' }),
    step('insert column B (at the left edge of the table)', { do: 'insert', range: 'B:B' }),
    step('insert column E (at the right edge of the table)', { do: 'insert', range: 'E:E' }),
    step('insert column D (before the last column)', { do: 'insert', range: 'D:D' }),
    step('delete column D (the last column)', { do: 'delete', range: 'D:D' }),
    step('delete columns C:E (two table columns and one after)', { do: 'delete', range: 'C:E' }),
    step('insert B13:D13 shifting down (under the table)', { do: 'insert', range: 'B13:D13', shift: 'down' }),
    step('insert E3:E4 shifting right (beside the table)', { do: 'insert', range: 'E3:E4', shift: 'right' }),
    step('insert B3:D3 shifting right (the table width)', { do: 'insert', range: 'B3:D3', shift: 'right' }),
    step('delete B2:D2 shifting up (the header row cells)', { do: 'delete', range: 'B2:D2', shift: 'up' }),
  );
}
console.log(JSON.stringify({ sheet, cases }, null, 1));
