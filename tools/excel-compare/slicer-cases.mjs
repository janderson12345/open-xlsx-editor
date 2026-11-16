// Inserting and deleting rows, columns and cells on a sheet with slicers and a timeline, for excel-slicers.ps1: a
// table Sales on B2:E14 (made-up dates, regions, items and units), a slicer on its Region column over G2:I11, a
// PivotTable of it at L20 with a slicer on Item over P2:R11 and a timeline on Date over L2:N9; one step each.
// usage: node tools/excel-compare/slicer-cases.mjs > tools/excel-compare/slicer-cases.json
const regions = ['North', 'South', 'East', 'West'];
const items = ['Bolt', 'Nut', 'Pin'];
// (Dates as Excel's serial numbers: 2026-01-05 is 46027, a fortnight apart.)
const rows = [['Date', 'Region', 'Item', 'Units'], ...Array.from({ length: 12 }, (_, i) => [46027 + i * 14, regions[i % 4], items[i % 3], (i * 7) % 11 + 1])];
const sheet = { rows, at: 'B2', tableSlicer: 'G2:I11', pivot: 'L20', pivotSlicer: 'P2:R11', timeline: 'L2:N9' };
const step = (name, s) => ({ name, step: s });
const cases = [
  step('start', { do: 'none', range: 'A1' }),
  step('insert row 1', { do: 'insert', range: '1:1' }),
  step('insert rows 5:6 (inside the table and the slicers)', { do: 'insert', range: '5:6' }),
  step('delete row 5 (inside the table and the slicers)', { do: 'delete', range: '5:5' }),
  step('delete rows 2:14 (all of the table, the slicers\' tops)', { do: 'delete', range: '2:14' }),
  step('delete rows 2:11 (all of the slicers)', { do: 'delete', range: '2:11' }),
  step('insert column A', { do: 'insert', range: 'A:A' }),
  step('insert column C (inside the table, before Region)', { do: 'insert', range: 'C:C' }),
  step('delete column B (the table\'s Date, the timeline\'s field)', { do: 'delete', range: 'B:B' }),
  step('delete column C (the table\'s Region, the slicer\'s column)', { do: 'delete', range: 'C:C' }),
  step('delete column D (the table\'s Item, the pivot slicer\'s field)', { do: 'delete', range: 'D:D' }),
  step('delete columns G:I (all of the table slicer)', { do: 'delete', range: 'G:I' }),
  step('insert column H (inside the table slicer)', { do: 'insert', range: 'H:H' }),
  step('insert G5 shifting down (under the slicer\'s top)', { do: 'insert', range: 'G5', shift: 'down' }),
  step('delete rows 21:22 (inside the pivot)', { do: 'delete', range: '21:22' }),
];
console.log(JSON.stringify({ sheet, cases }, null, 1));
