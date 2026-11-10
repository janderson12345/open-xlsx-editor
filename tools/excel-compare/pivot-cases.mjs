// Pivot table cases for excel-pivot.ps1: one made-up table of sales (region, product, quarter, units, price) and, for
// each case, the fields Excel's PivotTable puts in rows, columns and values. Excel builds each pivot in a new workbook,
// reads its cells, changes the source (a number and a region), refreshes it and reads them again.
// usage: node tools/excel-compare/pivot-cases.mjs > tools/excel-compare/pivot-cases.json
const header = ['Region', 'Product', 'Quarter', 'Units', 'Price'];
const regions = ['North', 'South', 'East'];
const products = ['Apple', 'Pear', 'Plum', 'Fig'];
const quarters = ['Q1', 'Q2', 'Q3', 'Q4'];
// (A fixed sequence, so the cases are the same each time.)
let seed = 7;
const next = () => { seed = (seed * 48271) % 2147483647; return seed; };
const rows = [];
for (let i = 0; i < 18; i++) rows.push([regions[next() % 3], products[next() % 4], quarters[next() % 4], 1 + (next() % 40), (5 + (next() % 60)) / 4]);
const data = [header, ...rows];
const sum = (field, caption = null) => ({ field, fn: 'sum', caption });
const cases = [
  { name: 'rows Region, sum of Units', rows: ['Region'], cols: [], values: [sum('Units')] },
  { name: 'rows Region then Product, sum of Units', rows: ['Region', 'Product'], cols: [], values: [sum('Units')] },
  { name: 'rows Product, columns Quarter, sum of Units', rows: ['Product'], cols: ['Quarter'], values: [sum('Units')] },
  { name: 'rows Region, sum of Units and average of Price', rows: ['Region'], cols: [], values: [sum('Units'), { field: 'Price', fn: 'average' }] },
  { name: 'rows Region, count of Product', rows: ['Region'], cols: [], values: [{ field: 'Product', fn: 'count' }] },
  { name: 'rows Product, max of Units and min of Price', rows: ['Product'], cols: [], values: [{ field: 'Units', fn: 'max' }, { field: 'Price', fn: 'min' }] },
  { name: 'columns Region only, sum of Units', rows: [], cols: ['Region'], values: [sum('Units')] },
  { name: 'rows Quarter then Region, columns Product, sum of Units', rows: ['Quarter', 'Region'], cols: ['Product'], values: [sum('Units')] },
];
// The change before refreshing: the first row's units to 100 and its region to a new one, Central (before the others A to Z: where Excel puts a new item).
const change = [{ r: 1, c: 3, value: 100 }, { r: 1, c: 0, value: 'Central' }];
console.log(JSON.stringify({ data, change, cases }, null, 1));
