// Filter, measured in Excel (excel-filter.ps1): sorting while rows are filtered out (AutoFilter.Sort: only the rows
// shown are sorted), and filtering dates by year, month or day (xlFilterValues with Criteria2's date groups).
// usage: node tools/excel-compare/filter2-cases.mjs > tools/excel-compare/filter2-cases.json
const table = {
  A1: 'Name', B1: 'Qty', C1: 'Team',
  A2: 'pear', B2: '5', C2: 'red', A3: 'fig', B3: '2', C3: 'blue', A4: 'kiwi', B4: '9', C4: 'red',
  A5: 'apple', B5: '1', C5: 'blue', A6: 'plum', B6: '7', C6: 'red', A7: 'date', B7: '3', C7: 'green',
};
const dates = {
  A1: 'When', B1: 'Row', A2: '1/5/2024', A3: '2/10/2024', A4: '2/20/2024', A5: '3/1/2023', A6: '2/10/2023', A7: '12/31/2024',
  ...Object.fromEntries([2, 3, 4, 5, 6, 7].map((r) => [`B${r}`, `r${r}`])),
};
const on = { do: 'on', cell: 'A1' };
const cases = [];
const add = (name, steps, cells, values = false) => cases.push({ name, setup: { cells }, steps, read: 'A1:C10', ...(values ? { values: true } : {}) });
add('red rows sorted A to Z while filtered', [on, { do: 'filter', range: 'A1', field: 3, values: ['red'] }, { do: 'sort', key: 'A1' }], table, true);
add('red rows sorted by Qty, largest first, while filtered', [on, { do: 'filter', range: 'A1', field: 3, values: ['red'] }, { do: 'sort', key: 'B1', order: 'desc' }], table, true);
add('dates: the year 2024', [on, { do: 'dates', range: 'A1', field: 1, groups: [[0, '1/1/2024']] }], dates);
add('dates: February 2024', [on, { do: 'dates', range: 'A1', field: 1, groups: [[1, '2/1/2024']] }], dates);
add('dates: a day, and the year 2023', [on, { do: 'dates', range: 'A1', field: 1, groups: [[2, '2/10/2024'], [0, '1/1/2023']] }], dates);
process.stdout.write(JSON.stringify(cases, null, 1) + '\n');
