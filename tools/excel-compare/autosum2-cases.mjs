// More cases of AutoSum measured in Excel (excel-autosum2.ps1, the ribbon's AutoSum and its menu on a selection of
// several cells): in a filtered list and a filtered table (SUBTOTAL), over rows hidden by hand, a table not filtered,
// several areas, text among the numbers selected, a blank column beside a block. None leaves Excel editing a cell
// (one cell selected, or nothing found), which would stop it answering over COM. (Several blank rows under the numbers,
// B7:C9, did: Excel opened B7 for editing as it does one cell, and answered nothing more; docs/AUTOSUM.md.)
// usage: node tools/excel-compare/autosum2-cases.mjs > tools/excel-compare/autosum2-cases.json
const list = { A1: 'Fruit', B1: 'Qty', C1: 'Amt', A2: 'Apple', A3: 'Pear', A4: 'Apple', A5: 'Plum', A6: 'Apple',
  B2: '1', B3: '2', B4: '3', B5: '4', B6: '5', C2: '10', C3: '20', C4: '30', C5: '40', C6: '50' };
const apples = { range: 'A1:C6', field: 1, criteria: 'Apple' };
const cases = [];
const add = (name, setup, select, { fn = 'sum', read = 'A1:F12' } = {}) => cases.push({ name, setup, select, fn, read });

add('filtered list: a column with the blank under it', { cells: list, autofilter: apples }, 'B2:B7');
add('filtered list: the blank row under it', { cells: list, autofilter: apples }, 'B7:C7');
add('filtered list: numbers in two columns selected', { cells: list, autofilter: apples }, 'B2:C6');
add('filtered list: totals both ways', { cells: list, autofilter: apples }, 'B2:D7');
for (const fn of ['average', 'count', 'max', 'min']) add(`filtered list: ${fn} under it`, { cells: list, autofilter: apples }, 'B7:C7', { fn });
add('a list with a filter that hides nothing', { cells: list, autofilter: { range: 'A1:C6' } }, 'B7:C7');
add('rows hidden by hand, no filter', { cells: list, hideRows: '3:3' }, 'B7:C7');
add('a table under it, filtered', { cells: list, table: { range: 'A1:C6', name: 'Fruit' }, tableFilter: { field: 1, criteria: 'Apple' } }, 'B7:C7');
add('a table under it, not filtered', { cells: list, table: { range: 'A1:C6', name: 'Fruit' } }, 'B7:C7');
add('a table: its numbers selected', { cells: list, table: { range: 'A1:C6', name: 'Fruit' } }, 'B2:C6');
add('two areas of numbers', { cells: { ...list, E1: 'More', E2: '7', E3: '8', E4: '9' } }, 'B2:B6,E2:E4');
add('text among the numbers selected', { cells: { ...list, B4: 'x' } }, 'B2:B6');
add('a blank column beside a block', { cells: list }, 'B2:D6');
add('a blank row inside the numbers selected', { cells: { ...list, B4: '', C4: '' } }, 'B2:C6');

process.stdout.write(JSON.stringify(cases, null, 1) + '\n');
