// The cases of Filter measured in Excel (excel-filter.ps1: Range.AutoFilter, then the XML Excel saves): turning it on
// from one cell, filtering by shown values (text, numbers as shown, blanks, case), two columns, and turning it off.
// usage: node tools/excel-compare/filter-cases.mjs > tools/excel-compare/filter-cases.json
const table = {
  A1: 'Name', B1: 'Qty', C1: 'Team',
  A2: 'apple', B2: '1', C2: 'red',
  A3: 'pear', B3: '2', C3: 'blue',
  A4: 'Apple', B4: '3', C4: 'red',
  A5: 'fig', B5: '1', C5: '',
  A6: 'plum', B6: '2.5', C6: 'green',
  A7: 'kiwi', B7: '10', C7: 'blue',
};
const cases = [];
const add = (name, steps, setup = { cells: table }) => cases.push({ name, setup, steps, read: 'A1:C10' });
const on = { do: 'on', cell: 'B3' };

add('turned on from one cell', [on]);
add('turned on, then off', [on, on]);
add('one text value', [on, { do: 'filter', range: 'A1', field: 3, values: ['red'] }]);
add('two text values', [on, { do: 'filter', range: 'A1', field: 3, values: ['red', 'green'] }]);
add('text in another case', [on, { do: 'filter', range: 'A1', field: 1, values: ['apple'] }]);
add('numbers by their text', [on, { do: 'filter', range: 'A1', field: 2, values: ['1', '10'] }]);
add('blanks', [on, { do: 'filter', range: 'A1', field: 3, values: ['='] }]);
add('blanks and a value', [on, { do: 'filter', range: 'A1', field: 3, values: ['=', 'blue'] }]);
add('two columns', [on, { do: 'filter', range: 'A1', field: 3, values: ['red', 'blue'] }, { do: 'filter', range: 'A1', field: 2, values: ['1'] }]);
add('a filter cleared', [on, { do: 'filter', range: 'A1', field: 3, values: ['red'] }, { do: 'filter', range: 'A1', field: 3, values: ['red', 'blue', 'green', '='] }]);
add('numbers shown with decimals', [on, { do: 'filter', range: 'A1', field: 2, values: ['1.00', '2.50'] }], { cells: table, formats: [{ range: 'B2:B7', numFmt: '0.00' }] });
add('one column cleared', [on, { do: 'filter', range: 'A1', field: 3, values: ['red'] }, { do: 'filter', range: 'A1', field: 2, values: ['1'] }, { do: 'clear', range: 'A1', field: 3 }]);
add('on, then off with a filter set', [on, { do: 'filter', range: 'A1', field: 3, values: ['red'] }, on]);

process.stdout.write(JSON.stringify(cases, null, 1) + '\n');
