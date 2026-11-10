// The cases of data validation measured in Excel (excel-dv.ps1: Validation.Add, then Range.Validation.Value for each
// cell): whole numbers, decimals, lists (typed and from cells), dates, times, text length and custom formulas, against
// numbers, text, numbers stored as text, booleans, errors and blanks.
// usage: node tools/excel-compare/dv-cases.mjs > tools/excel-compare/dv-cases.json
const col = (values) => Object.fromEntries(values.map((v, i) => [`A${i + 1}`, v]).filter(([, v]) => v !== ''));
const read = (n) => `A1:A${n}`;
const cases = [];
const add = (name, values, rule, extra = {}) => cases.push({ name, setup: { cells: { ...col(values), ...extra.cells }, formats: extra.formats }, rule: { range: read(values.length), ...rule }, read: read(values.length) });

const mixed = ['0', '1', '5', '10', '11', '5.5', 'abc', '', "'7", 'TRUE', '=1/0', '=2+3', '-3'];
add('whole number between 1 and 10', mixed, { type: 'whole', op: 'between', f1: '1', f2: '10' });
add('whole number not between 1 and 10', mixed, { type: 'whole', op: 'notBetween', f1: '1', f2: '10' });
add('whole number greater than 4', mixed, { type: 'whole', op: 'greater', f1: '4' });
add('whole number, blanks not ignored', mixed, { type: 'whole', op: 'between', f1: '1', f2: '10', ignoreBlank: false });
add('decimal between 0 and 1', ['0', '0.5', '1', '1.5', '-0.1', 'x', '', "'0.5"], { type: 'decimal', op: 'between', f1: '0', f2: '1' });
add('decimal equal to a cell', ['3', '4', '3.0', 'x'], { type: 'decimal', op: 'equal', f1: '=$B$1' }, { cells: { B1: '3' } });
add('list of words', ['a', 'A', 'b', 'd', 'b ', ' b', '', 'c,d'], { type: 'list', f1: 'a,b,c' });
add('list of numbers', ['1', "'1", '2', '2.0', '4', 'one'], { type: 'list', f1: '1,2,3' });
add('list from cells', ['y', 'Y', 'w', '', '5', '5.0'], { type: 'list', f1: '=$B$1:$B$4' }, { cells: { B1: 'x', B2: 'y', B3: 'z', B4: '5' } });
add('list, blanks not ignored', ['a', ''], { type: 'list', f1: 'a,b', ignoreBlank: false });
add('date on or after 1/1/2024', ['1/1/2024', '12/31/2023', '6/15/2024', '45292', '45291', 'x', "'1/1/2024"], { type: 'date', op: 'greaterEqual', f1: '1/1/2024' });
add('time between 9:00 and 17:00', ['9:00', '8:59', '12:30', '17:00', '17:01', '0.5', '1.5'], { type: 'time', op: 'between', f1: '9:00', f2: '17:00' });
add('text length at most 3', ['abc', 'abcd', '123', '1234', '12.5', '', 'TRUE', '=1/0'], { type: 'length', op: 'lessEqual', f1: '3' });
add('custom: a number', ['1', 'x', '', 'TRUE', '=1/0'], { type: 'custom', f1: '=ISNUMBER(A1)' });
add('custom: relative to the row', ['5', '1', '3', ''], { type: 'custom', f1: '=A1>B1' }, { cells: { B1: '4', B2: '4', B3: '3', B4: '0' } });

process.stdout.write(JSON.stringify(cases, null, 1) + '\n');
