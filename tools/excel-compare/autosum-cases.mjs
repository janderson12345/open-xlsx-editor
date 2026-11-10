// The cases of AutoSum measured in Excel (excel-autosum.ps1: the ribbon's AutoSum, Average, Count Numbers, Max and
// Min on a selection of several cells): where the formula goes, which cells it takes, past blanks, text and subtotals,
// and the format it is given. (Where nothing is found, as over dates, Excel leaves the cell being edited, =AVERAGE()
// waiting for Enter, and answers nothing over COM: seen on screen, not cased here.)
// usage: node tools/excel-compare/autosum-cases.mjs > tools/excel-compare/autosum-cases.json
const table = { B1: 'Qty', C1: 'Amt', B2: '1', B3: '2', B4: '3', B5: '4', C2: '10', C3: '20', C4: '30', C5: '40' };
const cases = [];
const add = (name, cells, select, { fn = 'sum', formats, read = 'A1:F10' } = {}) => cases.push({ name, setup: { cells, formats }, select, fn, read });

add('numbers in a column selected', table, 'B2:B5');
add('numbers in two columns selected', table, 'B2:C5');
add('a column with the blank under it', table, 'B2:B6');
add('two columns with the blank row under them', table, 'B2:C6');
add('with a blank row and a blank column', table, 'B2:D6');
add('the blank row under the table', table, 'B6:C6');
add('the blank column beside the table', table, 'D2:D5');
add('numbers in a row selected', { B8: '1', C8: '2', D8: '3', E8: '4' }, 'B8:E8');
add('a row with the blank beside it', { B8: '1', C8: '2', D8: '3', E8: '4' }, 'B8:F8');
add('a blank in the column', { ...table, B4: '' , C4: '' }, 'B6:C6');
add('text in the column', { ...table, B4: 'x', C4: 'y' }, 'B6:C6');
add('a number over the numbers', { ...table, B1: '2024', C1: '2025' }, 'B6:C6');
add('subtotals in the column', { B1: 'Qty', B2: '1', B3: '2', B4: '=SUM(B2:B3)', B5: '5', B6: '6', B7: '=SUM(B5:B6)', C1: 'Amt', C2: '10', C3: '20', C4: '=SUM(C2:C3)', C5: '50', C6: '60', C7: '=SUM(C5:C6)' }, 'B8:C8');
add('a gap of two rows', { ...table, B6: '', D1: 'x' }, 'B7:C7');
add('the cell under is filled', { ...table, B6: '99' }, 'B2:B5');
add('currency numbers', table, 'B6:C6', { formats: [{ range: 'B2:B5', numFmt: '$#,##0.00' }, { range: 'C2:C5', numFmt: '0%' }] });
for (const fn of ['average', 'count', 'max', 'min']) add(`${fn} under the table`, table, 'B6:C6', { fn });
add('count over text and numbers', { ...table, B3: 'x' }, 'B6:C6', { fn: 'count' });
add('a whole block with totals both ways', table, 'B2:D6', { fn: 'average' });

process.stdout.write(JSON.stringify(cases, null, 1) + '\n');
