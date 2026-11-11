// Go To Special's choices on test/fixtures/special.xlsx, for excel-special.ps1 (Excel's SpecialCells and the Range
// methods behind the box) and test/special.test.js (ours): each a selection (ranges, comma separated), its active
// cell, the choice and its option (types for Constants and Formulas: numbers 1, text 2, logicals 4, errors 16).
// usage: node tools/excel-compare/special-cases.mjs > tools/excel-compare/special-cases.json
const ALL = 1 | 2 | 4 | 16;
const one = (at, kind, extra = {}) => ({ select: at, active: at, kind, ...extra });
const cases = [];
const add = (name, c) => cases.push({ name, ...c });
for (const [word, bits] of [['all', ALL], ['numbers', 1], ['text', 2], ['logicals', 4], ['errors', 16], ['numbers and text', 3]]) {
  add(`constants, ${word}, from A1`, one('A1', 'constants', { types: bits }));
  add(`formulas, ${word}, from A1`, one('A1', 'formulas', { types: bits }));
}
add('constants, numbers, in A1:D6', { select: 'A1:D6', active: 'A1', kind: 'constants', types: 1 });
add('constants, all, in two ranges', { select: 'A1:B3,D2:D6', active: 'D2', kind: 'constants', types: ALL });
add('formulas, all, in C1:E4', { select: 'C1:E4', active: 'C1', kind: 'formulas', types: ALL });
add('blanks, from A1', one('A1', 'blanks'));
add('blanks, in A1:E8', { select: 'A1:E8', active: 'A1', kind: 'blanks' });
add('blanks, in A8:C10 (a merged range)', { select: 'A8:C10', active: 'A8', kind: 'blanks' });
add('notes, from A1', one('A1', 'notes'));
add('notes, in B1:E6', { select: 'B1:E6', active: 'B1', kind: 'notes' });
add('current region, B3', one('B3', 'region'));
add('current region, J12', one('J12', 'region'));
add('current region, an empty cell next to data (F7)', one('F7', 'region'));
add('current region, an empty cell apart (M20)', one('M20', 'region'));
add('current array, G3', one('G3', 'array'));
add('current array, not in one (A1)', one('A1', 'array'));
add('row differences, A12:C16 by column A', { select: 'A12:C16', active: 'A12', kind: 'rowDiff' });
add('row differences, A12:C16 by column B', { select: 'A12:C16', active: 'B13', kind: 'rowDiff' });
add('column differences, A12:C16 by row 12', { select: 'A12:C16', active: 'A12', kind: 'colDiff' });
add('column differences, A12:C16 by row 14', { select: 'A12:C16', active: 'B14', kind: 'colDiff' });
add('row differences, A17:C17 (case)', { select: 'A17:C17', active: 'A17', kind: 'rowDiff' });
add('row differences, A21:C21 (relative formulas)', { select: 'A21:C21', active: 'A21', kind: 'rowDiff' });
add('column differences, A20:C21 by row 20', { select: 'A20:C21', active: 'A20', kind: 'colDiff' });
add('precedents, direct, E3', one('E3', 'precedents', { levels: 'direct' }));
add('precedents, all levels, E3', one('E3', 'precedents', { levels: 'all' }));
add('precedents, direct, C2:C3', { select: 'C2:C3', active: 'C2', kind: 'precedents', levels: 'direct' });
add('dependents, direct, B2', one('B2', 'dependents', { levels: 'direct' }));
add('dependents, all levels, B2', one('B2', 'dependents', { levels: 'all' }));
add('dependents, all levels, B4', one('B4', 'dependents', { levels: 'all' }));
add('precedents, none (A1)', one('A1', 'precedents', { levels: 'direct' }));
add('last cell, from A1', one('A1', 'lastCell'));
add('visible cells only, from A1', one('A1', 'visible'));
add('visible cells only, in A1:H8', { select: 'A1:H8', active: 'A1', kind: 'visible' });
add('visible cells only, in A1:B3 (nothing hidden)', { select: 'A1:B3', active: 'A1', kind: 'visible' });
add('conditional formats, all, from A1', one('A1', 'cfAll'));
add('conditional formats, same as C3', one('C3', 'cfSame'));
add('conditional formats, all, in A1:B6', { select: 'A1:B6', active: 'A1', kind: 'cfAll' });
add('data validation, all, from A1', one('A1', 'dvAll'));
add('data validation, same as B3', one('B3', 'dvSame'));
add('data validation, same as D2', one('D2', 'dvSame'));
add('data validation, same as A1 (none)', one('A1', 'dvSame'));
console.log(JSON.stringify({ cases }, null, 1));
