// The cases of conditional formatting measured in Excel (excel-cf.ps1: FormatConditions, then what each cell shows by
// Range.DisplayFormat): cell values compared, formulas, text, top and bottom, average, duplicates, blanks and
// errors, colour scales, several rules on one range, and Stop If True.
// usage: node tools/excel-compare/cf-cases.mjs > tools/excel-compare/cf-cases.json
const RED = 255;
const GREEN = 65280;
const YELLOW = 65535;
const BLUE_TEXT = 16711680;
const nums = { A1: '5', A2: '12', A3: '7', A4: '20', A5: '7', A6: '1', A7: '15', A8: '', A9: '=1/0', A10: 'x' };
const words = { A1: 'apple', A2: 'Pineapple', A3: 'pear', A4: 'grape', A5: 'APPLE pie', A6: '', A7: '42', A8: 'crab apple' };
const cases = [];
const add = (name, cells, rules, read = 'A1:B10', formats) => cases.push({ name, setup: { cells, formats }, rules, read });
const fill = (f) => ({ fill: f });

add('greater than', nums, [{ range: 'A1:A10', kind: 'cell', op: 'greater', f1: '10', ...fill(RED) }]);
add('between', nums, [{ range: 'A1:A10', kind: 'cell', op: 'between', f1: '5', f2: '12', ...fill(GREEN) }]);
add('not between', nums, [{ range: 'A1:A10', kind: 'cell', op: 'notBetween', f1: '5', f2: '12', ...fill(GREEN) }]);
add('equal, with text and blanks around', nums, [{ range: 'A1:A10', kind: 'cell', op: 'equal', f1: '7', ...fill(YELLOW), bold: true }]);
add('not equal', nums, [{ range: 'A1:A10', kind: 'cell', op: 'notEqual', f1: '7', ...fill(YELLOW) }]);
add('less or equal, to a cell', { ...nums, B1: '7' }, [{ range: 'A1:A10', kind: 'cell', op: 'lessEqual', f1: '=$B$1', color: BLUE_TEXT }]);
add('equal to text', words, [{ range: 'A1:A8', kind: 'cell', op: 'equal', f1: '="apple"', ...fill(RED) }]);
add('a formula, relative', { ...nums, B1: '4', B2: '20', B3: '7', B4: '1', B5: '9', B6: '1', B7: '15' }, [{ range: 'A1:A7', kind: 'expr', f1: '=A1>B1', ...fill(GREEN) }]);
add('a formula over the row', { ...nums, B1: 'y', B2: 'n', B3: 'y' }, [{ range: 'A1:B3', kind: 'expr', f1: '=$B1="y"', ...fill(YELLOW) }]);
add('text contains', words, [{ range: 'A1:A8', kind: 'text', op: 'contains', text: 'apple', ...fill(RED) }]);
add('text does not contain', words, [{ range: 'A1:A8', kind: 'text', op: 'notContains', text: 'apple', ...fill(RED) }]);
add('text begins with', words, [{ range: 'A1:A8', kind: 'text', op: 'begins', text: 'ap', ...fill(RED) }]);
add('text ends with', words, [{ range: 'A1:A8', kind: 'text', op: 'ends', text: 'apple', ...fill(RED) }]);
add('top 3', nums, [{ range: 'A1:A10', kind: 'top', rank: 3, ...fill(GREEN) }]);
add('bottom 2', nums, [{ range: 'A1:A10', kind: 'top', rank: 2, bottom: true, ...fill(RED) }]);
add('top 25 percent', nums, [{ range: 'A1:A10', kind: 'top', rank: 25, percent: true, ...fill(GREEN) }]);
add('above average', nums, [{ range: 'A1:A10', kind: 'average', how: 0, ...fill(GREEN) }]);
add('below average', nums, [{ range: 'A1:A10', kind: 'average', how: 1, ...fill(RED) }]);
add('equal or above average', { A1: '1', A2: '2', A3: '3', A4: '2' }, [{ range: 'A1:A4', kind: 'average', how: 2, ...fill(GREEN) }]);
add('duplicates', { ...words, A2: 'apple', A3: 'APPLE' }, [{ range: 'A1:A8', kind: 'unique', dup: true, ...fill(RED) }]);
add('unique values', { ...nums }, [{ range: 'A1:A10', kind: 'unique', dup: false, ...fill(GREEN) }]);
add('blanks', nums, [{ range: 'A1:A10', kind: 'blanks', ...fill(YELLOW) }]);
add('no blanks', nums, [{ range: 'A1:A10', kind: 'noBlanks', ...fill(YELLOW) }]);
add('errors', nums, [{ range: 'A1:A10', kind: 'errors', ...fill(RED) }]);
add('two rules, both apply', nums, [{ range: 'A1:A10', kind: 'cell', op: 'greater', f1: '10', ...fill(RED) }, { range: 'A1:A10', kind: 'cell', op: 'greater', f1: '6', bold: true, ...fill(GREEN) }]);
add('two rules, stop if true', nums, [{ range: 'A1:A10', kind: 'cell', op: 'greater', f1: '6', bold: true, ...fill(GREEN) }, { range: 'A1:A10', kind: 'cell', op: 'greater', f1: '10', ...fill(RED), stop: true }]);
add('a rule over formatted cells', nums, [{ range: 'A1:A10', kind: 'cell', op: 'greater', f1: '10', color: BLUE_TEXT }], 'A1:B10', [{ range: 'A1:A3', numFmt: '0.00' }]);
add('two-colour scale', { A1: '0', A2: '25', A3: '50', A4: '75', A5: '100', A6: '' }, [{ range: 'A1:A6', kind: 'scale', colors: [16777215, 6526841] }]);
add('three-colour scale', { A1: '1', A2: '2', A3: '3', A4: '4', A5: '5', A6: '9' }, [{ range: 'A1:A6', kind: 'scale', colors: [7039480, 8711167, 8109667] }]);

process.stdout.write(JSON.stringify(cases, null, 1) + '\n');
