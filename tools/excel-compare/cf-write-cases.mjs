// The rules of Excel's Conditional Formatting menu, for excel-cf-write.ps1: each added (as the menu adds it, with one
// of its preset formats) to A1:A10 (numbers 1 to 10) or B1:B10 (words), alone or after another to see the priorities.
// usage: node tools/excel-compare/cf-write-cases.mjs > tools/excel-compare/cf-write-cases.json
const A = 'A1:A10';
const B = 'B1:B10';
const one = (name, rule) => ({ name, rules: [rule] });
const cases = [
  one('greater than 5, light red fill with dark red text', { range: A, kind: 'cellIs', operator: 'greaterThan', f1: '5', preset: 'lightRed' }),
  one('less than 3, yellow fill with dark yellow text', { range: A, kind: 'cellIs', operator: 'lessThan', f1: '3', preset: 'yellow' }),
  one('between 3 and 7, green fill with dark green text', { range: A, kind: 'cellIs', operator: 'between', f1: '3', f2: '7', preset: 'green' }),
  one('equal to "apple", light red fill', { range: B, kind: 'cellIs', operator: 'equal', f1: '="apple"', preset: 'redFill' }),
  one('text that contains "an", red text', { range: B, kind: 'containsText', text: 'an', preset: 'redText' }),
  one('duplicate values, light red fill with dark red text', { range: B, kind: 'duplicate', preset: 'lightRed' }),
  one('top 3 items', { range: A, kind: 'top10', rank: 3, preset: 'lightRed' }),
  one('top 20%', { range: A, kind: 'top10', rank: 20, percent: true, preset: 'lightRed' }),
  one('bottom 2 items', { range: A, kind: 'top10', rank: 2, bottom: true, preset: 'lightRed' }),
  one('above average', { range: A, kind: 'average', preset: 'lightRed' }),
  one('below average', { range: A, kind: 'average', below: true, preset: 'yellow' }),
  one('3-color scale', { range: A, kind: 'colorScale', points: 3 }),
  one('2-color scale', { range: A, kind: 'colorScale', points: 2 }),
  { name: 'greater than 5, then less than 3 on the same cells', rules: [
    { range: A, kind: 'cellIs', operator: 'greaterThan', f1: '5', preset: 'lightRed' },
    { range: A, kind: 'cellIs', operator: 'lessThan', f1: '3', preset: 'green' }] },
];
console.log(JSON.stringify({ cases }, null, 1));
