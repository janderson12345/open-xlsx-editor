// Dynamic array formulas (Excel 365's spilling), for excel-dynamic.ps1: each case a formula entered with Formula2 in
// D1 of a sheet whose A1:A6 holds numbers and B1:B6 words (`block`: a cell typed first, in the spill's way; `then`:
// another formula after it, reading the spill).
// usage: node tools/excel-compare/dynamic-cases.mjs > tools/excel-compare/dynamic-cases.json
const one = (name, formula, extra = {}) => ({ name, formula, ...extra });
const cases = [
  one('SEQUENCE(4)', '=SEQUENCE(4)'),
  one('SEQUENCE(2,3,10,5)', '=SEQUENCE(2,3,10,5)'),
  one('SORT a column of numbers', '=SORT(A1:A6)'),
  one('SORT words, descending', '=SORT(B1:B6,,-1)'),
  one('SORT two columns by the second', '=SORT(A1:B6,2)'),
  one('SORTBY words by numbers', '=SORTBY(B1:B6,A1:A6)'),
  one('UNIQUE numbers', '=UNIQUE(A1:A6)'),
  one('UNIQUE words, exactly once', '=UNIQUE(B1:B6,FALSE,TRUE)'),
  one('FILTER words by numbers over 4', '=FILTER(B1:B6,A1:A6>4)'),
  one('FILTER finding nothing, with if_empty', '=FILTER(B1:B6,A1:A6>100,"none")'),
  one('FILTER finding nothing', '=FILTER(B1:B6,A1:A6>100)'),
  one('a range times 2', '=A1:A3*2'),
  one('TRANSPOSE across', '=TRANSPOSE(A1:A3)'),
  one('a spill blocked', '=SEQUENCE(4)', { block: { ref: 'D3', value: 'x' } }),
  one('a spill read by D1#', '=SEQUENCE(4)', { then: { ref: 'F1', formula: '=SUM(D1#)' } }),
  one('LET', '=LET(x,A1,y,A2,x+y)'),
  one('LAMBDA called at once', '=LAMBDA(x,x*2)(A1)'),
  one('@ on a range (row 1)', '=@A1:A3'),
  one('SEQUENCE then a sum over the spill', '=SUM(SEQUENCE(4))'),
  one('a spill of one cell', '=SEQUENCE(1)'),
];
console.log(JSON.stringify({ cases }, null, 1));
