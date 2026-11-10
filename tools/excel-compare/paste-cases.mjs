// The cases of Paste Special and Paste Link measured in Excel (excel-steps.ps1): A1:C3 copied (numbers, text, a
// blank, relative and $ formulas, a percentage; A1 bold, filled, two decimals and a bottom border; column A wide) onto
// E1:G3, which has its own values, a formula, a text, a blank, a fill and a format.
// usage: node tools/excel-compare/paste-cases.mjs > tools/excel-compare/paste-cases.json
const base = () => ({
  cells: { A1: '1', B1: '2', C1: 't', A2: '=A1*2', B2: '=$A$1+B1', A3: '5', B3: '=SUM(A1:B1)', C3: '7',
    E1: '10', F1: '20', G1: 'x', E2: '40', F2: '=E1+1', G2: '50', F3: '60', G3: '70' },
  formats: [{ range: 'A1', bold: true, color: 65535, numFmt: '0.00', border: true }, { range: 'A3', numFmt: '0%' },
    { range: 'F1', color: 5287936 }, { range: 'G2', numFmt: '0.0' }, { range: 'E3', bold: true }],
  widths: { A: 20 },
});
const read = 'A1:H16';
const paste = (more) => [{ do: 'paste', from: 'A1:C3', to: 'E1', ...more }];
const cases = [
  ['all', paste({ paste: 'all' })],
  ['formulas', paste({ paste: 'formulas' })],
  ['values', paste({ paste: 'values' })],
  ['formats', paste({ paste: 'formats' })],
  ['column widths', paste({ paste: 'widths' })],
  ['formulas and number formats', paste({ paste: 'formulasFormats' })],
  ['values and number formats', paste({ paste: 'valuesFormats' })],
  ['all except borders', paste({ paste: 'noBorders' })],
  ['values, add', paste({ paste: 'values', op: 'add' })],
  ['all, subtract', paste({ paste: 'all', op: 'subtract' })],
  ['values, multiply', paste({ paste: 'values', op: 'multiply' })],
  ['values, divide', paste({ paste: 'values', op: 'divide' })],
  ['formulas, add', paste({ paste: 'formulas', op: 'add' })],
  ['all, skip blanks', paste({ paste: 'all', skip: true })],
  ['all, transpose', paste({ paste: 'all', transpose: true })],
  ['formulas, transpose', paste({ paste: 'formulas', transpose: true })],
  ['values, transpose, skip blanks', paste({ paste: 'values', transpose: true, skip: true })],
  ['formats, add', paste({ paste: 'formats', op: 'add' })],
  ['values onto a bigger selection', [{ do: 'paste', from: 'A1:C3', to: 'E10:G15', paste: 'values' }]],
  ['all onto one row', [{ do: 'paste', from: 'A1:C1', to: 'E10:J10', paste: 'all' }]],
  ['link', [{ do: 'link', from: 'A1:C3', to: 'E1' }]],
  ['link one cell', [{ do: 'link', from: 'A1', to: 'E5' }]],
].map(([name, steps]) => ({ name, setup: base(), steps, read }));
// A second sheet for what the first doesn't show: text and blanks under an operation, formulas reading cells outside
// the copy (D9), negative numbers, a border on the destination, and Paste Link's formats.
const more = () => ({
  cells: { A1: 'txt', B1: '3', C1: '=D9*2', A2: '=$D$9+1', B2: '=B1+D9', C2: '4', D9: '100',
    E1: '5', F1: 'y', E2: '-5', F2: '=TRUE', G2: 'z' },
  formats: [{ range: 'F2', border: true }, { range: 'B1', numFmt: '0%' }],
});
const paste2 = (more2) => [{ do: 'paste', from: 'A1:C2', to: 'E1', ...more2 }];
cases.push(...[
  ['text and blanks, values, add', paste2({ paste: 'values', op: 'add' })],
  ['text and blanks, formulas, add', paste2({ paste: 'formulas', op: 'add' })],
  ['text and blanks, all, multiply', paste2({ paste: 'all', op: 'multiply' })],
  ['outside references, transpose', paste2({ paste: 'formulas', transpose: true })],
  ['outside references, all', paste2({ paste: 'all' })],
  ['border kept, all except borders', paste2({ paste: 'noBorders' })],
  ['link from a general cell', [{ do: 'link', from: 'B1:C2', to: 'E5' }]],
  ['link onto a bigger selection', [{ do: 'link', from: 'B1:C1', to: 'E5:H6' }]],
].map(([name, steps]) => ({ name, setup: more(), steps, read })));
process.stdout.write(JSON.stringify(cases, null, 1) + '\n');
