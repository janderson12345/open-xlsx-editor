// The cases of Home > Clear measured in Excel (excel-steps.ps1: Range.Clear, ClearContents, ClearFormats): a small
// table with numbers, a formula, formats, a border and merged ranges, part of it cleared.
// usage: node tools/excel-compare/clear-cases.mjs > tools/excel-compare/clear-cases.json
const read = 'A1:F10';
const table = () => ({
  cells: { A1: 'Name', B1: 'Q1', A2: 'a', B2: '1', C2: '10', B3: '=B2*2', D8: 'merged', A9: 'kept' },
  formats: [{ range: 'B2:B3', bold: true, numFmt: '0.00' }, { range: 'C2', color: 65535, border: true }, { range: 'A9', bold: true }],
  merges: ['D8:E8', 'A9:B9'],
});
const cases = [
  ['clear all', [{ do: 'clear', range: 'B2:E8', what: 'all' }]],
  ['clear contents', [{ do: 'clear', range: 'B2:E8', what: 'contents' }]],
  ['clear formats', [{ do: 'clear', range: 'B2:E8', what: 'formats' }]],
  ['clear all, part of a merge', [{ do: 'clear', range: 'B9:C10', what: 'all' }]],
  ['clear formats, part of a merge', [{ do: 'clear', range: 'B9:C10', what: 'formats' }]],
  ['clear contents, part of a merge', [{ do: 'clear', range: 'B9:C10', what: 'contents' }]],
].map(([name, steps]) => ({ name, setup: table(), steps, read }));
process.stdout.write(JSON.stringify(cases, null, 1) + '\n');
