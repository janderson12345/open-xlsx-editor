// The cases of inserting and deleting cells measured in Excel (excel-steps.ps1): one sheet of numbers in B1:C12, and
// formulas reading them in column H, on Sheet2 and in defined names, with each step done to a new copy of it.
// usage: node tools/excel-compare/shift-cases.mjs > tools/excel-compare/shift-cases.json
const formulas = ['=B5', '=C7', '=SUM(B1:B10)', '=SUM(B1:C10)', '=SUM(A1:D10)', '=SUM(B1:B2)', '=B$5', '=$B$5', '=SUM(B5:C5)',
  '=A5', '=D5', '=SUM(B3:B3)', '=SUM(B3:C4)', '=SUM(B3:D4)', '=SUM(B2:C3)', '=SUM(A3:C3)', '=SUM(B:B)', '=SUM(3:3)',
  '=SUM(B5:B12)', '=C12', '=D3', '=E4', '=SUM(B4:C12)', '=SUM(A1:C1)'];
const base = () => {
  const cells = {};
  for (let r = 1; r <= 12; r++) { cells[`B${r}`] = String(r); cells[`C${r}`] = String(r * 10); }
  cells.C11 = '=B5+C1';
  cells.D3 = 'x';
  cells.E4 = '7';
  formulas.forEach((f, i) => { cells[`H${i + 1}`] = f; });
  return {
    cells,
    sheet2: { A1: '=Sheet1!B5', A2: '=SUM(Sheet1!B1:B10)', A3: '=Sheet1!D3', A4: '=Sheet1!B3:C4' },
    names: { Four: '=Sheet1!$B$4', Band: '=Sheet1!$B$3:$C$4', Col: '=Sheet1!$B$1:$B$10' },
    formats: [{ range: 'B2', color: 65535 }, { range: 'C2', bold: true }, { range: 'B3', color: 5287936 },
      { range: 'A3', color: 15773696 }, { range: 'A4', bold: true }],
  };
};
const with_ = (more) => { const s = base(); for (const [k, v] of Object.entries(more)) s[k] = Array.isArray(v) ? [...(s[k] || []), ...v] : { ...s[k], ...v }; return s; };
const read = 'A1:L26';
const cases = [
  ['insert B3:C4 down', [{ do: 'insert', range: 'B3:C4', shift: 'down' }]],
  ['insert B3:C4 right', [{ do: 'insert', range: 'B3:C4', shift: 'right' }]],
  ['delete B3:C4 up', [{ do: 'delete', range: 'B3:C4', shift: 'up' }]],
  ['delete B3:C4 left', [{ do: 'delete', range: 'B3:C4', shift: 'left' }]],
  ['insert B3 down', [{ do: 'insert', range: 'B3', shift: 'down' }]],
  ['delete B5:B6 up', [{ do: 'delete', range: 'B5:B6', shift: 'up' }]],
  ['insert B3:C4 shape', [{ do: 'insert', range: 'B3:C4' }]],
  ['insert B3:D4 shape', [{ do: 'insert', range: 'B3:D4' }]],
  ['insert B3:C5 shape', [{ do: 'insert', range: 'B3:C5' }]],
  ['insert B3 shape', [{ do: 'insert', range: 'B3' }]],
  ['delete B3:D4 shape', [{ do: 'delete', range: 'B3:D4' }]],
  ['delete B3:C5 shape', [{ do: 'delete', range: 'B3:C5' }]],
  ['delete B3:C4 shape', [{ do: 'delete', range: 'B3:C4' }]],
  ['merge inside, insert down', [{ do: 'insert', range: 'B3:C4', shift: 'down' }], { merges: ['B7:C7'] }],
  ['merge across the edge, insert down', [{ do: 'insert', range: 'B3:C4', shift: 'down' }], { merges: ['A7:B7'] }],
  ['merge above, insert down', [{ do: 'insert', range: 'B3:C4', shift: 'down' }], { merges: ['B1:C1'] }],
  ['merge across the edge, insert right', [{ do: 'insert', range: 'B3:C4', shift: 'right' }], { merges: ['A3:B3'] }],
  ['merge inside, insert right', [{ do: 'insert', range: 'B3:C4', shift: 'right' }], { merges: ['D3:E3'] }],
  ['merge across the rows, insert right', [{ do: 'insert', range: 'B3:C4', shift: 'right' }], { merges: ['F3:F5'] }],
  ['array inside, insert right', [{ do: 'insert', range: 'B3:C4', shift: 'right' }], { arrays: { 'J3:J4': '=B3:B4*2' } }],
  ['array across the rows, insert right', [{ do: 'insert', range: 'B3:C4', shift: 'right' }], { arrays: { 'J4:J5': '=B4:B5*2' } }],
  ['array below, insert down', [{ do: 'insert', range: 'B3:C4', shift: 'down' }], { arrays: { 'B14:B15': '=C1:C2*2' } }],
  ['filled last row, insert down', [{ do: 'insert', range: 'B3', shift: 'down' }], { cells: { B1048576: '1' } }],
  ['merge deleted, delete up', [{ do: 'delete', range: 'B3:C4', shift: 'up' }], { merges: ['B3:C4'] }],
  ['merge shrinks, delete up', [{ do: 'delete', range: 'B3:C3', shift: 'up' }], { merges: ['B4:B5'] }],
].map(([name, steps, more]) => ({ name, setup: more ? with_(more) : base(), steps, read }));
process.stdout.write(JSON.stringify(cases, null, 1) + '\n');
