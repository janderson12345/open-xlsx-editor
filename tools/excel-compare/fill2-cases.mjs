// The fill handle's Auto Fill Options for dates (Fill Days, Fill Weekdays, Fill Months, Fill Years), measured in Excel
// (excel-steps.ps1, AutoFill's xlFillDays to xlFillYears): a date or two in column A filled down.
// usage: node tools/excel-compare/fill2-cases.mjs > tools/excel-compare/fill2-cases.json
const read = 'A1:H12';
const col = (...texts) => Object.fromEntries(texts.map((t, i) => [`A${i + 1}`, t]));
const down = (cells, n, type) => ({ cells, steps: [{ do: 'autofill', from: `A1:A${Object.keys(cells).length}`, to: `A1:A${n}`, type }] });
const cases = [
  ['a date, days', down(col('1/30/2024'), 5, 'days')],
  ['a date, weekdays', down(col('11/12/2026'), 7, 'weekdays')],
  ['a Saturday, weekdays', down(col('11/14/2026'), 4, 'weekdays')],
  ['a date, months', down(col('1/31/2024'), 6, 'months')],
  ['a date, years', down(col('2/29/2024'), 4, 'years')],
  ['dates a week apart, months', down(col('1/1/2024', '1/8/2024'), 5, 'months')],
  ['dates two days apart, weekdays', down(col('11/12/2026', '11/14/2026'), 6, 'weekdays')],
  ['dates a month apart, days', down(col('1/31/2024', '2/29/2024'), 5, 'days')],
  ['a date and time, days', down(col('1/30/2024 9:30'), 4, 'days')],
  ['three dates in a row, weekdays', down(col('11/16/2026', '11/17/2026', '11/18/2026'), 9, 'weekdays')],
  ['two weekdays a day apart, weekdays', down(col('11/16/2026', '11/17/2026'), 8, 'weekdays')],
  ['three dates, months', down(col('1/15/2024', '1/20/2024', '1/31/2024'), 9, 'months')],
  ['two years apart, years', down(col('3/1/2020', '3/1/2022'), 5, 'years')],
  ['three dates, days', down(col('1/1/2024', '1/2/2024', '1/4/2024'), 6, 'days')],
  ['a date with a format, years', { cells: col('3/15/2020'), formats: [{ range: 'A1', numFmt: 'd-mmm-yy', bold: true }], steps: [{ do: 'autofill', from: 'A1', to: 'A1:A4', type: 'years' }] }],
].map(([name, { cells, formats, steps }]) => ({ name, setup: { cells, ...(formats ? { formats } : {}) }, steps, read }));
process.stdout.write(JSON.stringify(cases, null, 1) + '\n');
