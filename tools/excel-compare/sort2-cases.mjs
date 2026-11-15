// Sort On Cell Color and Font Color, and a custom list as the order, measured in Excel (excel-steps.ps1: Worksheet.Sort
// with SortFields.Add's SortOn, SortOnValue.Color and CustomOrder, as Excel's Sort box applies them).
// usage: node tools/excel-compare/sort2-cases.mjs > tools/excel-compare/sort2-cases.json
const YELLOW = 65535;
const GREEN = 5296274;
const RED = 255;
const data = { A1: 'Item', B1: 'Row', A2: 'pear', A3: 'fig', A4: 'kiwi', A5: 'apple', A6: 'plum', A7: 'date', ...Object.fromEntries([2, 3, 4, 5, 6, 7].map((r) => [`B${r}`, `r${r}`])) };
const fills = [{ range: 'A3', color: YELLOW }, { range: 'A5', color: GREEN }, { range: 'A6', color: YELLOW }];
const inks = [{ range: 'A4', fontColor: RED }, { range: 'A7', fontColor: RED }];
const sort = (keys) => [{ do: 'sort', range: 'A1:B7', keys, header: 'yes' }];
const levels = { A1: 'Level', B1: 'Row', A2: 'Low', A3: 'High', A4: 'None', A5: 'Medium', A6: 'high', A7: '', A8: 'Low', ...Object.fromEntries([2, 3, 4, 5, 6, 7, 8].map((r) => [`B${r}`, `r${r}`])) };
const cases = [
  ['cell colour yellow on top', { cells: data, formats: fills }, sort([{ key: 'A1', on: 'cellColor', color: YELLOW }])],
  ['cell colour yellow on bottom', { cells: data, formats: fills }, sort([{ key: 'A1', on: 'cellColor', color: YELLOW, order: 'desc' }])],
  ['font colour red on top', { cells: data, formats: inks }, sort([{ key: 'A1', on: 'fontColor', color: RED }])],
  ['yellow on top, then green on top, then A to Z', { cells: data, formats: fills }, sort([{ key: 'A1', on: 'cellColor', color: YELLOW }, { key: 'A1', on: 'cellColor', color: GREEN }, { key: 'A1' }])],
  ['custom list High, Medium, Low', { cells: levels }, [{ do: 'sort', range: 'A1:B8', keys: [{ key: 'A1', customOrder: 'High,Medium,Low' }], header: 'yes' }]],
  ['custom list High, Medium, Low, descending', { cells: levels }, [{ do: 'sort', range: 'A1:B8', keys: [{ key: 'A1', customOrder: 'High,Medium,Low', order: 'desc' }], header: 'yes' }]],
].map(([name, setup, steps]) => ({ name, setup, steps, read: 'A1:H12' }));
process.stdout.write(JSON.stringify(cases, null, 1) + '\n');
