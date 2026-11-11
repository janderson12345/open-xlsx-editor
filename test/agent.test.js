import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { openWorkbook } from '../src/agent.js';

const book = () => openWorkbook(fs.readFileSync('test/fixtures/formulas.xlsx'));

test('an agent reads the sheets, names and tables, and a range by reference, name or table', () => {
  const a = book();
  assert.deepEqual(a.sheets().map((s) => [s.name, s.used, s.tables.join()]), [['Calc', 'A2:F1387', ''], ['Data', 'A1:Q9', 'Fruit'], ['My Data', 'A1:A2', '']]);
  assert.deepEqual(a.names().map((n) => [n.name, n.refersTo, n.table]), [['Fruit', 'Data!A1:E9', true], ['Items', 'Data!$A$2:$A$9', false], ['Rate', 'Data!$C$2', false]]);
  const qty = a.read('Fruit[Qty]');
  assert.equal(qty.range, 'Data!B2:B9');
  assert.deepEqual(qty.rows.map((row) => row[0].value), [3, 5, '12', 2, 8, null, 4, 6], 'B4 is a number stored as text in the fixture');
  assert.deepEqual(a.read('Items').rows.map((row) => row[0].text).slice(0, 2), ['Apple', 'Banana']);
  assert.deepEqual(a.read('B2', { sheet: 'Data' }).rows, [[{ ref: 'B2', value: 3, text: '3', formula: null }]]);
  assert.equal(a.read("'My Data'!A1:A2").range, "'My Data'!A1:A2", 'a sheet name with a space comes back quoted');
  assert.throws(() => a.read('Nope!A1'), /not a range/);
  const whole = a.read('Data!B:B');
  assert.equal(whole.range, 'Data!B1:B9', 'a whole column is cut to the used rows');
});

test('an agent gets a Markdown table, finds cells and reads formulas', () => {
  const a = book();
  const md = a.markdown('Data!A1:B3');
  assert.equal(md, 'Data!A1:B3\n\n|  | A | B |\n|---|---|---|\n| 1 | Name | Qty |\n| 2 | Apple | 3 |\n| 3 | Banana | 5 |');
  assert.deepEqual(a.find('apple', { lookIn: 'values', within: 'sheet', sheet: 'Data' }).map((h) => h.ref), ['Data!A2', 'Data!L2', 'Data!A5', 'Data!A9']);
  const f = a.read('Calc!B2').rows[0][0];
  assert.equal(f.formula, '=Data!B2:B9');
});

test('an agent enters text as typed: numbers, percentages, formulas, refusals; recalculates what reads it', () => {
  const a = book();
  const out = a.write({ 'Data!B2': '7', 'Data!G2': '50%', 'Data!H2': '=Data!B2*2', 'Data!I2': "'007", 'Data!A1:A2': 'x', 'Data!J2': '=1+' });
  assert.deepEqual(out.written, ['Data!B2', 'Data!G2', 'Data!H2', 'Data!I2']);
  assert.deepEqual(out.refused.map((x) => x.ref), ['Data!A1:A2', 'Data!J2']);
  const row = a.read('Data!G2:I2').rows[0];
  assert.deepEqual(row.map((x) => [x.value, x.text]), [[0.5, '50%'], [14, '14'], ['007', '007']]);
  assert.equal(row[1].formula, '=Data!B2*2');
  assert.equal(a.read('Fruit[Qty]').rows[0][0].value, 7);
});

test('an agent writes blocks of values, inserts rows, adds and renames sheets, and saves', () => {
  const a = book();
  const name = a.addSheet('Summary');
  assert.equal(name, 'Summary');
  assert.equal(a.setValues('Summary!A1', [['Fruit', 'Qty'], ['=not a formula', 3], [true, null]]), 'Summary!A1:B3');
  assert.deepEqual(a.read('Summary!A1:B3').rows.map((r) => r.map((x) => x.value)), [['Fruit', 'Qty'], ['=not a formula', 3], [true, null]]);
  a.insertRows(2, 1, { sheet: 'Summary' });
  assert.deepEqual(a.read('Summary!A1:A4').rows.map((r) => r[0].value), ['Fruit', null, '=not a formula', true]);
  a.insertRows(3, 1, { sheet: 'Data' });
  assert.equal(a.names().find((x) => x.name === 'Fruit').refersTo, 'Data!A1:E10', 'a row inserted inside the Fruit table widens it');
  a.renameSheet('Summary', 'Totals');
  assert.throws(() => a.renameSheet('Totals', 'Data'), /already|taken|name/i);
  const again = openWorkbook(a.save());
  assert.deepEqual(again.sheets().map((s) => s.name), ['Calc', 'Data', 'My Data', 'Totals']);
  assert.equal(again.read('Totals!A3').rows[0][0].value, '=not a formula');
  again.deleteSheet('Totals');
  assert.equal(again.sheets().length, 3);
});

test("an agent's entry is refused by a cell's data validation, as Excel refuses it, unless forced", async () => {
  const { build } = await import('./steps.js');
  const { AgentWorkbook } = await import('../src/agent.js');
  const { parseXml, kid } = await import('../src/ooxml/xml.js');
  const { readSheet } = await import('../src/ooxml/sheet.js');
  const { wb, si } = build({ cells: { B1: 'x' } });
  const { doc } = wb._part(si);
  const root = doc.documentElement;
  const dv = parseXml(`<dataValidations xmlns="${root.namespaceURI}" count="1"><dataValidation type="whole" allowBlank="1" showErrorMessage="1" error="1 to 10" sqref="A1:A5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>`).documentElement;
  root.insertBefore(doc.importNode(dv, true), kid(root, 'pageMargins') || null);
  wb._parsed.set(wb.sheets[si], { doc, model: readSheet(doc, wb.sharedStrings) });
  const a = new AgentWorkbook(wb);
  assert.deepEqual(a.write({ A1: '5', A2: '11', A3: '2.5' }, { sheet: si }), { written: [`${wb.sheets[si].name}!A1`], refused: [{ ref: 'A2', error: '1 to 10' }, { ref: 'A3', error: '1 to 10' }] });
  assert.deepEqual(a.write({ A2: '11' }, { sheet: si, force: true }).written.length, 1, 'forced, as Yes in Excel');
});
