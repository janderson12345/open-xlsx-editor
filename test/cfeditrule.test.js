import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { build, area } from './steps.js';
import { serializeXml } from '../src/ooxml/xml.js';
import { editRule } from '../src/ooxml/cfmanage.js';

// Edit Rule against Excel (test/cfedit-excel.json, excel-cf-edit.ps1: FormatCondition.Modify on rules of A1:A10): the
// rule in its place in the file, its type and criteria new, its format kept. COM also sends it last and sets Stop If
// True; the editor keeps its priority and its own Stop If True, as the Rules Manager shows it (docs/CONDFORMAT.md).
const EXCEL = JSON.parse(fs.readFileSync('test/cfedit-excel.json', 'utf8'));
const cells = Object.fromEntries(Array.from({ length: 10 }, (_, i) => [`A${i + 1}`, String(i + 1)]));
const A = [area('A1:A10')];
const shape = (xml) => [...xml.matchAll(/<cfRule\b[\s\S]*?<\/cfRule>/g)].map((m) => m[0].replace(/ (priority|stopIfTrue|dxfId)="\d+"/g, ''));
const dxfIds = (xml) => [...xml.matchAll(/<cfRule\b[^>]*?dxfId="(\d+)"/g)].map((m) => Number(m[1]));
const CASES = {
  'greater than 5 to between 3 and 7': [[{ kind: 'cellIs', operator: 'greaterThan', f1: '5' }, { kind: 'expression', formula: '=A1=2' }], 0, { kind: 'cellIs', operator: 'between', f1: '3', f2: '7' }],
  'greater than 5 to a formula': [[{ kind: 'cellIs', operator: 'greaterThan', f1: '5' }, { kind: 'expression', formula: '=A1=2' }], 0, { kind: 'expression', formula: '=MOD(A1,2)=0' }],
  'formula to less than 4': [[{ kind: 'cellIs', operator: 'greaterThan', f1: '5' }, { kind: 'expression', formula: '=A1=2' }], 1, { kind: 'cellIs', operator: 'lessThan', f1: '4' }],
  'the middle of three, stop if true off': [[{ kind: 'cellIs', operator: 'greaterThan', f1: '8' }, { kind: 'cellIs', operator: 'lessThan', f1: '3' }, { kind: 'expression', formula: '=A1=5' }], 1, { kind: 'cellIs', operator: 'equal', f1: '6' }],
};

for (const x of EXCEL) {
  test(`Edit Rule: ${x.name}, as Excel edits it`, () => {
    const [rules, k, change] = CASES[x.name];
    const { wb, si } = build({ cells });
    for (const r of rules) wb.setConditional(si, A, { ...r, preset: 'redFill' });
    const cfXml = () => [...wb._part(si).doc.documentElement.childNodes].filter((n) => n.localName === 'conditionalFormatting').map((n) => serializeXml(n).replace(/ xmlns(:\w+)?="[^"]*"/g, '')).join('');
    const before = cfXml();
    const p = Number([...before.matchAll(/priority="(\d+)"/g)][k][1]);
    const priorities = [...before.matchAll(/priority="(\d+)"/g)].map((m) => m[1]);
    wb.changeConditional(si, () => editRule(wb, si, p, { ...change, preset: null }));
    const after = cfXml();
    assert.deepEqual(shape(after), shape(x.conditionalFormatting), 'the rules, in their places, the edited one new');
    assert.deepEqual(dxfIds(after), dxfIds(before), 'each keeping its format');
    assert.deepEqual([...after.matchAll(/priority="(\d+)"/g)].map((m) => m[1]), priorities, 'and its priority');
    assert.equal(wb.sheet(si).condFormats[0].rules[k].priority, p);
  });
}

test('Edit Rule: a new format when one is picked, Stop If True kept, a data bar given the place of the rule it replaces', () => {
  const { wb, si } = build({ cells });
  wb.setConditional(si, A, { kind: 'cellIs', operator: 'greaterThan', f1: '5', preset: 'redFill' });
  wb.setConditional(si, A, { kind: 'expression', formula: '=A1=2', preset: 'redFill' });
  wb.setConditional(si, A, { kind: 'average', preset: 'redFill' });
  const doc = () => wb._part(si).doc;
  const ruleAt = (p) => [...doc().getElementsByTagNameNS('*', 'cfRule')].find((e) => e.getAttribute('priority') === String(p));
  ruleAt(2).setAttribute('stopIfTrue', '1');
  const dxfs = wb.styles.dxfs.length;
  wb.changeConditional(si, () => editRule(wb, si, 2, { kind: 'cellIs', operator: 'lessThan', f1: '3', preset: 'green' }));
  assert.equal(ruleAt(2).getAttribute('operator'), 'lessThan');
  assert.equal(ruleAt(2).getAttribute('stopIfTrue'), '1');
  assert.equal(Number(ruleAt(2).getAttribute('dxfId')), dxfs, "the green preset's new dxf");
  // (The rule of priority 2 a data bar: it comes after priority 1 and before the one after it.)
  wb.changeConditional(si, () => editRule(wb, si, 2, { kind: 'dataBar', color: 'FF638EC6', solid: false }));
  const order = wb.sheet(si).condFormats.flatMap((cf) => cf.rules).sort((a, b) => a.priority - b.priority).map((r) => r.type);
  assert.deepEqual(order, ['aboveAverage', 'dataBar', 'cellIs']);
});
