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

test("a rule's own format (New Rule's Format...): its dxf as Excel writes it, its number format and borders drawn", async () => {
  const { conditionalLooks } = await import('../src/ooxml/condformat.js');
  const X = Object.fromEntries(JSON.parse(fs.readFileSync('test/cfformat-excel.json', 'utf8')).map((c) => [c.name, c]));
  const FORMATS = {
    'number format 0.00%': { numFmt: '0.00%' },
    'bold, italic, red text, green fill': { bold: true, italic: true, color: 'FFFF0000', fill: 'FF92D050' },
    'thin border all round': { border: { left: { style: 'thin' }, right: { style: 'thin' }, top: { style: 'thin' }, bottom: { style: 'thin' } } },
    'red bottom border': { border: { bottom: { style: 'thin', color: 'FFFF0000' } } },
    'dashed top border': { border: { top: { style: 'dashed' } } },
    'all of it': { numFmt: '#,##0.0', bold: true, fill: 'FFFFFF00', border: { left: { style: 'thin' }, right: { style: 'thin' }, top: { style: 'thin' }, bottom: { style: 'thin' } } },
  };
  // (Excel's numFmtId for a built-in code in a dxf (14 for 0.00%) and the <i val="0"/> COM's Font leaves are set aside.)
  const norm = (xml) => xml.replace(/ numFmtId="\d+"/g, '').replace(/<i val="0"\/>/g, '');
  for (const [name, format] of Object.entries(FORMATS)) {
    const { wb, si } = build({ cells });
    const before = wb.styles.dxfs.length;
    wb.setConditional(si, A, { kind: 'cellIs', operator: 'greaterThan', f1: '5', preset: format });
    const added = [...wb.stylesDoc.documentElement.getElementsByTagNameNS('*', 'dxf')].slice(before).map((n) => serializeXml(n).replace(/ xmlns(:\w+)?="[^"]*"/g, '')).join('');
    assert.equal(norm(added), norm(/<dxf>[\s\S]*<\/dxf>/.exec(X[name].dxfs)[0]), name);
    const look = conditionalLooks(wb, si).get('6,0');
    if (format.numFmt) assert.equal(look.numFmt, format.numFmt, `${name}: the number format`);
    if (format.border) assert.deepEqual(Object.keys(look.border).sort(), Object.keys(format.border).sort(), `${name}: the sides`);
    assert.equal(conditionalLooks(wb, si).get('1,0'), undefined, `${name}: not on 2`);
  }
});

test('a colour scale with thresholds and colours of its own: written and shown as Excel does (cfscale-excel.json)', async () => {
  const { conditionalLooks } = await import('../src/ooxml/condformat.js');
  const X = JSON.parse(fs.readFileSync('test/cfscale-excel.json', 'utf8'));
  const STOPS = {
    '3-color: number 2, percent 50, highest': [{ type: 'num', val: 2, color: 'FFFF0000' }, { type: 'percent', val: 50, color: 'FFFFFF00' }, { type: 'max', color: 'FF008000' }],
    '2-color: percentile 10, percent 90': [{ type: 'percentile', val: 10, color: 'FFFFFFFF' }, { type: 'percent', val: 90, color: 'FF0070C0' }],
    '2-color: lowest, number 6': [{ type: 'min', color: 'FFFFFFCC' }, { type: 'num', val: 6, color: 'FF006600' }],
  };
  const hex = (n) => `#${[n & 255, (n >> 8) & 255, (n >> 16) & 255].map((x) => x.toString(16).padStart(2, '0')).join('')}`;
  const rgb = (h) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16));
  for (const x of X) {
    const { wb, si } = build({ cells });
    wb.setConditional(si, A, { kind: 'colorScale', points: STOPS[x.name].length, stops: STOPS[x.name] });
    const cf = [...wb._part(si).doc.documentElement.childNodes].filter((n) => n.localName === 'conditionalFormatting').map((n) => serializeXml(n).replace(/ xmlns(:\w+)?="[^"]*"/g, '')).join('');
    assert.equal(cf.replace(/ priority="\d+"/, ''), x.conditionalFormatting.replace(/ priority="\d+"/, ''), `${x.name}: the rule`);
    const looks = conditionalLooks(wb, si);
    for (const [ref, r] of [['A1', 0], ['A5', 4], ['A10', 9]]) {
      const mine = rgb(looks.get(`${r},0`).fill.toLowerCase());
      const theirs = rgb(hex(x.shown[ref]));
      assert.ok(mine.every((v, i) => Math.abs(v - theirs[i]) <= 1), `${x.name}, ${ref}: ${mine} against Excel's ${theirs}`);
    }
  }
});

test("Excel's twelve Color Scales: the gallery's colours in its order, as Excel applies them (cfgallery-excel.json)", async () => {
  const { CF_GALLERY } = await import('../src/ooxml/cfedit.js');
  const X = JSON.parse(fs.readFileSync('test/cfgallery-excel.json', 'utf8'));
  assert.deepEqual(CF_GALLERY.map(([, colors]) => colors.map((c) => c.slice(2))), X.presets);
});
