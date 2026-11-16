import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseXml } from '../src/ooxml/xml.js';
import { readSharedStrings } from '../src/ooxml/strings.js';

// (The sharedStrings Excel 365 saved after Characters(...).Font set over COM: bold, a colour, a size, sub- and
// superscript, another font, underline; the first run without rPr, as Excel writes it.)
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const rPr = (inner) => `<rPr>${inner}<sz val="11"/><color theme="1"/><rFont val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></rPr>`;
const SST = `<sst xmlns="${NS}" count="3" uniqueCount="3">`
  + `<si><r><t xml:space="preserve">Plain </t></r><r>${rPr('<b/>')}<t>bold</t></r><r><rPr><sz val="16"/><color rgb="FFFF0000"/><rFont val="Courier New"/><family val="3"/></rPr><t>big</t></r></si>`
  + `<si><r><t>H</t></r><r>${rPr('<vertAlign val="subscript"/>')}<t>2</t></r><r>${rPr('<u/>')}<t>O</t></r></si>`
  + '<si><t>plain</t></si></sst>';

test("rich text as Excel writes it: each run's text and font, the first one the cell's own", () => {
  const list = readSharedStrings(parseXml(SST));
  assert.deepEqual([...list], ['Plain boldbig', 'H2O', 'plain']);
  assert.deepEqual([...list.runs.keys()], [0, 1]);
  const [plain, bold, big] = list.runs.get(0);
  assert.deepEqual(plain, { text: 'Plain ', font: null });
  assert.deepEqual(bold, { text: 'bold', font: { name: 'Aptos Narrow', size: 11, scheme: 'minor', b: true, i: false, u: null, strike: false, color: { rgb: null, theme: 1, indexed: null, auto: false, tint: 0 }, vertAlign: null, family: 2 } });
  assert.deepEqual([big.font.name, big.font.size, big.font.scheme, big.font.color.rgb], ['Courier New', 16, null, 'FFFF0000']);
  const h2o = list.runs.get(1);
  assert.deepEqual(h2o.map((r) => [r.text, r.font?.vertAlign ?? null, r.font?.u ?? null]), [['H', null, null], ['2', 'subscript', null], ['O', null, 'single']]);
});

test("rich text written as Excel writes it: the cell's own font's run bare, the others with every property", async () => {
  const { build } = await import('./steps.js');
  const { serializeXml } = await import('../src/ooxml/xml.js');
  const { wb, si } = build({ cells: { A1: 'x' } });
  wb.apply(si, [{ r: 1, c: 0, value: 'Plain boldbig', runs: [{ text: 'Plain ', font: null }, { text: 'bold', font: { b: true } }, { text: 'big', font: { size: 16, color: 'FFFF0000', name: 'Courier New' } }] }]);
  const el = [...wb._part(si).doc.documentElement.getElementsByTagNameNS(NS, 'c')].find((c) => c.getAttribute('r') === 'A2');
  const runs = [...el.getElementsByTagNameNS(NS, 'r')].map((r) => serializeXml(r).replace(/ xmlns="[^"]*"/g, ''));
  // (Excel's own, from the sharedStrings it saved (SST above): the base font here is the fixture's, Calibri 11.)
  const base = wb._runBase(0);
  assert.equal(runs[0], '<r><t xml:space="preserve">Plain </t></r>');
  assert.equal(runs[1], `<r><rPr><b/><sz val="${base.size}"/>${base.color?.theme != null ? `<color theme="${base.color.theme}"/>` : ''}<rFont val="${base.name}"/>${base.family != null ? `<family val="${base.family}"/>` : ''}${base.scheme ? `<scheme val="${base.scheme}"/>` : ''}</rPr><t>bold</t></r>`);
  assert.equal(runs[2], '<r><rPr><sz val="16"/><color rgb="FFFF0000"/><rFont val="Courier New"/></rPr><t>big</t></r>');
  // (And read back as runs.)
  assert.deepEqual(wb.sheet(si).cells.get(1).get(0).runs.map((r) => [r.text, !!r.font?.b, r.font?.size ?? null]), [['Plain ', false, null], ['bold', true, base.size], ['big', false, 16]]);
});

test('a run read from a file written back with its own colour and family, as the file had them', async () => {
  const { richRunsXml } = await import('../src/ooxml/strings.js');
  const base = { name: 'Calibri', size: 11, color: { theme: 1 }, family: 2, scheme: 'minor' };
  // (runFont's colour is an element's attributes, not an 'FFRRGGBB' string.)
  const red = { name: 'Courier New', size: 16, i: true, family: 3, color: { rgb: 'FFFF0000', theme: null, indexed: null, auto: false, tint: 0 } };
  const themed = { name: 'Calibri', b: true, color: { rgb: null, theme: 4, indexed: null, auto: false, tint: 0.4 } };
  assert.equal(richRunsXml([{ text: 'red', font: red }, { text: 'blue', font: themed }], base),
    '<r><rPr><i/><sz val="16"/><color rgb="FFFF0000"/><rFont val="Courier New"/><family val="3"/></rPr><t>red</t></r>'
    + '<r><rPr><b/><sz val="11"/><color theme="4" tint="0.4"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t>blue</t></r>');
});

test("a run's rPr children not modelled (outline, shadow, condense, extend, charset) kept and written back in Excel's order", async () => {
  const { build } = await import('./steps.js');
  const { serializeXml } = await import('../src/ooxml/xml.js');
  const { stringItemRuns } = await import('../src/ooxml/strings.js');
  const si = parseXml(`<si xmlns="${NS}"><r><t>a</t></r><r><rPr><b/><condense val="0"/><outline/><shadow/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><charset val="161"/><scheme val="minor"/></rPr><t>b</t></r></si>`).documentElement;
  const runs = stringItemRuns(si, NS);
  assert.deepEqual(runs[1].font.extra, ['<condense val="0"/>', '<outline/>', '<shadow/>', '<charset val="161"/>']);
  const { wb, si: sheet } = build({ cells: { A1: 'x' } });
  wb.apply(sheet, [{ r: 0, c: 0, value: 'ab', runs }]);
  const el = [...wb._part(sheet).doc.documentElement.getElementsByTagNameNS(NS, 'c')].find((c) => c.getAttribute('r') === 'A1');
  const written = serializeXml([...el.getElementsByTagNameNS(NS, 'rPr')][0]).replace(/ xmlns="[^"]*"/g, '');
  assert.equal(written, '<rPr><b/><condense val="0"/><outline/><shadow/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><charset val="161"/><scheme val="minor"/></rPr>');
});
