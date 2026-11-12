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
  assert.deepEqual(bold, { text: 'bold', font: { name: 'Aptos Narrow', size: 11, scheme: 'minor', b: true, i: false, u: null, strike: false, color: { rgb: null, theme: 1, indexed: null, auto: false, tint: 0 }, vertAlign: null } });
  assert.deepEqual([big.font.name, big.font.size, big.font.scheme, big.font.color.rgb], ['Courier New', 16, null, 'FFFF0000']);
  const h2o = list.runs.get(1);
  assert.deepEqual(h2o.map((r) => [r.text, r.font?.vertAlign ?? null, r.font?.u ?? null]), [['H', null, null], ['2', 'subscript', null], ['O', null, 'single']]);
});
