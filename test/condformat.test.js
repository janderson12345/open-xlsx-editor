import './setup.js';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseXml, kid } from '../src/ooxml/xml.js';
import { readSheet } from '../src/ooxml/sheet.js';
import { readStyles } from '../src/ooxml/styles.js';
import { conditionalLooks } from '../src/ooxml/condformat.js';
import { readCases, build, area } from './steps.js';

// Conditional formatting, each case as Excel drew it (FormatConditions over COM, then each cell's
// Range.DisplayFormat: tools/excel-compare/excel-cf.ps1 on cf-cases.json, from cf-cases.mjs, giving
// test/cf-excel.json): the case's sheet built here with the conditionalFormatting elements and dxfs Excel saved for
// it, and every cell's fill, font colour and bold compared with what Excel showed.
const cases = readCases('tools/excel-compare/cf-cases.json', 'test/cf-excel.json');
const X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
/** Excel's colour number (blue, green, red from the top byte down) as "#RRGGBB". */
const hex = (n) => `#${[n & 255, (n >> 8) & 255, (n >> 16) & 255].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase()}`;

/** The case's sheet with Excel's rules and formats put in, read again. */
function withRules(c) {
  const built = build(c.setup);
  const { wb, si } = built;
  const { doc } = wb._part(si);
  const root = doc.documentElement;
  const before = kid(root, 'pageMargins') || null;
  for (const xml of c.excel.conditionalFormatting) {
    const el = parseXml(xml.replace('<conditionalFormatting ', `<conditionalFormatting xmlns="${X}" `)).documentElement;
    root.insertBefore(doc.importNode(el, true), before);
  }
  if (c.excel.dxfs) {
    const styles = wb.stylesDoc.documentElement;
    const old = kid(styles, 'dxfs');
    const el = wb.stylesDoc.importNode(parseXml(c.excel.dxfs.replace('<dxfs ', `<dxfs xmlns="${X}" `)).documentElement, true);
    if (old) styles.replaceChild(el, old); else styles.appendChild(el);
    wb.styles = readStyles(wb.stylesDoc);
  }
  wb._parsed.set(wb.sheets[si], { doc, model: readSheet(doc, wb.sharedStrings) });
  wb.calculate();
  return built;
}

for (const c of cases) {
  test(`conditional format: ${c.name}, as Excel`, () => {
    const { wb, si } = withRules(c);
    const looks = conditionalLooks(wb, si);
    const a = area(c.read);
    const mine = {};
    for (let r = a.r0; r <= a.r1; r++) {
      for (let col = a.c0; col <= a.c1; col++) {
        const l = looks.get(`${r},${col}`);
        if (!l) continue;
        const one = {};
        if (l.fill) one.fill = l.fill;
        if (l.color) one.color = l.color;
        if (l.bold) one.bold = true;
        if (Object.keys(one).length) mine[String.fromCharCode(65 + col) + (r + 1)] = one;
      }
    }
    const theirs = Object.fromEntries(Object.entries(c.excel.cells).map(([k, v]) => [k, { ...(v.fill != null ? { fill: hex(v.fill) } : {}), ...(v.color != null ? { color: hex(v.color) } : {}), ...(v.bold ? { bold: true } : {}) }]));
    assert.deepEqual(mine, theirs);
  });
}
