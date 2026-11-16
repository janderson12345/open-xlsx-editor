// Writes src/engine/tablestyles.js from Excel's measured built-in table styles (test/tablestyles-excel.json, from
// excel-tablestyles.ps1): each style's elements as the differential formats Excel saved for its duplicate, the dxfs
// shared between styles listed once.
// usage: node tools/excel-compare/tablestyles-data.mjs
import { readFileSync, writeFileSync } from 'node:fs';

const measured = JSON.parse(readFileSync('test/tablestyles-excel.json', 'utf8'));
const dxfs = [];
const at = new Map();
const styles = {};
for (const [name, s] of Object.entries(measured.styles)) {
  const els = {};
  for (const [type, xml] of Object.entries(s.dxfs)) {
    const inner = xml.replace(/^<dxf>|<\/dxf>$/g, '');
    if (!at.has(inner)) { at.set(inner, dxfs.length); dxfs.push(inner); }
    els[type] = at.get(inner);
  }
  // (Every built-in style's stripes are one row or column: TableStyleElement.StripeSize over COM.)
  styles[name] = els;
}
const out = [
  "// Excel's built-in table styles (TableStyleLight1..21, Medium1..28, Dark1..11), as measured: each style duplicated in",
  '// Excel (TableStyle.Duplicate over COM, tools/excel-compare/excel-tablestyles.ps1) and the differential formats it',
  '// saved read back (test/tablestyles-excel.json); every stripe of them one row or column wide. Written by',
  '// tools/excel-compare/tablestyles-data.mjs: not edited by hand.',
  '',
  "/** The dxfs' contents (a <dxf>'s children), each once. */",
  'export const TABLE_DXFS = [',
  ...dxfs.map((d) => `  ${JSON.stringify(d)},`),
  '];',
  '',
  '/** Each built-in style: its elements (wholeTable, headerRow, ...) as indexes into TABLE_DXFS. */',
  'export const TABLE_STYLES = {',
  ...Object.entries(styles).map(([n, els]) => `  ${n}: ${JSON.stringify(els).replace(/"(\w+)":/g, '$1: ').replace(/,/g, ', ')},`),
  '};',
  '',
];
writeFileSync('src/engine/tablestyles.js', out.join('\n'));
console.log(`${Object.keys(styles).length} styles, ${dxfs.length} dxfs`);
