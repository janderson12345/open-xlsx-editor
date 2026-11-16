// Writes src/engine/cellstyles.js from Excel's measured built-in cell styles (test/cellstyles-excel.json, from
// excel-cellstyles.ps1): for each, the cellStyle's builtinId, the style's xf (what it applies), its number format,
// and the font, fill and border Excel wrote for it in a new workbook, the body font's name written as MINOR and the
// heading font's as MAJOR (the workbook's own are put in when a style is added).
// usage: node tools/excel-compare/cellstyles-data.mjs
import { readFileSync, writeFileSync } from 'node:fs';

const m = JSON.parse(readFileSync('test/cellstyles-excel.json', 'utf8'));
const items = (xml, tag) => [...(xml || '').replace(new RegExp(`^<${tag}s[^>]*>`), '').matchAll(new RegExp(`<${tag}(?: [^>]*)?(?:/>|>[\\s\\S]*?</${tag}>)`, 'g'))].map((x) => x[0]);
const out = [];
for (const st of m.styles) {
  const s = m.saved[st.name];
  // (Every character a regular expression gives meaning to escaped, backslashes too: CodeQL js/incomplete-sanitization.)
  const esc = st.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const cs = new RegExp(`<cellStyle name="${esc}" xfId="(\\d+)" builtinId="(\\d+)"/>`).exec(s.cellStyles);
  const xfId = Number(cs[1]);
  const xf = items(s.cellStyleXfs, 'xf')[xfId];
  const id = (k) => Number(new RegExp(`${k}="(\\d+)"`).exec(xf)?.[1] || 0);
  const flags = Object.fromEntries([...xf.matchAll(/apply(\w+)="0"/g)].map((x) => [x[1], 0]));
  const font = items(s.fonts, 'font')[id('fontId')]
    .replace(/<name val="Aptos Narrow"\/><family val="2"\/><scheme val="minor"\/>/, 'MINOR')
    .replace(/<name val="Aptos Display"\/><family val="2"\/><scheme val="major"\/>/, 'MAJOR');
  const fill = id('fillId') > 1 ? items(s.fills, 'fill')[id('fillId')] : null;
  const border = id('borderId') > 0 ? items(s.borders, 'border')[id('borderId')] : null;
  const numFmtId = id('numFmtId');
  const code = new RegExp(`<numFmt numFmtId="${numFmtId}" formatCode="([^"]*)"/>`).exec(s.numFmts || '')?.[1] ?? null;
  out.push({ name: st.name, builtinId: Number(cs[2]), off: Object.keys(flags), numFmtId, ...(code ? { code: code.replace(/&quot;/g, '"') } : {}), font, ...(fill ? { fill } : {}), ...(border ? { border } : {}) });
}
const lines = [
  "// Excel's 47 built-in cell styles (Home > Cell Styles), as measured: each applied in a new workbook over COM",
  '// (Range.Style, tools/excel-compare/excel-cellstyles.ps1) and what Excel saved read back (test/cellstyles-excel.json):',
  "// its builtinId, the parts its xf doesn't apply (`off`: NumberFormat, Font, Fill, Border, Alignment, Protection), its",
  '// number format and its font, fill and border as Excel writes them (MINOR and MAJOR: the theme\'s body and heading',
  '// font). Written by tools/excel-compare/cellstyles-data.mjs: not edited by hand.',
  '',
  'export const BUILTIN_CELL_STYLES = [',
  ...out.map((x) => `  ${JSON.stringify(x)},`),
  '];',
  '',
];
writeFileSync('src/engine/cellstyles.js', lines.join('\n'));
console.log(`${out.length} styles`);
