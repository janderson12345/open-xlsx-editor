// Generates the test workbooks in this folder from code, so the repository holds no third-party files.
// usage: node test/fixtures/make-fixtures.mjs   (npm run fixtures)
//
// basic.xlsx     three sheets (one hidden), shared strings with rich text, a phonetic run and an _x000D_ escape,
//                numbers, booleans, an error, formulas (one shared) with cached values and a calcChain, an inline
//                string, column widths, a hidden column and row, a tall row, a merged range, long text, and parts and
//                elements the reader doesn't model (conditional format, validation, extLst, docProps, customXml).
// prefixed.xlsx  the same vocabulary under an `x:` prefix, with rows and cells that leave `r` implied, inline strings
//                only (no shared strings part).
// large.xlsx     5,000 rows by 20 columns, for the virtualised grid.
//
// The zips are written with a fixed date, so the files only change when this script does.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { zipSync, strToU8 } from 'fflate';

const here = path.dirname(fileURLToPath(import.meta.url));
const X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const RT = (t) => `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${t}`;
const CT = (t) => `application/vnd.openxmlformats-officedocument.spreadsheetml.${t}+xml`;
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const col = (c) => { let s = ''; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };

const rels = (list) => `${DECL}<Relationships xmlns="${PR}">${list.map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`).join('')}</Relationships>`;

function contentTypes(overrides) {
  return `${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    overrides.map(([part, type]) => `<Override PartName="/${part}" ContentType="${type}"/>`).join('') + '</Types>';
}

const STYLES = `${DECL}<styleSheet xmlns="${X}">` +
  '<fonts count="2"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>' +
  '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
  '<dxfs count="1"><dxf><font><b/></font></dxf></dxfs><tableStyles count="0"/></styleSheet>';

const docProps = () => ({
  'docProps/core.xml': `${DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>open-xlsx-editor fixtures</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created></cp:coreProperties>`,
  'docProps/app.xml': `${DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>open-xlsx-editor fixtures</Application></Properties>`,
});
const PROPS_CT = [['docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'], ['docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml']];
const ROOT_RELS = rels([['rId1', RT('officeDocument'), 'xl/workbook.xml'], ['rId2', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml'], ['rId3', RT('extended-properties'), 'docProps/app.xml']]);

function write(name, parts) {
  const files = {};
  for (const [k, v] of Object.entries(parts)) files[k] = [strToU8(v), { mtime: new Date('2026-01-01T00:00:00') }];
  fs.writeFileSync(path.join(here, name), zipSync(files, { level: 6 }));
  console.log(`wrote test/fixtures/${name}`);
}

// Shared strings collected while writing cells.
function stringTable() {
  const list = [];
  const index = new Map();
  return {
    add(s) { if (!index.has(s)) { index.set(s, list.length); list.push(`<si><t${/^\s|\s$/.test(s) ? ' xml:space="preserve"' : ''}>${esc(s)}</t></si>`); } return index.get(s); },
    raw(xml) { list.push(xml); return list.length - 1; },
    xml: (count) => `${DECL}<sst xmlns="${X}" count="${count}" uniqueCount="${list.length}">${list.join('')}</sst>`,
  };
}

function basic() {
  const sst = stringTable();
  let uses = 0;
  const s = (ref, text, style) => { uses++; return `<c r="${ref}" t="s"${style ? ` s="${style}"` : ''}><v>${sst.add(text)}</v></c>`; };
  const n = (ref, v) => `<c r="${ref}"><v>${v}</v></c>`;
  const rows = [];
  rows.push(`<row r="1" spans="1:6">${['Item', 'Amount', 'Double', 'Hidden col', 'Choice', 'Wide column'].map((h, i) => s(`${col(i)}1`, h, 1)).join('')}</row>`);
  for (let r = 2; r <= 10; r++) {
    const attrs = r === 4 ? ' ht="30" customHeight="1"' : r === 6 ? ' hidden="1"' : '';
    const f = r === 2 ? `<f t="shared" ref="C2:C10" si="0">B2*2</f>` : '<f t="shared" si="0"/>';
    rows.push(`<row r="${r}" spans="1:6"${attrs}>${s(`A${r}`, `Item ${r - 1}`)}${n(`B${r}`, r - 1)}<c r="C${r}">${f}<v>${(r - 1) * 2}</v></c>${n(`D${r}`, (r - 1) / 8)}${s(`E${r}`, r % 2 ? 'Yes' : 'No')}</row>`);
  }
  rows.push(`<row r="11" spans="1:6">${s('A11', 'Total', 1)}<c r="B11" s="1"><f>SUM(B2:B10)</f><v>45</v></c><c r="C11" t="b"><v>1</v></c><c r="D11" t="e"><f>1/0</f><v>#DIV/0!</v></c><c r="E11" t="inlineStr"><is><t>inline text</t></is></c><c r="F11"><v>1234567.891</v></c></row>`);
  // Rich text with a phonetic run (not part of the cell's text), and an escaped carriage return.
  uses++;
  const rich = sst.raw('<si><r><t>Rich </t></r><r><rPr><b/><sz val="11"/><rFont val="Calibri"/></rPr><t>bold</t></r><rPh sb="0" eb="1"><t>ruby</t></rPh><phoneticPr fontId="0"/></si>');
  uses++;
  const cr = sst.raw('<si><t>Line_x000D_end</t></si>');
  rows.push(`<row r="12" spans="1:6">${s('A12', 'Merged across A12:C13')}<c r="E12" t="s"><v>${rich}</v></c><c r="F12" t="s"><v>${cr}</v></c></row>`);
  rows.push('<row r="13" spans="1:6"><c r="F13"><v>0.000012345</v></c></row>');
  rows.push(`<row r="15">${s('A15', 'This long text runs on across the empty cells to its right, as text does in a sheet.')}</row>`);
  rows.push(`<row r="40">${n('A40', 40)}${n('H40', 123456789012)}</row>`);
  const sheet1 = `${DECL}<worksheet xmlns="${X}" xmlns:r="${R}">` +
    '<dimension ref="A1:H40"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>' +
    '<sheetFormatPr defaultRowHeight="15"/>' +
    '<cols><col min="2" max="2" width="20.7109375" customWidth="1"/><col min="4" max="4" width="9.140625" hidden="1" customWidth="1"/><col min="6" max="6" width="30.7109375" customWidth="1"/></cols>' +
    `<sheetData>${rows.join('')}</sheetData>` +
    '<mergeCells count="1"><mergeCell ref="A12:C13"/></mergeCells>' +
    '<conditionalFormatting sqref="B2:B10"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>' +
    '<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="E2:E10"><formula1>"Yes,No"</formula1></dataValidation></dataValidations>' +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
    '<extLst><ext uri="{6F2D7B1E-0000-4000-8000-0000000000AA}" xmlns:oxe="urn:open-xlsx-editor:test"><oxe:note>kept as it came</oxe:note></ext></extLst></worksheet>';
  const sheet2 = `${DECL}<worksheet xmlns="${X}" xmlns:r="${R}"><dimension ref="A1:B2"/><sheetData>` +
    `<row r="1">${s('A1', 'Notes')}</row><row r="2">${s('A2', '  leading spaces kept')}${n('B2', -3.5)}</row></sheetData>` +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
  const sheet3 = `${DECL}<worksheet xmlns="${X}"><sheetData><row r="1">${s('A1', 'Hidden sheet')}</row></sheetData></worksheet>`;
  const calcChain = `${DECL}<calcChain xmlns="${X}">${[2, 3, 4, 5, 6, 7, 8, 9, 10].map((r, i) => `<c r="C${r}"${i ? '' : ' i="1"'}/>`).join('')}<c r="B11"/><c r="D11"/></calcChain>`;
  const workbook = `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><workbookPr defaultThemeVersion="164011"/>` +
    '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="20000" windowHeight="10000"/></bookViews>' +
    '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Notes &amp; more" sheetId="2" r:id="rId2"/><sheet name="Hidden" sheetId="3" state="hidden" r:id="rId3"/></sheets>' +
    '<definedNames><definedName name="Amounts">Data!$B$2:$B$10</definedName></definedNames><calcPr calcId="191029"/></workbook>';
  write('basic.xlsx', {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/worksheets/sheet2.xml', CT('worksheet')], ['xl/worksheets/sheet3.xml', CT('worksheet')], ['xl/sharedStrings.xml', CT('sharedStrings')], ['xl/styles.xml', CT('styles')], ['xl/calcChain.xml', CT('calcChain')], ...PROPS_CT]),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': workbook,
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), 'worksheets/sheet1.xml'], ['rId2', RT('worksheet'), 'worksheets/sheet2.xml'], ['rId3', RT('worksheet'), 'worksheets/sheet3.xml'], ['rId4', RT('styles'), 'styles.xml'], ['rId5', RT('sharedStrings'), 'sharedStrings.xml'], ['rId6', RT('calcChain'), 'calcChain.xml'], ['rId7', RT('customXml'), '../customXml/item1.xml']]),
    'xl/worksheets/sheet1.xml': sheet1,
    'xl/worksheets/sheet2.xml': sheet2,
    'xl/worksheets/sheet3.xml': sheet3,
    'xl/sharedStrings.xml': sst.xml(uses),
    'xl/styles.xml': STYLES,
    'xl/calcChain.xml': calcChain,
    'customXml/item1.xml': `${DECL}<oxe:data xmlns:oxe="urn:open-xlsx-editor:test">a part nobody models</oxe:data>`,
  });
}

function prefixed() {
  const is = (t) => `<x:c t="inlineStr"><x:is><x:t>${esc(t)}</x:t></x:is></x:c>`;
  // Rows 1-3 leave `r` implied; row 5 gives it, row 6 follows on. Cells leave it implied except C5.
  const data = `<x:row>${is('First')}${is('Second')}</x:row><x:row><x:c><x:v>1</x:v></x:c><x:c><x:v>2</x:v></x:c></x:row><x:row/>` +
    `<x:row r="5"><x:c><x:v>5</x:v></x:c><x:c r="C5"><x:v>3</x:v></x:c><x:c><x:v>4</x:v></x:c></x:row><x:row>${is('Row six')}</x:row>`;
  write('prefixed.xlsx', {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/styles.xml', CT('styles')], ...PROPS_CT]),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': `${DECL}<x:workbook xmlns:x="${X}" xmlns:r="${R}"><x:sheets><x:sheet name="Prefixed" sheetId="1" r:id="rId1"/></x:sheets></x:workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), '/xl/worksheets/sheet1.xml'], ['rId2', RT('styles'), 'styles.xml']]),
    'xl/worksheets/sheet1.xml': `${DECL}<x:worksheet xmlns:x="${X}"><x:sheetData>${data}</x:sheetData></x:worksheet>`,
    'xl/styles.xml': STYLES,
  });
}

function large() {
  const sst = stringTable();
  const ROWS = 5000;
  const COLS = 20;
  const rows = [];
  rows.push(`<row r="1">${Array.from({ length: COLS }, (_, c) => `<c r="${col(c)}1" t="s" s="1"><v>${sst.add(`Column ${col(c)}`)}</v></c>`).join('')}</row>`);
  for (let r = 2; r <= ROWS; r++) {
    let cells = `<c r="A${r}" t="s"><v>${sst.add(`Row ${r}`)}</v></c>`;
    for (let c = 1; c < COLS; c++) cells += `<c r="${col(c)}${r}"><v>${r * 100 + c}</v></c>`;
    rows.push(`<row r="${r}">${cells}</row>`);
  }
  write('large.xlsx', {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/sharedStrings.xml', CT('sharedStrings')], ['xl/styles.xml', CT('styles')], ...PROPS_CT]),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><sheets><sheet name="Large" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), 'worksheets/sheet1.xml'], ['rId2', RT('styles'), 'styles.xml'], ['rId3', RT('sharedStrings'), 'sharedStrings.xml']]),
    'xl/worksheets/sheet1.xml': `${DECL}<worksheet xmlns="${X}" xmlns:r="${R}"><dimension ref="A1:${col(COLS - 1)}${ROWS}"/><sheetData>${rows.join('')}</sheetData></worksheet>`,
    'xl/sharedStrings.xml': sst.xml(ROWS + COLS - 1),
    'xl/styles.xml': STYLES,
  });
}

// A probe for measuring Excel's printed geometry: 40 columns of different widths, each with left-aligned text in
// row 1 (where it starts gives the column's left edge) and a right-aligned number in row 2 (its right edge), and rows
// of different heights with text in column A (baselines step by the row heights).
export const PROBE_CHARS = [0.5, 1, 1.5, 2, 2.5, 3, 4, 5, 6, 7, 8, 8.43, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 20, 22, 25, 28, 30, 33, 36, 40, 1.29, 2.71, 4.57, 10.14, 12.71, 19.29, 26.86, 3.86, 6.43, 7.71];
export const PROBE_HEIGHTS = [8, 9, 10, 11, 12, 12.75, 13.5, 14.25, 15, 15.75, 16.5, 18, 20, 22.5, 24, 27, 30, 36, 45, 60];
function widths() {
  const stored = (chars) => Math.trunc(((chars * 7 + 5) / 7) * 256) / 256;
  const cols = PROBE_CHARS.map((ch, i) => `<col min="${i + 1}" max="${i + 1}" width="${stored(ch)}" customWidth="1"/>`).join('');
  const is = (ref, t) => `<c r="${ref}" t="inlineStr"><is><t>${esc(t)}</t></is></c>`;
  const rows = [`<row r="1">${PROBE_CHARS.map((_, i) => is(`${col(i)}1`, `c${i + 1}`)).join('')}</row>`,
    `<row r="2">${PROBE_CHARS.map((_, i) => `<c r="${col(i)}2"><v>${i + 1}</v></c>`).join('')}</row>`,
    ...PROBE_HEIGHTS.map((ht, i) => `<row r="${i + 4}" ht="${ht}" customHeight="1">${is(`A${i + 4}`, `h${ht}`)}</row>`)];
  write('widths.xlsx', {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/styles.xml', CT('styles')], ...PROPS_CT]),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><sheets><sheet name="Widths" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), 'worksheets/sheet1.xml'], ['rId2', RT('styles'), 'styles.xml']]),
    'xl/worksheets/sheet1.xml': `${DECL}<worksheet xmlns="${X}"><sheetFormatPr defaultRowHeight="15"/><cols>${cols}</cols><sheetData>${rows.join('')}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`,
    'xl/styles.xml': STYLES,
  });
}

// A probe for Excel's General format: each value in five columns of different widths (with empty columns between,
// so the extracted texts stay apart), to see how many digits Excel shows and when it switches to scientific.
export const GENERAL_VALUES = [0, 1, -1, 0.5, 1 / 3, 2 / 3, -2 / 3, 0.1, 0.1 + 0.2, 0.0001, 0.00001, 0.000015, 0.000012345,
  0.0000123456789, 0.000001, 0.00012345678, 1e-9, 1.5e-10, 1e-15, 1.23456789e-300, 0.123456789012, -0.123456789012, 12.3456789012,
  3.14159265358979, 1234.5, 100000, 123456.7891234, 1234567.891, 123456789, 1000000000, 1234567890, 12345678901,
  -12345678901, 99999999999, 99999999999.5, 123456789012, 1234567890123, 1e15, 1e20, 1e100, 9.99999999999e10, 5e-5, 0.00099999];
export const GENERAL_CHARS = [30, 8.43, 6, 4, 2.5];
function general() {
  const stored = (chars) => Math.trunc(((chars * 7 + 5) / 7) * 256) / 256;
  // Value columns A, C, E, G, I; the empty ones between are 1 character wide.
  const cols = GENERAL_CHARS.flatMap((ch, i) => [`<col min="${2 * i + 1}" max="${2 * i + 1}" width="${stored(ch)}" customWidth="1"/>`,
    `<col min="${2 * i + 2}" max="${2 * i + 2}" width="${stored(1)}" customWidth="1"/>`]).join('');
  const rows = GENERAL_VALUES.map((v, r) => `<row r="${r + 1}">${GENERAL_CHARS.map((_, i) => `<c r="${col(2 * i)}${r + 1}"><v>${v}</v></c>`).join('')}</row>`);
  write('general.xlsx', {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/styles.xml', CT('styles')], ...PROPS_CT]),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><sheets><sheet name="General" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), 'worksheets/sheet1.xml'], ['rId2', RT('styles'), 'styles.xml']]),
    'xl/worksheets/sheet1.xml': `${DECL}<worksheet xmlns="${X}"><sheetFormatPr defaultRowHeight="15"/><cols>${cols}</cols><sheetData>${rows.join('')}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`,
    'xl/styles.xml': STYLES,
  });
}

// A theme part of our own: Calibri for body text and headings, plain colours. (A workbook without one is drawn in
// Excel 365's default theme, Aptos Narrow, which browsers don't have.)
const THEME = `${DECL}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="open-xlsx-editor fixtures"><a:themeElements>` +
  '<a:clrScheme name="fixtures">' + [['dk1', '000000'], ['lt1', 'FFFFFF'], ['dk2', '1F3864'], ['lt2', 'E7E6E6'], ['accent1', '4472C4'], ['accent2', 'ED7D31'],
  ['accent3', 'A5A5A5'], ['accent4', 'FFC000'], ['accent5', '5B9BD5'], ['accent6', '70AD47'], ['hlink', '0563C1'], ['folHlink', '954F72']]
    .map(([n, c]) => `<a:${n}><a:srgbClr val="${c}"/></a:${n}>`).join('') + '</a:clrScheme>' +
  '<a:fontScheme name="fixtures"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
  '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>' +
  '<a:fmtScheme name="fixtures"><a:fillStyleLst>' + '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'.repeat(3) + '</a:fillStyleLst>' +
  '<a:lnStyleLst>' + '<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'.repeat(3) + '</a:lnStyleLst>' +
  '<a:effectStyleLst>' + '<a:effectStyle><a:effectLst/></a:effectStyle>'.repeat(3) + '</a:effectStyleLst>' +
  '<a:bgFillStyleLst>' + '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'.repeat(3) + '</a:bgFillStyleLst></a:fmtScheme>' +
  '</a:themeElements></a:theme>';

// A probe for fonts: one row per font and size (its own cell format), text in A and a number in B, each row tall
// enough for its font (customHeight, so Excel doesn't resize it), with the theme above.
export const PROBE_FONTS = [['Calibri', 11, 'minor'], ['Calibri', 11, 'minor', 'b'], ['Calibri', 11, 'minor', 'i'], ['Calibri', 8, 'minor'],
  ['Calibri', 14, 'minor'], ['Calibri', 20, 'minor'], ['Cambria', 11], ['Cambria', 20, null, 'i'], ['Arial', 11], ['Arial', 14],
  ['Times New Roman', 11], ['Times New Roman', 16], ['Courier New', 11], ['Courier New', 9], ['Calibri', 11]];
function fonts() {
  const font = ([name, sz, scheme, style]) => `<font>${style === 'b' ? '<b/>' : ''}${style === 'i' ? '<i/>' : ''}<sz val="${sz}"/><color theme="1"/><name val="${name}"/>${scheme ? `<scheme val="${scheme}"/>` : ''}</font>`;
  const styles = `${DECL}<styleSheet xmlns="${X}"><fonts count="${PROBE_FONTS.length}">${PROBE_FONTS.map(font).join('')}</fonts>` +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    `<cellXfs count="${PROBE_FONTS.length}">${PROBE_FONTS.map((_, i) => `<xf numFmtId="0" fontId="${i}" fillId="0" borderId="0" xfId="0"${i ? ' applyFont="1"' : ''}/>`).join('')}</cellXfs>` +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0"/></styleSheet>';
  const rows = PROBE_FONTS.map((f, i) => {
    const ht = Math.max(15, Math.ceil(f[1] * 1.6));
    return `<row r="${i + 1}" ht="${ht}" customHeight="1"><c r="A${i + 1}" s="${i}" t="inlineStr"><is><t>${esc(`${f[0]} ${f[1]} ${f[3] || ''} Sample 0123`)}</t></is></c><c r="B${i + 1}" s="${i}"><v>${12345.678 + i}</v></c></row>`;
  });
  // Then, in the Normal style (Calibri 11 here), three rows of the default height and the widths probe's row heights:
  // printed row heights depend on the Normal style's font (docs/LAYOUT.md).
  const at = PROBE_FONTS.length + 2;
  const txt = (r, t) => `<c r="A${r}" t="inlineStr"><is><t>${esc(t)}</t></is></c>`;
  for (let k = 0; k < 3; k++) rows.push(`<row r="${at + k}">${txt(at + k, `d${k + 1}`)}</row>`);
  PROBE_HEIGHTS.forEach((ht, k) => rows.push(`<row r="${at + 3 + k}" ht="${ht}" customHeight="1">${txt(at + 3 + k, `h${ht}`)}</row>`));
  write('fonts.xlsx', {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/styles.xml', CT('styles')], ['xl/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'], ...PROPS_CT]),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><sheets><sheet name="Fonts" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), 'worksheets/sheet1.xml'], ['rId2', RT('styles'), 'styles.xml'], ['rId3', RT('theme'), 'theme/theme1.xml']]),
    'xl/worksheets/sheet1.xml': `${DECL}<worksheet xmlns="${X}"><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="45.7109375" customWidth="1"/><col min="2" max="2" width="20.7109375" customWidth="1"/></cols><sheetData>${rows.join('')}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`,
    'xl/styles.xml': styles,
    'xl/theme/theme1.xml': THEME,
  });
}

// A probe for fills and borders: each sample is one cell with empty cells around it (every other row and column), with
// the theme above. Fills in columns B and D (RGB, theme colours 0-9, tints, indexed colours, a pattern), boxes in each
// border style in F and H, and the sample's number in J.
export const PROBE_FILLS = [['rgb', 'FFFF0000'], ['rgb', 'FF00B050'], ...[0, 1, 2, 3, 4, 5, 6, 7, 8, 9].map((t) => ['theme', t, 0]),
  ['theme', 4, -0.5], ['theme', 4, -0.25], ['theme', 4, 0.4], ['theme', 4, 0.6], ['theme', 4, 0.8], ['theme', 1, 0.5],
  ['theme', 0, -0.15], ['theme', 9, -0.249977111117893], ['indexed', 10], ['indexed', 13], ['indexed', 22], ['indexed', 30],
  ['pattern', 'darkGray', 'FF4472C4', 'FFFFFF00']];
export const PROBE_BORDERS = [['thin'], ['medium'], ['thick'], ['hair'], ['dotted'], ['dashed'], ['double'], ['mediumDashed'],
  ['dashDot'], ['thin', 'rgb', 'FFFF0000'], ['medium', 'theme', 5], ['thick', 'indexed', 12]];
function fills() {
  const color = (tag, kind, a, b) => (kind === 'rgb' ? `<${tag} rgb="${a}"/>` : kind === 'theme' ? `<${tag} theme="${a}"${b ? ` tint="${b}"` : ''}/>` : kind === 'indexed' ? `<${tag} indexed="${a}"/>` : `<${tag} auto="1"/>`);
  const fillXml = PROBE_FILLS.map(([kind, a, b, c]) => (kind === 'pattern'
    ? `<fill><patternFill patternType="${a}">${color('fgColor', 'rgb', b)}${color('bgColor', 'rgb', c)}</patternFill></fill>`
    : `<fill><patternFill patternType="solid">${color('fgColor', kind, a, b)}<bgColor indexed="64"/></patternFill></fill>`));
  const side = (name, [style, kind, a]) => `<${name} style="${style}">${kind ? color('color', kind, a) : '<color indexed="64"/>'}</${name}>`;
  const borderXml = PROBE_BORDERS.map((b) => `<border>${side('left', b)}${side('right', b)}${side('top', b)}${side('bottom', b)}<diagonal/></border>`);
  const nf = PROBE_FILLS.length;
  const styles = `${DECL}<styleSheet xmlns="${X}"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>` +
    `<fills count="${nf + 2}"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>${fillXml.join('')}</fills>` +
    `<borders count="${PROBE_BORDERS.length + 1}"><border><left/><right/><top/><bottom/><diagonal/></border>${borderXml.join('')}</borders>` +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    `<cellXfs count="${1 + nf + PROBE_BORDERS.length}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
    PROBE_FILLS.map((_, i) => `<xf numFmtId="0" fontId="0" fillId="${i + 2}" borderId="0" xfId="0" applyFill="1"/>`).join('') +
    PROBE_BORDERS.map((_, i) => `<xf numFmtId="0" fontId="0" fillId="0" borderId="${i + 1}" xfId="0" applyBorder="1"/>`).join('') + '</cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0"/></styleSheet>';
  const rows = new Map();
  const put = (r, c, xml) => rows.set(r, (rows.get(r) || []).concat([[c, xml]]));
  PROBE_FILLS.forEach((_, i) => {
    const r = 2 + 2 * Math.floor(i / 2);
    put(r, i % 2 ? 3 : 1, `<c r="${col(i % 2 ? 3 : 1)}${r}" s="${1 + i}"/>`);
  });
  PROBE_BORDERS.forEach((_, i) => {
    const r = 2 + 2 * Math.floor(i / 2);
    put(r, i % 2 ? 7 : 5, `<c r="${col(i % 2 ? 7 : 5)}${r}" s="${1 + nf + i}"/>`);
  });
  for (const r of rows.keys()) put(r, 9, `<c r="J${r}"><v>${r}</v></c>`);
  const data = [...rows].sort((a, b) => a[0] - b[0]).map(([r, cells]) => `<row r="${r}">${cells.sort((a, b) => a[0] - b[0]).map((x) => x[1]).join('')}</row>`);
  write('fills.xlsx', {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/styles.xml', CT('styles')], ['xl/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'], ...PROPS_CT]),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><sheets><sheet name="Fills" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), 'worksheets/sheet1.xml'], ['rId2', RT('styles'), 'styles.xml'], ['rId3', RT('theme'), 'theme/theme1.xml']]),
    'xl/worksheets/sheet1.xml': `${DECL}<worksheet xmlns="${X}"><sheetFormatPr defaultRowHeight="15"/><sheetData>${data.join('')}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`,
    'xl/styles.xml': styles,
    'xl/theme/theme1.xml': THEME,
  });
}

// A probe for alignment: one sample per row (45pt tall, wrapped ones 75pt), in column B (20 characters wide) or, for
// text running on, in D with empty neighbours; with our theme. Each sample is [text or number, alignment attributes].
const WRAP_TEXT = 'Wrapped words flow onto several lines within one cell when they do not fit its width';
export const PROBE_ALIGN = [
  ['L bottom', { horizontal: 'left' }], ['C bottom', { horizontal: 'center' }], ['R bottom', { horizontal: 'right' }],
  ['G text', {}], [12345, {}], [12346, { horizontal: 'left' }], [12347, { horizontal: 'center' }],
  ['L top', { horizontal: 'left', vertical: 'top' }], ['L middle', { horizontal: 'left', vertical: 'center' }],
  ['C middle', { horizontal: 'center', vertical: 'center' }], ['R top', { horizontal: 'right', vertical: 'top' }],
  ['Indent one', { horizontal: 'left', indent: 1 }], ['Indent two', { horizontal: 'left', indent: 2 }],
  ['Indent R three', { horizontal: 'right', indent: 3 }],
  [WRAP_TEXT, { wrapText: 1, vertical: 'top' }], [WRAP_TEXT, { wrapText: 1, horizontal: 'center' }],
  [WRAP_TEXT, { wrapText: 1, vertical: 'center' }], [WRAP_TEXT, { wrapText: 1, horizontal: 'right', vertical: 'bottom' }],
  ['Centred text running both ways', { horizontal: 'center' }, 'D'], ['Right text running to the left', { horizontal: 'right' }, 'D'],
];
function align() {
  const attrs = (a) => Object.entries(a).map(([k, v]) => ` ${k}="${v}"`).join('');
  const styles = `${DECL}<styleSheet xmlns="${X}"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>` +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    `<cellXfs count="${PROBE_ALIGN.length + 1}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
    PROBE_ALIGN.map(([, a]) => `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment${attrs(a)}/></xf>`).join('') + '</cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0"/></styleSheet>';
  const rows = PROBE_ALIGN.map(([v, a, where = 'B'], i) => {
    const r = 2 + 2 * i;
    const ht = a.wrapText ? 75 : 45;
    const cell = typeof v === 'number' ? `<c r="${where}${r}" s="${i + 1}"><v>${v}</v></c>` : `<c r="${where}${r}" s="${i + 1}" t="inlineStr"><is><t>${esc(v)}</t></is></c>`;
    return `<row r="${r}" ht="${ht}" customHeight="1">${cell}</row>`;
  });
  write('align.xlsx', {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/styles.xml', CT('styles')], ['xl/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'], ...PROPS_CT]),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><sheets><sheet name="Align" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), 'worksheets/sheet1.xml'], ['rId2', RT('styles'), 'styles.xml'], ['rId3', RT('theme'), 'theme/theme1.xml']]),
    'xl/worksheets/sheet1.xml': `${DECL}<worksheet xmlns="${X}"><sheetFormatPr defaultRowHeight="15"/><cols><col min="2" max="2" width="20.7109375" customWidth="1"/></cols><sheetData>${rows.join('')}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`,
    'xl/styles.xml': styles,
    'xl/theme/theme1.xml': THEME,
  });
}

// A probe for number formats: every format with every value, one pair per row: the pair's number as text in A
// ("f3v7") and the value in B (40 characters wide) in that format. Built-in formats by id, the rest as custom codes.
export const PROBE_NUMFMT_BUILTIN = [1, 2, 3, 4, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 37, 38, 39, 40, 45, 46, 47, 48, 49];
export const PROBE_NUMFMT_CUSTOM = ['0.000', '#,##0.0,,"M"', '$#,##0.00', '[Red]-0.00;[Blue]0.00', '0.00;(0.00);"zero";"text: "@',
  '[>=100]"big";[<0]"neg";"small"', 'yyyy-mm-dd', 'dddd, mmmm d, yyyy', 'mmm d', 'hh:mm:ss.000', '[mm]:ss', 'h "h" mm "m"',
  '000-00-0000', '# ?/8', '0.0%', '"Total: "0', '\\$0', '0_);(0)', '#,##0_);[Red](#,##0)', '#', '#.##', '0.#', '???.???',
  '"x"General', ';;;', 'm/d/yy h:mm AM/PM', '[$-409]mmmm yyyy', '0.00E+00', '##0.00E+0', 'd'];
export const PROBE_NUMFMT_VALUES = [0, 1, -1, 0.5, 1234.5678, -1234.5678, 0.001234, 1234567.891, 45678.75, 0.75, 2.5, 60, 61, 'text'];
function numfmts() {
  const custom = PROBE_NUMFMT_CUSTOM.map((code, i) => `<numFmt numFmtId="${164 + i}" formatCode="${esc(code)}"/>`).join('');
  const ids = [...PROBE_NUMFMT_BUILTIN, ...PROBE_NUMFMT_CUSTOM.map((_, i) => 164 + i)];
  const styles = `${DECL}<styleSheet xmlns="${X}"><numFmts count="${PROBE_NUMFMT_CUSTOM.length}">${custom}</numFmts>` +
    '<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>' +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    `<cellXfs count="${ids.length + 1}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
    ids.map((id) => `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`).join('') + '</cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0"/></styleSheet>';
  const rows = [];
  ids.forEach((_, f) => PROBE_NUMFMT_VALUES.forEach((v, k) => {
    const r = rows.length + 1;
    const cell = typeof v === 'number' ? `<c r="B${r}" s="${f + 1}"><v>${v}</v></c>` : `<c r="B${r}" s="${f + 1}" t="inlineStr"><is><t>${v}</t></is></c>`;
    rows.push(`<row r="${r}"><c r="A${r}" t="inlineStr"><is><t>f${f}v${k}</t></is></c>${cell}</row>`);
  }));
  write('numfmt.xlsx', {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/styles.xml', CT('styles')], ['xl/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'], ...PROPS_CT]),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><sheets><sheet name="Formats" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), 'worksheets/sheet1.xml'], ['rId2', RT('styles'), 'styles.xml'], ['rId3', RT('theme'), 'theme/theme1.xml']]),
    'xl/worksheets/sheet1.xml': `${DECL}<worksheet xmlns="${X}"><sheetFormatPr defaultRowHeight="15"/><cols><col min="2" max="2" width="40.7109375" customWidth="1"/></cols><sheetData>${rows.join('')}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`,
    'xl/styles.xml': styles,
    'xl/theme/theme1.xml': THEME,
  });
}

// A probe for row heights Excel works out itself: rows without customHeight, sized by their fonts and wrapped text;
// two rows whose stored ht disagrees with their content. Column A of every row is filled (two colours in turn), so
// the fills' edges in Excel's PDF give each row's printed height exactly. Each sample: [label, cells, row attributes];
// a cell is [column, value, font index, wrap].
export const PROBE_ROW_FONTS = [['Calibri', 11, 'minor'], ['Calibri', 8, 'minor'], ['Calibri', 14, 'minor'], ['Calibri', 20, 'minor'],
  ['Calibri', 36, 'minor'], ['Cambria', 20], ['Times New Roman', 16], ['Arial', 14], ['Courier New', 11]];
const LONG = 'Several words of text that wrap inside a cell of twenty characters and go on';
export const PROBE_ROWS = [
  ['empty', [], {}], ['Calibri 8', [['B', 'Small', 1]], {}], ['Calibri 11', [['B', 'Normal', 0]], {}], ['Calibri 14', [['B', 'Bigger', 2]], {}],
  ['Calibri 20', [['B', 'Large', 3]], {}], ['Calibri 36', [['B', 'Huge', 4]], {}], ['Cambria 20', [['B', 'Serif', 5]], {}],
  ['Times 16', [['B', 'Times', 6]], {}], ['Arial 14', [['B', 'Arial', 7]], {}], ['Courier 11', [['B', 'Mono', 8]], {}],
  ['mixed 11 and 20', [['B', 'Small one', 0], ['C', 'Large one', 3]], {}], ['styled, empty', [['B', null, 3]], {}],
  ['wrap 2 lines', [['B', LONG.slice(0, 32), 0, true]], {}], ['wrap 3 lines', [['B', LONG.slice(0, 52), 0, true]], {}],
  ['wrap 5 lines', [['B', LONG + ' and on and on further', 0, true]], {}], ['wrap Calibri 14', [['B', LONG.slice(0, 40), 2, true]], {}],
  ['ht 30, not custom', [['B', 'Normal', 0]], { ht: 30 }], ['ht 10, not custom', [['B', 'Large', 3]], { ht: 10 }],
  ['wrap line breaks', [['B', 'one\ntwo\nthree', 0, true]], {}], ['number 20', [['B', 12345, 3]], {}],
];
function rowHeights() {
  const font = ([name, sz, scheme]) => `<font><sz val="${sz}"/><color theme="1"/><name val="${name}"/>${scheme ? `<scheme val="${scheme}"/>` : ''}</font>`;
  // cellXfs: 0 normal, 1 and 2 the two fills, then per font: plain (3 + 2k) and wrapped (4 + 2k).
  const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>', '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>',
    '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>',
    ...PROBE_ROW_FONTS.flatMap((_, k) => [`<xf numFmtId="0" fontId="${k}" fillId="0" borderId="0" xfId="0" applyFont="1"/>`,
      `<xf numFmtId="0" fontId="${k}" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment wrapText="1"/></xf>`])];
  const styles = `${DECL}<styleSheet xmlns="${X}"><fonts count="${PROBE_ROW_FONTS.length}">${PROBE_ROW_FONTS.map(font).join('')}</fonts>` +
    '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/><bgColor indexed="64"/></patternFill></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    `<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs>` +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0"/></styleSheet>';
  const rows = PROBE_ROWS.map(([label, cells, attrs], i) => {
    const r = i + 1;
    const cellXml = cells.map(([c, v, f, wrap]) => {
      const s = 3 + 2 * f + (wrap ? 1 : 0);
      if (v == null) return `<c r="${c}${r}" s="${s}"/>`;
      if (typeof v === 'number') return `<c r="${c}${r}" s="${s}"><v>${v}</v></c>`;
      return `<c r="${c}${r}" s="${s}" t="inlineStr"><is><t${/\n/.test(v) ? ' xml:space="preserve"' : ''}>${esc(v)}</t></is></c>`;
    }).join('');
    const ht = attrs.ht ? ` ht="${attrs.ht}"` : '';
    return `<row r="${r}"${ht}><c r="A${r}" s="${1 + (i % 2)}" t="inlineStr"><is><t>${esc(label)}</t></is></c>${cellXml}</row>`;
  });
  write('rows.xlsx', {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/styles.xml', CT('styles')], ['xl/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'], ...PROPS_CT]),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><sheets><sheet name="Rows" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), 'worksheets/sheet1.xml'], ['rId2', RT('styles'), 'styles.xml'], ['rId3', RT('theme'), 'theme/theme1.xml']]),
    'xl/worksheets/sheet1.xml': `${DECL}<worksheet xmlns="${X}"><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="18.7109375" customWidth="1"/><col min="2" max="2" width="20.7109375" customWidth="1"/></cols><sheetData>${rows.join('')}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`,
    'xl/styles.xml': styles,
    'xl/theme/theme1.xml': THEME,
  });
}

if (process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1])) {
  rowHeights();
  numfmts();
  align();
  fills();
  fonts();
  general();
  basic();
  prefixed();
  large();
  widths();
}
