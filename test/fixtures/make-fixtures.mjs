// Generates the test workbooks in this folder from code, so the repository holds no third-party files.
// usage: node test/fixtures/make-fixtures.mjs   (npm run fixtures)
//
// basic.xlsx     three sheets (one hidden, one with a tab colour), shared strings with rich text, a phonetic run and
//                an _x000D_ escape, numbers, booleans, an error, formulas (one shared) with cached values and a
//                calcChain, an inline string, column widths, a hidden column and row, a tall row, a merged range, long
//                text, and parts and elements the reader doesn't model (conditional format, validation, extLst,
//                docProps, customXml).
// prefixed.xlsx  the same vocabulary under an `x:` prefix, with rows and cells that leave `r` implied, inline strings
//                only (no shared strings part).
// large.xlsx     5,000 rows by 20 columns, for the virtualised grid.
// charts.xlsx    four charts (clustered columns, lines, a pie, stacked bars) over a sheet of monthly sales, written as
//                Excel writes chart parts.
// notes.xlsx     notes and a threaded comment with a reply, written as Excel writes them (comments, the VML that
//                shows notes, threadedComments, persons), made-up people.
// anchors.xlsx   (and anchors-table.xlsx, with the table Stock and a formula reading it in A20)
//                anchor-cases.mjs's sheet without its table: a stock list on B2:D12, a column chart of Qty
//                anchored F3 to L15, a rectangle M5 to P9 and notes at A1, B6, D12 and F20 (test/anchors.test.js).
//
// The zips are written with a fixed date, so the files only change when this script does.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { zipSync, strToU8 } from 'fflate';
import { normalizeFormula } from '../../src/calc/tokens.js';

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
  const sheet2 = `${DECL}<worksheet xmlns="${X}" xmlns:r="${R}"><sheetPr><tabColor theme="5" tint="0.3999755851924192"/></sheetPr><dimension ref="A1:B2"/><sheetData>` +
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

// A probe for the calculation engine: formulas without saved values, for Excel to work out when it opens the file, on
// a sheet of mixed data (numbers, text, a number as text, TRUE and FALSE, blanks, an error), with defined names, a
// second sheet whose name needs quotes, a shared formula and a chain read in both directions. Column A shows each
// formula as text, column B holds it.
export const PROBE_FORMULAS = [
  // (These first rows sit beside Data rows 2 to 9: a range in a single-value place is cut to the formula's row.)
  'Data!B2:B9', 'Data!B2:B9*1', 'Data!A2:A9&""', 'SUM(Data!B2:B9*1)', 'Data!A1:E1', '-Data!C2:C9', 'ROUND(Data!C2:C9,0)', 'Data!B2:B9',
  // Arithmetic, precedence, coercion
  '1+2*3', '(1+2)*3', '-2^2', '2^3^2', '2^-1', '10/4', '7/0', '0^0', '(-8)^(1/3)', '2^1024', '5%', '50%*4', '-50%', '2*-3',
  '"3"+1', '"1,234"+1', '" 4 "+1', '"$5"+0', '"(5)"+0', '"50%"+0', '"1e3"+0', '"abc"+1', 'TRUE+1', 'FALSE*5', '"2024-01-02"+0',
  '"1/2/2024"+0', '"10:30"+0', '"2:15 PM"+0', '""&1', '1&2', '"a"&TRUE', 'Data!F2+1', 'Data!F2&"x"',
  // Numbers turned into text
  '1/3&""', '2/3&""', '-1/3&""', '0.1+0.2&""', '1E-5&""', '1E-9&""', '1E-10&""', '1.5E-7&""', '1E-15&""', '1E-19&""', '1E-20&""',
  '123456789012345&""', '1234567890123456&""', '1E15&""', '1E20&""', '1E21&""', '1E22&""', '12345678901234567890&""', '100000*3&""',
  '0.000123456789012345678&""', 'LEN(1/7)',
  // Floating point
  '0.1+0.2-0.3', '(0.1+0.2-0.3)', '1*(0.1+0.2-0.3)', '0.1+0.2=0.3', '1/3*3=1', '0.3-0.1-0.2', '1-0.9-0.1', '3*0.1=0.3',
  '1E15+0.3-1E15', '0.1*3', '1.1-1', '(1.1-1)*10',
  // Comparisons
  '1<2', '"a"="A"', '"a"<"b"', '"B">"a"', '1<"a"', '"a"<TRUE', 'FALSE<TRUE', '1=TRUE', 'Data!F2=0', 'Data!F2=""', 'Data!F2=FALSE',
  '"10"<"9"', '1<>1', '"abc"<"abd"', '"a"<"a "', '"Z"<"a"', '"-"<"a"',
  // References and names
  'Data!B2', 'Data!B2+Data!B3', "'My Data'!A1*2", "'My Data'!A2", 'Rate*10', 'SUM(Items)', 'COUNTA(Items)', 'Data!$C$3', 'D11',
  // Aggregates
  'SUM(Data!B2:B9)', 'SUM(Data!B:B)', 'AVERAGE(Data!B2:B9)', 'MIN(Data!B2:B9)', 'MAX(Data!C2:C9)', 'COUNT(Data!A1:E9)',
  'COUNTA(Data!A1:E9)', 'COUNTBLANK(Data!A1:E9)', 'COUNTBLANK(Data!F:F)', 'PRODUCT(Data!B2:B4)', 'SUM(1,"2",TRUE)', 'SUM(Data!B2:B9,10)',
  'COUNT(1,"2",TRUE,"x")', 'COUNT(1,)', 'COUNTA(1,)', 'AVERAGE(1,)', 'SUM(Data!G2:G3)', 'MEDIAN(Data!B2:B9)', 'MEDIAN(1,2,3,4)',
  'LARGE(Data!B2:B9,2)', 'SMALL(Data!C2:C9,3)', 'STDEV(Data!B2:B9)', 'STDEV.P(Data!B2:B9)', 'VAR(Data!C2:C9)', 'VAR.P(Data!C2:C9)',
  'SUMSQ(1,2,3)', 'SUMPRODUCT(Data!B2:B9,Data!C2:C9)', 'SUMPRODUCT((Data!B2:B9>3)*Data!C2:C9)', 'SUMPRODUCT((Data!A2:A9="Apple")*1)',
  'SUMPRODUCT(--ISNUMBER(SEARCH("an",Data!A2:A9)))', 'SUMPRODUCT({1,2,3},{4,5,6})', 'SUM({1,2,3}*2)', 'MAX(Data!E2:E9)',
  'SUM(Data!E2:E9)', 'AVERAGE(Data!A2:A9)',
  // Math
  'ROUND(2.675,2)', 'ROUND(-2.5,0)', 'ROUND(1234.5678,-2)', 'ROUND(0.1+0.2,15)', 'ROUNDUP(3.14159,2)', 'ROUNDUP(-3.14159,2)',
  'ROUNDDOWN(-3.9,0)', 'INT(-3.5)', 'TRUNC(-3.5)', 'TRUNC(3.14159,3)', 'MOD(10,3)', 'MOD(-10,3)', 'MOD(10,-3)', 'MOD(5.5,2)', 'MOD(1,0)',
  'MOD(0.3,0.1)', 'ABS(-4)', 'SQRT(16)', 'SQRT(-1)', 'POWER(2,10)', 'EXP(1)', 'LN(10)', 'LOG(100)', 'LOG(8,2)', 'LOG10(1000)', 'LN(0)',
  'PI()', 'SIGN(-3)', 'CEILING(2.5,1)', 'CEILING(-2.5,-2)', 'CEILING(2.5,-1)', 'CEILING(-2.5,2)', 'FLOOR(2.5,1)', 'FLOOR(-2.5,-2)',
  'FLOOR(-2.5,2)', 'MROUND(10,3)', 'MROUND(-10,3)', 'EVEN(1.5)', 'ODD(1.5)', 'EVEN(-1.5)', 'ODD(0)', 'FACT(5)', 'FACT(0)',
  'QUOTIENT(7,2)', 'QUOTIENT(-7,2)', 'INT(0.1+0.7)*10',
  // Logic
  'IF(1>2,"y","n")', 'IF(Data!F2,"t","f")', 'IF(1,"a")', 'IF(0,"a")', 'IF(TRUE,)', 'IF("x",1,2)', 'IFERROR(1/0,"err")',
  'IFERROR(5,"err")', 'IFNA(NA(),0)', 'IFNA(1/0,0)', 'AND(TRUE,1)', 'OR(FALSE,0)', 'NOT(1)', 'XOR(TRUE,TRUE,TRUE)', 'AND(Data!A1:A3)',
  'AND(Data!E2:E3)', 'OR(Data!E2:E9)', 'IFS(1>2,"a",TRUE,"b")', 'IFS(FALSE,1)', 'SWITCH(2,1,"one",2,"two")', 'SWITCH(3,1,"one","none")',
  // Information
  'ISBLANK(Data!F2)', 'ISBLANK("")', 'ISNUMBER(Data!B2)', 'ISNUMBER(Data!B4)', 'ISNUMBER("1")', 'ISTEXT(Data!A2)', 'ISERROR(1/0)',
  'ISERR(NA())', 'ISNA(NA())', 'ISLOGICAL(TRUE)', 'ISEVEN(4)', 'ISODD(3.7)', 'N(TRUE)', 'N("x")', 'ISNONTEXT(Data!F2)', 'ISREF(Data!A1)',
  'ISERROR(Data!G2)', 'T(1)', 'T("a")',
  // Text
  'CONCATENATE("a",1,TRUE)', 'CONCAT(Data!A2:A4)', 'CONCAT(Data!B2:C3)', 'TEXTJOIN("-",TRUE,Data!A2:A5,"x")',
  'TEXTJOIN(",",FALSE,Data!E4:E6)', 'LEN("hello")', 'LEFT("hello",2)', 'LEFT("hello")', 'RIGHT("hello",3)', 'MID("hello",2,3)',
  'MID("hello",0,1)', 'MID("hello",9,2)', 'UPPER("abc")', 'LOWER("ABC")', 'PROPER("hello wORLD o\'neil 2nd")', 'TRIM("  a   b  ")',
  'REPT("ab",3)', 'SUBSTITUTE("aaa","a","b")', 'SUBSTITUTE("aaa","a","b",2)', 'REPLACE("abcdef",2,3,"X")', 'FIND("l","hello")',
  'FIND("L","hello")', 'SEARCH("L","hello")', 'SEARCH("h?l","ahelp")', 'SEARCH("~?","a?b")', 'EXACT("a","A")', 'VALUE("1,234.5")',
  'VALUE("abc")', 'VALUE("12%")', 'TEXT(1234.567,"#,##0.00")', 'TEXT(0.25,"0%")', 'TEXT(45678,"yyyy-mm-dd")', 'TEXT("abc","0")',
  'TEXT(-1.5,"0.0;(0.0)")', 'TEXT(TRUE,"0")', 'CHAR(65)', 'CODE("A")', '"x"&Data!F2&"y"', 'LEN(Data!B4)', 'UPPER(Data!B2)',
  // Lookup
  'VLOOKUP("Banana",Data!A2:C9,3,FALSE)', 'VLOOKUP("banana",Data!A2:C9,2,FALSE)', 'VLOOKUP("B*",Data!A2:C9,2,FALSE)',
  'VLOOKUP("Zzz",Data!A2:C9,2,FALSE)', 'VLOOKUP(65,Data!I2:J6,2)', 'VLOOKUP(80,Data!I2:J6,2,TRUE)', 'VLOOKUP(-1,Data!I2:J6,2)',
  'VLOOKUP(65,Data!I2:J6,3)', 'VLOOKUP(65,Data!I2:J6,0)', 'VLOOKUP("Fig",Data!A2:E9,5,0)', 'VLOOKUP("Fig",Data!A2:F9,6,0)',
  'HLOOKUP("Price",Data!A1:E9,3,FALSE)', 'MATCH("Cherry",Data!A2:A9,0)', 'MATCH(65,Data!I2:I6,1)', 'MATCH(65,Data!I2:I6)',
  'MATCH("c*",Data!A2:A9,0)', 'MATCH(99,Data!I2:I6,0)', 'INDEX(Data!A2:C9,3,2)', 'INDEX(Data!A2:A9,4)', 'INDEX(Data!A1:E1,3)',
  'INDEX(Data!A2:C9,9,1)', 'SUM(INDEX(Data!B2:C9,0,1))', 'SUM(INDEX(Data!B2:C9,2,0))', 'INDEX({1,2;3,4},2,1)',
  'XLOOKUP("Cherry",Data!A2:A9,Data!C2:C9)', 'XLOOKUP("nope",Data!A2:A9,Data!C2:C9,"none")', 'XLOOKUP(65,Data!I2:I6,Data!J2:J6,,-1)',
  'XLOOKUP(65,Data!I2:I6,Data!J2:J6,,1)', 'XLOOKUP("Apple",Data!A2:A9,Data!C2:C9,,0,-1)', 'CHOOSE(2,"a","b","c")', 'CHOOSE(4,"a")',
  'ROW()', 'COLUMN()', 'ROW(Data!C5)', 'COLUMN(Data!C5)', 'ROWS(Data!A2:C9)', 'COLUMNS(Data!A2:C9)', 'SUM(OFFSET(Data!B2,1,0,3,1))',
  'OFFSET(Data!A1,2,2)', 'INDIRECT("Data!B3")', 'SUM(INDIRECT("Data!C2:C4"))', 'INDIRECT("Nowhere!A1")',
  // Conditional
  'SUMIF(Data!B2:B9,">3")', 'SUMIF(Data!A2:A9,"Apple",Data!C2:C9)', 'SUMIF(Data!A2:A9,"A*",Data!B2:B9)', 'SUMIF(Data!A2:A9,"apple",Data!C2)',
  'COUNTIF(Data!B2:B9,">=4")', 'COUNTIF(Data!A2:A9,"apple")', 'COUNTIF(Data!A2:A9,"<>Apple")', 'COUNTIF(Data!B2:B9,12)',
  'COUNTIF(Data!B2:B9,"12")', 'COUNTIF(Data!B2:B9,"=12")', 'COUNTIF(Data!F1:F9,"")', 'COUNTIF(Data!F1:F9,"=")', 'COUNTIF(Data!A1:E9,"<>")',
  'COUNTIF(Data!A2:A9,"*")', 'COUNTIF(Data!E2:E9,TRUE)', 'COUNTIF(Data!E2:E9,"TRUE")', 'COUNTIF(Data!G2:G3,"#N/A")',
  'COUNTIF(Data!A2:A9,"<C")', 'COUNTIF(Data!F:F,"")', 'COUNTIF(Data!B2:B9,"<>3")', 'AVERAGEIF(Data!B2:B9,">2")',
  'AVERAGEIF(Data!B2:B9,">100")', 'SUMIFS(Data!C2:C9,Data!A2:A9,"Apple",Data!B2:B9,">2")', 'COUNTIFS(Data!A2:A9,"Apple",Data!E2:E9,FALSE)',
  'AVERAGEIFS(Data!C2:C9,Data!A2:A9,"<>Apple")', 'MAXIFS(Data!C2:C9,Data!A2:A9,"Apple")', 'MINIFS(Data!C2:C9,Data!A2:A9,"Apple")',
  // Dates and times
  'DATE(2024,1,31)', 'DATE(2024,14,1)', 'DATE(2024,1,0)', 'DATE(99,1,1)', 'DATE(1900,3,1)', 'DATE(1900,1,1)', 'YEAR(45678)', 'MONTH(45678)',
  'DAY(45678)', 'WEEKDAY(45678)', 'WEEKDAY(45678,2)', 'WEEKDAY(45678,3)', 'TIME(12,30,0)', 'TIME(25,0,0)', 'HOUR(0.75)', 'MINUTE(0.7604)',
  'SECOND(0.5000115)', 'EDATE(DATE(2024,1,31),1)', 'EOMONTH(DATE(2024,1,15),1)', 'EOMONTH(DATE(2024,1,15),-1)',
  'DAYS(DATE(2024,3,1),DATE(2024,1,1))', 'DATE(2024,2,29)-DATE(2023,2,28)', 'YEAR("2024-03-05")', 'YEAR(-1)', 'DATEVALUE("2024-03-05")',
  'MONTH(Data!D5)',
  // Errors
  '#N/A+1', '#DIV/0!&"x"', 'Data!G2', 'Data!G2+1', '#REF!', 'SUM(1,#VALUE!)', 'IF(#N/A,1,2)', 'ISERROR(#NULL!)',
  // More numbers as text, and the last + or - of a formula near zero
  '1E-16&""', '1E-17&""', '1E-18&""', '1.2345E-14&""', '1.23456E-14&""', '1.5E-18&""', '1.23456789012345E-5&""', '99999999999999999999&""',
  '1.5E+19&""', '-1E-15&""', '0.5-0.4-0.1', '0.1+0.2-0.3+0', '1+1E-15-1', '1-1E-15-1', '2.3-2.2-0.1', 'Data!C2-1.5', '-(0.1+0.2-0.3)',
  '0.1+0.2-0.3>0', '10*(0.1+0.2)-3', '(0.1+0.2)*10-3', '100-99.99', '1E-14+1-1', 'SUM(0.1,0.2,-0.3)', '{1,2,3}*2', 'ROUND({1.2,2.7},0)',
  // SUBTOTAL, ranks, percentiles, quartiles, the mode (functions-more.js)
  'SUBTOTAL(9,Data!C2:C9)', 'SUBTOTAL(1,Data!C2:C9)', 'SUBTOTAL(2,Data!B2:B9)', 'SUBTOTAL(3,Data!A2:A9)', 'SUBTOTAL(4,Data!C2:C9)',
  'SUBTOTAL(109,Data!C2:C9)', 'SUBTOTAL(12,Data!C2:C9)', 'RANK(4,Data!C2:C9)', 'RANK(4,Data!C2:C9,1)', 'RANK.EQ(2,Data!B2:B9)',
  'RANK.AVG(2,Data!B2:B9)', 'RANK(9,Data!C2:C9)', 'PERCENTILE(Data!C2:C9,0.3)', 'PERCENTILE.INC(Data!C2:C9,1)', 'PERCENTILE.EXC(Data!C2:C9,0.25)',
  'PERCENTILE.EXC(Data!C2:C9,0.05)', 'QUARTILE(Data!C2:C9,1)', 'QUARTILE.INC(Data!C2:C9,3)', 'QUARTILE.EXC(Data!C2:C9,2)', 'QUARTILE(Data!C2:C9,5)',
  'MODE(1,2,2,3,3)', 'MODE.SNGL(Data!B2:B9)', 'MODE(1,2,3)', 'AVERAGEA(Data!E2:E9)', 'MAXA(Data!E2:E9)', 'MINA(Data!A2:A3,-1)',
  // Working days, DATEDIF, weeks, YEARFRAC
  'NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31))', 'NETWORKDAYS(DATE(2024,1,31),DATE(2024,1,1))', 'NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31),{45293,45296})',
  'WORKDAY(DATE(2024,1,5),1)', 'WORKDAY(DATE(2024,1,8),-3)', 'WORKDAY(DATE(2024,1,1),10,DATE(2024,1,2))',
  'DATEDIF(DATE(2020,5,15),DATE(2024,3,10),"Y")', 'DATEDIF(DATE(2020,5,15),DATE(2024,3,10),"M")', 'DATEDIF(DATE(2020,5,15),DATE(2024,3,10),"D")',
  'DATEDIF(DATE(2020,5,15),DATE(2024,3,10),"YM")', 'DATEDIF(DATE(2020,5,15),DATE(2024,3,10),"MD")', 'DATEDIF(DATE(2020,5,15),DATE(2024,3,10),"YD")',
  'DATEDIF(DATE(2024,3,1),DATE(2024,1,1),"D")', 'WEEKNUM(DATE(2024,1,1))', 'WEEKNUM(DATE(2024,1,7))', 'WEEKNUM(DATE(2024,1,7),2)',
  'WEEKNUM(DATE(2024,12,31),21)', 'ISOWEEKNUM(DATE(2021,1,1))', 'ISOWEEKNUM(DATE(2024,12,30))', 'YEARFRAC(DATE(2024,1,1),DATE(2024,7,1))',
  'YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),1)', 'YEARFRAC(DATE(2023,1,31),DATE(2024,2,29),0)', 'YEARFRAC(DATE(2020,3,15),DATE(2024,9,1),1)',
  'YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),2)', 'YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),3)', 'YEARFRAC(DATE(2024,1,31),DATE(2024,3,31),4)',
  // Loans and investments
  'PMT(0.05/12,360,200000)', 'PMT(0,12,1200)', 'PMT(0.06/12,60,-20000,0,1)', 'PV(0.08/12,240,-500)', 'FV(0.06/12,120,-100)', 'FV(0.05,10,-100,-1000,1)',
  'NPER(0.01,-100,1000)', 'NPER(0,-100,1000)', 'RATE(48,-200,8000)', 'RATE(10,-100,1000)', 'NPV(0.1,-10000,3000,4200,6800)',
  'IRR({-70000,12000,15000,18000,21000,26000})', 'IRR({-100,10})',
  // LOOKUP, XMATCH, ADDRESS
  'LOOKUP(4.5,{1,2,3,4,5},{"a","b","c","d","e"})', 'LOOKUP("Cherry",Data!A2:A9)', 'LOOKUP(0,{1,2})', 'XMATCH(4,Data!C2:C9)', 'XMATCH(5,{1,3,5,7},1)',
  'XMATCH(4,{1,3,5,7},-1)', 'XMATCH(4,{1,3,5,7},1)', 'XMATCH("App*",Data!A2:A9,2)', 'XMATCH(2,{1,2,2,3},0,-1)', 'ADDRESS(2,3)', 'ADDRESS(2,3,4)',
  'ADDRESS(2,3,2,FALSE)', 'ADDRESS(1,28,1,TRUE,"My Data")',
  // Text
  'TEXTBEFORE("a-b-c","-")', 'TEXTBEFORE("a-b-c","-",2)', 'TEXTAFTER("a-b-c","-",-1)', 'TEXTAFTER("a-b-c","x")', 'TEXTAFTER("a-b-c","x",1,0,0,"none")',
  'TEXTBEFORE("A-B","b",1,1)', 'NUMBERVALUE("1.234,5",",",".")', 'NUMBERVALUE("12%")', 'NUMBERVALUE(" 3 000 ")', 'FIXED(1234.567)', 'FIXED(1234.567,1,TRUE)',
  'FIXED(-1234.5,-2)', 'DOLLAR(-1234.567)', 'DOLLAR(0.5,0)', 'UNICODE("A")', 'UNICHAR(9731)',
  // Trigonometry, GCD, LCM, rounding to a multiple, COMBIN, TYPE, ERROR.TYPE
  'SIN(PI()/6)', 'COS(0)', 'TAN(1)', 'ASIN(2)', 'ATAN2(1,1)', 'DEGREES(PI())', 'RADIANS(180)', 'GCD(24,36)', 'GCD(7,0)', 'LCM(4,6)', 'LCM(4,0)',
  'CEILING.MATH(4.3)', 'CEILING.MATH(-4.3)', 'CEILING.MATH(-4.3,1,1)', 'FLOOR.MATH(4.7,2)', 'FLOOR.MATH(-4.7)', 'CEILING.MATH(6.7,0.5)', 'COMBIN(8,3)',
  'COMBIN(3,5)', 'TYPE(1)', 'TYPE("a")', 'TYPE(TRUE)', 'TYPE(#N/A)', 'TYPE({1,2})', 'ERROR.TYPE(#DIV/0!)', 'ERROR.TYPE(1)',
  // Table references (the table Fruit on Data!A1:E9: ooxml/tables.js)
  'SUM(Fruit[Qty])', 'SUM(Fruit[Price])', 'COUNTA(Fruit[Name])', 'ROWS(Fruit[#Data])', 'ROWS(Fruit[#All])', 'INDEX(Fruit[#Headers],1,3)',
  'SUM(Fruit[[Qty]:[Price]])', 'SUMIFS(Fruit[Price],Fruit[Name],"Apple")', 'XLOOKUP("Fig",Fruit[Name],Fruit[Price])', 'SUM(fruit[qty])',
  'Fruit[[#Headers],[Price]]', 'COLUMNS(Fruit[#All])', 'SUBTOTAL(109,Fruit[Price])', 'Fruit[[#This Row],[Name]]',
  'AVERAGE(Fruit[[#Data],[Price]])', 'COUNTA(Fruit[[#Headers],[#Data],[Name]])', 'MATCH("Cherry",Fruit[Name],0)', 'SUM(Fruit[[#All],[Qty]:[Price]])',
];
function formulas() {
  const str = (ref, s) => `<c r="${ref}" t="inlineStr"><is><t${/^\s|\s$/.test(s) ? ' xml:space="preserve"' : ''}>${esc(s)}</t></is></c>`;
  const num = (ref, v) => `<c r="${ref}"><v>${v}</v></c>`;
  const cell = (ref, v) => (v == null ? '' : typeof v === 'number' ? num(ref, v) : typeof v === 'boolean' ? `<c r="${ref}" t="b"><v>${v ? 1 : 0}</v></c>` :
    typeof v === 'object' ? `<c r="${ref}" t="e"><v>${v.err}</v></c>` : str(ref, v));
  const data = [
    ['Name', 'Qty', 'Price', 'Date', 'Flag', null, 'Err', null, 'Score', 'Grade'],
    ['Apple', 3, 1.5, 45292, true, null, { err: '#N/A' }, null, 0, 'F'],
    ['Banana', 5, 0.25, 45323, false, null, 5, null, 50, 'D'],
    ['Cherry', '12', 4, 45352, true, null, null, null, 60, 'C'],
    ['Apple', 2, 2, 45383, null, null, null, null, 70, 'B'],
    ['Elder', 8, 3.5, 45413, false, null, null, null, 80, 'A'],
    ['Fig', null, 6, 45444, true],
    ['Grape', 4, 0.75, 45474, 'yes'],
    ['Apple', 6, 1.25, 45505, false],
  ];
  const dataRows = data.map((row, i) => `<row r="${i + 1}">${row.map((v, c) => cell(`${col(c)}${i + 1}`, v)).join('')}</row>`);
  const calcRows = PROBE_FORMULAS.map((f, i) => {
    const r = i + 2;
    // (D2:D10 a chain, D3:D10 as one shared formula; D11 their sum; E2 reads E3, below it.)
    const chain = r === 2 ? num('D2', 1) : r <= 10 ? `<c r="D${r}">${r === 3 ? '<f t="shared" ref="D3:D10" si="0">D2+1</f>' : '<f t="shared" si="0"/>'}</c>` :
      r === 11 ? '<c r="D11"><f>SUM(D2:D10)</f></c>' : '';
    const fwd = r === 2 ? '<c r="E2"><f>E3*2</f></c>' : r === 3 ? num('E3', 5) : '';
    // (F2:F4 and F12: a table's This Row from another sheet, by the row's number; row 12 is past the table's rows.)
    const thisRow = r <= 4 || r === 12 ? `<c r="F${r}"><f>Fruit[[#This Row],[Price]]*2</f></c>` : '';
    // (As the file has them: newer functions with their _xlfn. prefix.)
    return `<row r="${r}">${str(`A${r}`, f)}<c r="B${r}"><f>${esc(normalizeFormula(f))}</f></c>${chain}${fwd}${thisRow}</row>`;
  });
  const sheet = (rows, cols, tail = '') => `${DECL}<worksheet xmlns="${X}" xmlns:r="${R}"><sheetFormatPr defaultRowHeight="15"/>${cols}<sheetData>${rows.join('')}</sheetData>${tail}</worksheet>`;
  const table = `${DECL}<table xmlns="${X}" id="1" name="Fruit" displayName="Fruit" ref="A1:E9" totalsRowShown="0"><autoFilter ref="A1:E9"/>` +
    '<tableColumns count="5"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Qty"/><tableColumn id="3" name="Price"/><tableColumn id="4" name="Date"/><tableColumn id="5" name="Flag"/></tableColumns>' +
    '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>';
  write('formulas.xlsx', {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/worksheets/sheet2.xml', CT('worksheet')],
      ['xl/worksheets/sheet3.xml', CT('worksheet')], ['xl/styles.xml', CT('styles')], ['xl/tables/table1.xml', CT('table')], ...PROPS_CT]),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><sheets><sheet name="Calc" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/>` +
      '<sheet name="My Data" sheetId="3" r:id="rId3"/></sheets><definedNames><definedName name="Items">Data!$A$2:$A$9</definedName>' +
      '<definedName name="Rate">Data!$C$2</definedName></definedNames></workbook>',
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), 'worksheets/sheet1.xml'], ['rId2', RT('worksheet'), 'worksheets/sheet2.xml'],
      ['rId3', RT('worksheet'), 'worksheets/sheet3.xml'], ['rId4', RT('styles'), 'styles.xml']]),
    'xl/worksheets/sheet1.xml': sheet(calcRows, '<cols><col min="1" max="1" width="50.7109375" customWidth="1"/><col min="2" max="2" width="30.7109375" customWidth="1"/></cols>'),
    'xl/worksheets/sheet2.xml': sheet(dataRows, '', '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>'),
    'xl/worksheets/_rels/sheet2.xml.rels': rels([['rId1', RT('table'), '../tables/table1.xml']]),
    'xl/tables/table1.xml': table,
    'xl/worksheets/sheet3.xml': sheet([`<row r="1">${num('A1', 100)}</row>`, `<row r="2">${str('A2', 'text')}</row>`], ''),
    'xl/styles.xml': STYLES,
  });
}

// Notes and a threaded comment as Excel 365 writes them (its parts and XML read from a workbook Excel saved): a note's
// text in comments1.xml with its box in a VML drawing; a threaded comment in threadedComments with its people in
// persons, and a placeholder note for older versions. The people are made up.
function notes() {
  const VML_NS = 'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"';
  const shape = (n, r, c) => `<v:shape id="_x0000_s${1024 + n}" type="#_x0000_t202" style='position:absolute;margin-left:${60 + c * 48}pt;margin-top:${r * 15}pt;width:108pt;height:59.25pt;z-index:${n};visibility:hidden' fillcolor="infoBackground [80]" strokecolor="none [81]" o:insetmode="auto">` +
    `<v:fill color2="infoBackground [80]"/><v:shadow color="none [81]" obscured="t"/><v:path o:connecttype="none"/><v:textbox style='mso-direction-alt:auto'><div style='text-align:left'></div></v:textbox>` +
    `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${c + 1}, 15, ${r}, 10, ${c + 3}, 31, ${r + 4}, 9</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${r}</x:Row><x:Column>${c}</x:Column></x:ClientData></v:shape>`;
  const vml = `<xml ${VML_NS}><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>` +
    shape(1, 1, 1) + shape(2, 2, 2) + shape(3, 3, 3) + '</xml>';
  const run = (t, bold) => `<r><rPr>${bold ? '<b/>' : ''}<sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr><t xml:space="preserve">${esc(t)}</t></r>`;
  const TC = '{6F1D3C2A-0B4E-4D7A-9C1E-2A3B4C5D6E7F}';
  const placeholder = ['[Threaded comment]', '', 'Your version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924',
    '', 'Comment:', '    Is this the right total?', 'Reply:', '    Yes, checked against the invoice.'].join('\n');
  const comments = `${DECL}<comments xmlns="${X}"><authors><author>Avery Chen</author><author>tc=${TC}</author></authors><commentList>` +
    `<comment ref="B2" authorId="0" shapeId="0"><text>${run('Avery Chen:', true)}${run('\nCheck this total', false)}</text></comment>` +
    `<comment ref="C3" authorId="1" shapeId="0"><text><t>${esc(placeholder)}</t></text></comment>` +
    `<comment ref="D4" authorId="0" shapeId="0"><text>${run('Avery Chen:', true)}${run('\nFirst line\nSecond line', false)}</text></comment></commentList></comments>`;
  const TN = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';
  const AVERY = '{1A2B3C4D-5E6F-4A1B-8C2D-3E4F5A6B7C8D}';
  const SAM = '{9F8E7D6C-5B4A-4C3D-8E2F-1A0B9C8D7E6F}';
  const threaded = `${DECL}<ThreadedComments xmlns="${TN}" xmlns:x="${X}"><threadedComment ref="C3" dT="2026-01-01T09:30:00.00" personId="${AVERY}" id="${TC}"><text>Is this the right total?</text></threadedComment>` +
    `<threadedComment ref="C3" dT="2026-01-02T14:05:00.00" personId="${SAM}" id="{2B3C4D5E-6F7A-4B8C-9D0E-1F2A3B4C5D6E}" parentId="${TC}"><text>Yes, checked against the invoice.</text></threadedComment></ThreadedComments>`;
  const persons = `${DECL}<personList xmlns="${TN}" xmlns:x="${X}"><person displayName="Avery Chen" id="${AVERY}" userId="avery@example.com" providerId="None"/><person displayName="Sam Rivera" id="${SAM}" userId="sam@example.com" providerId="None"/></personList>`;
  const v = (ref, n) => `<c r="${ref}"><v>${n}</v></c>`;
  const t = (ref, text) => `<c r="${ref}" t="inlineStr"><is><t>${esc(text)}</t></is></c>`;
  const data = `<row r="1">${t('A1', 'Item')}${t('B1', 'Total')}</row><row r="2">${t('A2', 'apples')}${v('B2', 12)}</row><row r="3">${t('A3', 'pears')}${v('C3', 30)}</row><row r="4">${t('A4', 'figs')}${v('D4', 7)}</row><row r="5">${t('A5', 'plain')}${v('B5', 1)}</row>`;
  write('notes.xlsx', {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/styles.xml', CT('styles')],
      ['xl/comments1.xml', CT('comments')], ['xl/threadedComments/threadedComment1.xml', 'application/vnd.ms-excel.threadedcomments+xml'], ['xl/persons/person.xml', 'application/vnd.ms-excel.person+xml'], ...PROPS_CT])
      .replace('<Default Extension="xml"', '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Default Extension="xml"'),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><sheets><sheet name="Notes" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), 'worksheets/sheet1.xml'], ['rId2', RT('styles'), 'styles.xml'], ['rId3', 'http://schemas.microsoft.com/office/2017/10/relationships/person', 'persons/person.xml']]),
    'xl/worksheets/sheet1.xml': `${DECL}<worksheet xmlns="${X}" xmlns:r="${R}"><sheetData>${data}</sheetData><legacyDrawing r:id="rId1"/></worksheet>`,
    'xl/worksheets/_rels/sheet1.xml.rels': rels([['rId1', RT('vmlDrawing'), '../drawings/vmlDrawing1.vml'], ['rId2', RT('comments'), '../comments1.xml'], ['rId3', 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment', '../threadedComments/threadedComment1.xml']]),
    'xl/drawings/vmlDrawing1.vml': vml,
    'xl/comments1.xml': comments,
    'xl/threadedComments/threadedComment1.xml': threaded,
    'xl/persons/person.xml': persons,
    'xl/styles.xml': STYLES,
  });
}

// Charts as Excel writes them (a drawing of two-cell anchors, each a graphic frame pointing to a chart part): monthly
// sales for three regions, drawn four ways.
function charts() {
  const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
  const regions = { North: [120, 135, 150, 145, 170, 190], South: [80, 95, 90, 110, 105, 125], West: [60, 72, 85, 80, 98, 104] };
  const names = Object.keys(regions);
  const cols = ['B', 'C', 'D'];
  const text = (ref, t) => `<c r="${ref}" t="inlineStr"><is><t>${t}</t></is></c>`;
  const head = `<row r="1">${text('A1', 'Month')}${names.map((n, i) => text(`${cols[i]}1`, n)).join('')}</row>`;
  const body = months.map((m, k) => `<row r="${k + 2}">${text(`A${k + 2}`, m)}${names.map((n, i) => `<c r="${cols[i]}${k + 2}"><v>${regions[n][k]}</v></c>`).join('')}</row>`).join('');
  const strCache = (list) => `<c:strCache><c:ptCount val="${list.length}"/>${list.map((t, i) => `<c:pt idx="${i}"><c:v>${t}</c:v></c:pt>`).join('')}</c:strCache>`;
  const numCache = (list) => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${list.length}"/>${list.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache>`;
  const ser = (name, i, extra = '') => {
    const col = cols[names.indexOf(name)];
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:strRef><c:f>Sales!$${col}$1</c:f>${strCache([name])}</c:strRef></c:tx>${extra}` +
      `<c:cat><c:strRef><c:f>Sales!$A$2:$A$7</c:f>${strCache(months)}</c:strRef></c:cat><c:val><c:numRef><c:f>Sales!$${col}$2:$${col}$7</c:f>${numCache(regions[name])}</c:numRef></c:val></c:ser>`;
  };
  const axes = (dir) => `<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${dir === 'bar' ? 'l' : 'b'}"/><c:numFmt formatCode="General" sourceLinked="1"/>` +
    '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="20"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>' +
    `<c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${dir === 'bar' ? 'b' : 'l'}"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/>` +
    '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="10"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>';
  const title = (t) => (t ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${t}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>');
  const space = (inner, legend) => `${DECL}<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:roundedCorners val="0"/><c:chart>${inner}` +
    `${legend ? `<c:legend><c:legendPos val="${legend}"/><c:overlay val="0"/></c:legend>` : ''}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
  const parts = [
    space(`${title('Sales by month')}<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${ser('North', 0)}${ser('South', 1)}<c:gapWidth val="219"/><c:overlap val="-27"/><c:axId val="10"/><c:axId val="20"/></c:barChart>${axes('col')}</c:plotArea>`, 'b'),
    space(`${title(null)}<c:plotArea><c:layout/><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${names.map((n, i) => ser(n, i, '<c:marker><c:symbol val="none"/></c:marker>')).join('')}<c:marker val="1"/><c:axId val="10"/><c:axId val="20"/></c:lineChart>${axes('col')}</c:plotArea>`, 'r'),
    space(`<c:title><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:pieChart><c:varyColors val="1"/>${ser('North', 0)}<c:firstSliceAng val="0"/></c:pieChart></c:plotArea>`, 'r'),
    space(`${title('Sales, stacked')}<c:plotArea><c:layout/><c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/><c:varyColors val="0"/>${names.map((n, i) => ser(n, i)).join('')}<c:gapWidth val="150"/><c:overlap val="100"/><c:axId val="10"/><c:axId val="20"/></c:barChart>${axes('bar')}</c:plotArea>`, 'b'),
  ];
  const at = [[5, 1, 13, 16], [14, 1, 22, 16], [5, 17, 13, 32], [14, 17, 22, 32]];
  const anchor = ([c0, r0, c1, r1], i) => `<xdr:twoCellAnchor><xdr:from><xdr:col>${c0}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r0}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
    `<xdr:to><xdr:col>${c1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${i + 2}" name="Chart ${i + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" r:id="rId${i + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
  const files = {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/styles.xml', CT('styles')],
      ['xl/drawings/drawing1.xml', 'application/vnd.openxmlformats-officedocument.drawing+xml'], ...parts.map((x, i) => [`xl/charts/chart${i + 1}.xml`, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml']), ...PROPS_CT]),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><sheets><sheet name="Sales" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), 'worksheets/sheet1.xml'], ['rId2', RT('styles'), 'styles.xml']]),
    'xl/worksheets/sheet1.xml': `${DECL}<worksheet xmlns="${X}" xmlns:r="${R}"><sheetData>${head}${body}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><drawing r:id="rId1"/></worksheet>`,
    'xl/worksheets/_rels/sheet1.xml.rels': rels([['rId1', RT('drawing'), '../drawings/drawing1.xml']]),
    'xl/drawings/drawing1.xml': `${DECL}<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:r="${R}">${at.map(anchor).join('')}</xdr:wsDr>`,
    'xl/drawings/_rels/drawing1.xml.rels': rels(parts.map((x, i) => [`rId${i + 1}`, RT('chart'), `../charts/chart${i + 1}.xml`])),
    'xl/styles.xml': STYLES,
  };
  parts.forEach((x, i) => { files[`xl/charts/chart${i + 1}.xml`] = x; });
  write('charts.xlsx', files);
}

// The sheet of tools/excel-compare/anchor-cases.mjs, as Excel writes its parts: a drawing with a chart and a rectangle
// on two-cell anchors, notes in the comments part and the VML that shows them; with `table`, also the table Stock on
// B2:D12 and a formula reading two of its columns in A20 (anchors-table.xlsx).
function anchors(table = false) {
  const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
  const items = ['Bolt', 'Nut', 'Screw', 'Washer', 'Rivet', 'Pin', 'Clip', 'Hook', 'Ring', 'Peg'];
  const qty = items.map((_, i) => (i * 7) % 13 + 2);
  const price = items.map((_, i) => ((i * 5) % 9) + 0.5);
  const t = (ref, text) => `<c r="${ref}" t="inlineStr"><is><t>${esc(text)}</t></is></c>`;
  const v = (ref, n) => `<c r="${ref}"><v>${n}</v></c>`;
  const sum = qty.reduce((a, b) => a + b, 0) + items.length;
  const formulaRow = table ? `<row r="20"><c r="A20"><f>SUM(Stock[Qty])+COUNTA(Stock[Price])</f><v>${sum}</v></c></row>` : '';
  const rows = [`<row r="2">${t('B2', 'Item')}${t('C2', 'Qty')}${t('D2', 'Price')}</row>`, ...items.map((it, i) => `<row r="${i + 3}">${t(`B${i + 3}`, it)}${v(`C${i + 3}`, qty[i])}${v(`D${i + 3}`, price[i])}</row>`)].join('') + formulaRow;
  const num = `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="10"/>${qty.map((q, i) => `<c:pt idx="${i}"><c:v>${q}</c:v></c:pt>`).join('')}</c:numCache>`;
  const chart = `${DECL}<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:roundedCorners val="0"/><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>` +
    `<c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>Stock!$C$3:$C$12</c:f>${num}</c:numRef></c:val></c:ser><c:gapWidth val="219"/><c:axId val="10"/><c:axId val="20"/></c:barChart>` +
    '<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx></c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>';
  const at = (c0, r0, c1, r1) => `<xdr:from><xdr:col>${c0}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r0}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${c1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>`;
  const drawing = `${DECL}<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:r="${R}">` +
    `<xdr:twoCellAnchor>${at(5, 2, 11, 14)}<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>` +
    `<xdr:twoCellAnchor>${at(12, 4, 15, 8)}<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="3" name="Rectangle 2"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;
  const notesAt = [[0, 0], [5, 1], [11, 3], [19, 5]]; // A1, B6, D12, F20 (row, column)
  const VML_NS = 'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"';
  const box = ([r, c], n) => `<v:shape id="_x0000_s${1025 + n}" type="#_x0000_t202" style='position:absolute;margin-left:${60 + c * 48}pt;margin-top:${r * 15}pt;width:108pt;height:59.25pt;z-index:${n + 1};visibility:hidden' fillcolor="infoBackground [80]" strokecolor="none [81]" o:insetmode="auto">` +
    `<v:fill color2="infoBackground [80]"/><v:shadow color="none [81]" obscured="t"/><v:path o:connecttype="none"/><v:textbox style='mso-direction-alt:auto'><div style='text-align:left'></div></v:textbox>` +
    `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${c + 1}, 15, ${r}, 10, ${c + 3}, 31, ${r + 4}, 9</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${r}</x:Row><x:Column>${c}</x:Column></x:ClientData></v:shape>`;
  const vml = `<xml ${VML_NS}><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>${notesAt.map(box).join('')}</xml>`;
  const comments = `${DECL}<comments xmlns="${X}"><authors><author>Avery Chen</author></authors><commentList>${notesAt.map(([r, c]) => `<comment ref="${col(c)}${r + 1}" authorId="0" shapeId="0"><text><t>note ${col(c)}${r + 1}</t></text></comment>`).join('')}</commentList></comments>`;
  const tablePart = `${DECL}<table xmlns="${X}" id="1" name="Stock" displayName="Stock" ref="B2:D12" totalsRowShown="0"><autoFilter ref="B2:D12"/><tableColumns count="3"><tableColumn id="1" name="Item"/><tableColumn id="2" name="Qty"/><tableColumn id="3" name="Price"/></tableColumns>` +
    '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>';
  write(table ? 'anchors-table.xlsx' : 'anchors.xlsx', {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/styles.xml', CT('styles')], ['xl/comments1.xml', CT('comments')], ...(table ? [['xl/tables/table1.xml', CT('table')]] : []),
      ['xl/drawings/drawing1.xml', 'application/vnd.openxmlformats-officedocument.drawing+xml'], ['xl/charts/chart1.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'], ...PROPS_CT])
      .replace('<Default Extension="xml"', '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Default Extension="xml"'),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><sheets><sheet name="Stock" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), 'worksheets/sheet1.xml'], ['rId2', RT('styles'), 'styles.xml']]),
    'xl/worksheets/sheet1.xml': `${DECL}<worksheet xmlns="${X}" xmlns:r="${R}"><sheetData>${rows}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><drawing r:id="rId1"/><legacyDrawing r:id="rId2"/>${table ? '<tableParts count="1"><tablePart r:id="rId4"/></tableParts>' : ''}</worksheet>`,
    'xl/worksheets/_rels/sheet1.xml.rels': rels([['rId1', RT('drawing'), '../drawings/drawing1.xml'], ['rId2', RT('vmlDrawing'), '../drawings/vmlDrawing1.vml'], ['rId3', RT('comments'), '../comments1.xml'], ...(table ? [['rId4', RT('table'), '../tables/table1.xml']] : [])]),
    ...(table ? { 'xl/tables/table1.xml': tablePart } : {}),
    'xl/drawings/drawing1.xml': drawing,
    'xl/drawings/_rels/drawing1.xml.rels': rels([['rId1', RT('chart'), '../charts/chart1.xml']]),
    'xl/charts/chart1.xml': chart,
    'xl/drawings/vmlDrawing1.vml': vml,
    'xl/comments1.xml': comments,
    'xl/styles.xml': STYLES,
  });
}

if (process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1])) {
  formulas();
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
  notes();
  charts();
  anchors();
  anchors(true);
}
