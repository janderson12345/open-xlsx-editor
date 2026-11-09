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

if (process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1])) {
  general();
  basic();
  prefixed();
  large();
  widths();
}
