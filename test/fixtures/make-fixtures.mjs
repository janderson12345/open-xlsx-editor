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
// rotate.xlsx   rotated text (angles up and down with each alignment, stacked, rows Excel sizes) and shrink to fit,
//                for comparing with Excel's PDF (probes-format.mjs; docs/LAYOUT.md).
// special.xlsx  something for each of Go To Special's choices (special.mjs).
// sparklines.xlsx  sparklines of each type and option, as Excel writes them (sparklines.mjs).
// tables.xlsx    a table in each built-in table style, one in a style of its own, one with cells' own formats
//                (tables.mjs).
// print-*.xlsx   probes for the print model: print area, titles and breaks; scale; fit to a page wide; headers and
//                footers (probes-print.mjs).
// The zips are written with a fixed date, so the files only change when this script does.
//
// Layout: fixture-common.mjs (helpers, shared parts), probes-format.mjs (widths, General, fonts, fills, alignment,
// number formats, row heights), probes-formulas.mjs, parts-drawings.mjs (notes, charts, anchors), probes-print.mjs.
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { special } from './special.mjs';
import { sparklines } from './sparklines.mjs';
import { tables } from './tables.mjs';
import { write, contentTypes, rels, RT, CT, DECL, X, R, ROOT_RELS, docProps, PROPS_CT, STYLES, esc, col, stringTable } from './fixture-common.mjs';
import { widths, general, fonts, fills, align, rotate, numfmts, rowHeights } from './probes-format.mjs';
import { formulas } from './probes-formulas.mjs';
import { notes, charts, anchors } from './parts-drawings.mjs';
import { printProbes } from './probes-print.mjs';

// The tests read the probes' tables from here.
export * from './probes-format.mjs';
export * from './probes-formulas.mjs';

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

if (process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1])) {
  formulas();
  rowHeights();
  numfmts();
  align();
  rotate();
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
  printProbes();
  special({ write, contentTypes, rels, RT, CT, DECL, X, R, ROOT_RELS, docProps, PROPS_CT, STYLES, esc });
  sparklines();
  tables();
}
