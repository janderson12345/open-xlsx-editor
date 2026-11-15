// sparklines.xlsx: a sheet of sparklines of each type and option, written as Excel writes its x14:sparklineGroups
// extension (tools/excel-compare/excel-sparklines.ps1, docs/SPARKLINES.md), for test/sparklines.test.js, the browser
// test and the drawing comparison with Excel's PDF. Made-up numbers in B:M of each row, a label in A, the sparkline in
// O (30 characters wide, rows 36pt tall); column N is hidden and holds an outlier for the hidden-data rows; dates for
// the date axis in B32:M32.
import { write, contentTypes, rels, RT, CT, DECL, X, R, ROOT_RELS, docProps, PROPS_CT, STYLES, THEME, esc, col } from './fixture-common.mjs';

export const SPARK_ROWS = [
  // [label, numbers (null: empty, a string: text), group attributes, colours (default: Excel's Style 1 in theme
  //  colours, as its ribbon writes them), extra { at, date, n (column N's value) }]
  ['Line', [3, 5, 2, 8, 6, 9, 4, 7, 5, 10, 6, 8], ''],
  ['Line, every point', [3, -2, 5, 8, -4, 9, 1, 7, -6, 10, 2, 4], 'lineWeight="2.25" markers="1" high="1" low="1" first="1" last="1" negative="1"', 'vivid'],
  ['First high, last low', [10, 4, 6, 10, 3, 5, 2, 7, 8, 6, 4, 2], 'high="1" low="1" first="1" last="1"', 'vivid'],
  ['Negative points', [3, -2, 5, 8, -4, 9, 1, 7, -6, 10, 2, 4], 'negative="1"', 'vivid'],
  ['Column', [4, -3, 6, 2, -5, 7, 3, -1, 8, 5, -2, 6], 'type="column" negative="1"'],
  ['Column, points', [10, 4, 6, 10, 3, 5, 2, 7, 8, 6, 4, 2], 'type="column" high="1" low="1" first="1" last="1"', 'vivid'],
  ['Column, negative', [-3, -5, -2, -8, -6, -9, -4, -7, -5, -10, -6, -8], 'type="column"'],
  ['Column from 0', [5, 9, 3, 7, 11, 2, 8, 6, 10, 4, 12, 1], 'type="column" manualMin="0" minAxisType="custom"'],
  ['Win/loss', [1, -1, 2, -3, 0, 4, -2, 1, -1, 3, -4, 2], 'type="stacked" negative="1"'],
  ['Line, axis', [-2, -5, 1, 4, -3, 6, 2, -4, 5, -1, 3, 0], 'displayXAxis="1"', 'vivid'],
  ['Line, axis, positive', [3, 5, 2, 8, 6, 9, 4, 7, 5, 10, 6, 8], 'displayXAxis="1"', 'vivid'],
  ['Column, axis', [4, -3, 6, 2, -5, 7, 3, -1, 8, 5, -2, 6], 'type="column" displayXAxis="1"', 'vivid'],
  ['Empty as gaps', [4, 6, null, 8, 5, null, null, 7, 9, 3, null, 6], 'displayEmptyCellsAs="gap"'],
  ['Empty as zero', [4, 6, null, 8, 5, null, null, 7, 9, 3, null, 6], ''],
  ['Empty connected', [4, 6, null, 8, 5, null, null, 7, 9, 3, null, 6], 'displayEmptyCellsAs="span"'],
  ['Column, empty', [4, 6, null, 8, 5, null, null, 7, 9, 3, null, 6], 'type="column" displayEmptyCellsAs="gap"'],
  ['Group scale', [2, 4, 3, 5, 4, 6, 5, 7, 6, 8, 7, 9], 'type="column" minAxisType="group" maxAxisType="group"', null, { at: 'O18:O19' }],
  ['(same group)', [10, 14, 12, 16, 15, 18, 17, 20, 19, 22, 21, 24], null],
  ['Custom 0 to 20', [3, 8, 1, 12, 6, 15, 4, 9, 2, 11, 7, 14], 'manualMax="20" manualMin="0" minAxisType="custom" maxAxisType="custom"'],
  ['Custom max 8', [3, 8, 1, 12, 6, 15, 4, 9, 2, 11, 7, 14], 'type="column" manualMax="8" maxAxisType="custom"'],
  ['Date axis', [5, 7, 6, 9, 8, 11, 10, 9, 12, 14, 13, 15], 'dateAxis="1" markers="1"', null, { date: true }],
  ['Right to left', [3, 5, 2, 8, 6, 9, 4, 7, 5, 10, 6, 8], 'type="column" rightToLeft="1" first="1"', 'vivid'],
  ['Hidden data left out', [3, 5, 2, 8, 6, 9, 4, 7, 5, 10, 6, 8], '', null, { n: 30 }],
  ['Hidden data shown', [3, 5, 2, 8, 6, 9, 4, 7, 5, 10, 6, 8], 'displayHidden="1"', null, { n: 30 }],
  ['All the same', [5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5], 'markers="1"'],
  ['Text in the data', [4, 6, 'n/a', 8, 5, 7, 3, 'n/a', 9, 3, 6, 6], 'displayEmptyCellsAs="gap" markers="1"'],
  ['Thin line', [3, 5, 2, 8, 6, 9, 4, 7, 5, 10, 6, 8], 'lineWeight="0.25"'],
  ['Heavy line, markers', [3, 5, 2, 8, 6, 9, 4, 7, 5, 10, 6, 8], 'lineWeight="6" markers="1"'],
];

/** Dates for the date axis (B32:M32): serials with uneven gaps. */
export const SPARK_DATES = [45658, 45659, 45661, 45665, 45666, 45670, 45671, 45680, 45681, 45682, 45690, 45700];

// Excel's default style as its ribbon writes it; and one of plain colours, each part told apart.
const STYLE1 = '<x14:colorSeries theme="4" tint="-0.499984740745262"/><x14:colorNegative theme="5"/><x14:colorAxis rgb="FF000000"/>' +
  '<x14:colorMarkers theme="4" tint="-0.499984740745262"/><x14:colorFirst theme="4" tint="0.39997558519241921"/>' +
  '<x14:colorLast theme="4" tint="0.39997558519241921"/><x14:colorHigh theme="4"/><x14:colorLow theme="4"/>';
const VIVID = '<x14:colorSeries rgb="FF13458B"/><x14:colorNegative rgb="FFFF0000"/><x14:colorAxis rgb="FFC00000"/>' +
  '<x14:colorMarkers rgb="FFFFA500"/><x14:colorFirst rgb="FF00FFFF"/><x14:colorLast rgb="FF808080"/>' +
  '<x14:colorHigh rgb="FF00FF00"/><x14:colorLow rgb="FFFF00FF"/>';

export function sparklines() {
  const cell = (ref, v) => (typeof v === 'string' ? `<c r="${ref}" t="inlineStr"><is><t>${esc(v)}</t></is></c>` : `<c r="${ref}"><v>${v}</v></c>`);
  const rows = [];
  const groups = [];
  SPARK_ROWS.forEach(([label, values, attrs, colors, extra = {}], i) => {
    const r = i + 2;
    const cells = [cell(`A${r}`, label), ...values.map((v, j) => (v == null ? '' : cell(`${col(j + 1)}${r}`, v)))];
    if (extra.n != null) cells.push(cell(`N${r}`, extra.n));
    rows.push(`<row r="${r}" ht="36" customHeight="1">${cells.join('')}</row>`);
    if (attrs == null) return; // (a row of the group above)
    const at = extra.at || `O${r}`;
    const lines = at.includes(':') ? [[`B${r}:M${r}`, `O${r}`], [`B${r + 1}:M${r + 1}`, `O${r + 1}`]] : [[`B${r}:${extra.n != null ? 'N' : 'M'}${r}`, at]];
    groups.push(`<x14:sparklineGroup ${attrs}>${colors === 'vivid' ? VIVID : STYLE1}` +
      (extra.date ? '<xm:f>Sparklines!B32:M32</xm:f>' : '') +
      `<x14:sparklines>${lines.map(([f, sq]) => `<x14:sparkline><xm:f>Sparklines!${f}</xm:f><xm:sqref>${sq}</xm:sqref></x14:sparkline>`).join('')}</x14:sparklines></x14:sparklineGroup>`);
  });
  rows.push(`<row r="32">${[cell('A32', 'Dates'), ...SPARK_DATES.map((d, j) => cell(`${col(j + 1)}32`, d))].join('')}</row>`);
  // (Excel writes the groups newest first; the order makes no difference to the drawing.)
  const ext = '<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
    `<x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">${groups.join('')}</x14:sparklineGroups></ext></extLst>`;
  const sheet = `${DECL}<worksheet xmlns="${X}" xmlns:r="${R}"><sheetFormatPr defaultRowHeight="15"/>` +
    '<cols><col min="1" max="1" width="20" customWidth="1"/><col min="14" max="14" width="9.140625" hidden="1"/><col min="15" max="15" width="30" customWidth="1"/></cols>' +
    `<sheetData>${rows.join('')}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>${ext}</worksheet>`;
  write('sparklines.xlsx', {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/styles.xml', CT('styles')], ['xl/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'], ...PROPS_CT]),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><sheets><sheet name="Sparklines" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), 'worksheets/sheet1.xml'], ['rId2', RT('styles'), 'styles.xml'], ['rId3', RT('theme'), 'theme/theme1.xml']]),
    'xl/worksheets/sheet1.xml': sheet,
    'xl/styles.xml': STYLES,
    'xl/theme/theme1.xml': THEME,
  });
}
