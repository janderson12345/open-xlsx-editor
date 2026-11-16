// tables.xlsx: a table in each of Excel's 60 built-in table styles, with the Table Style Options taking turns (header
// and banded rows; everything on with a total row; no header with banded columns; first and last column), one in a
// style of the workbook's own (styles.xml's tableStyles: every element, stripes two rows wide) and one with cells of
// their own format over the style, written as Excel writes tables (tools/excel-compare/excel-tables.ps1); and a row of
// accent colours lightened 40, 60 and 80%, plain cells, to tell the tints apart. For test/tables.test.js, the browser
// test and the drawing comparison with Excel's PDF (docs/TABLES.md). Made-up data.
import { write, contentTypes, rels, RT, CT, DECL, X, R, ROOT_RELS, docProps, PROPS_CT, THEME, esc, col } from './fixture-common.mjs';

const STYLE_NAMES = [
  ...Array.from({ length: 21 }, (_, i) => `TableStyleLight${i + 1}`),
  ...Array.from({ length: 28 }, (_, i) => `TableStyleMedium${i + 1}`),
  ...Array.from({ length: 11 }, (_, i) => `TableStyleDark${i + 1}`),
];
/** The options each table takes, in turn: [header, totals, rowStripes, colStripes, first, last]. */
const OPTIONS = [[1, 0, 1, 0, 0, 0], [1, 1, 1, 1, 1, 1], [0, 1, 0, 1, 0, 0], [1, 1, 1, 0, 1, 1]];
const HEAD = ['Region', 'Item', 'Qty', 'Price'];
const DATA = [['North', 'kiwi', 30, 1.5], ['South', 'apple', 10, 0.5], ['East', 'pear', 20, 2], ['West', 'date', 50, 3]];
const CUSTOM = 'Fixture Style';

/** Every table of the fixture: { name, style, ref, r0 (0-based), c0, options, own } (own: cells of their own format). */
export function tableList() {
  const out = [];
  const add = (style, options, own = null) => {
    const i = out.length;
    const r0 = 1 + Math.floor(i / 2) * 8;
    const c0 = i % 2 ? 6 : 1;
    const [header, totals] = options;
    const r1 = r0 + header + DATA.length - 1 + totals;
    out.push({ name: `Table${i + 1}`, style, r0, c0, r1, c1: c0 + 3, options, own, ref: `${col(c0)}${r0 + 1}:${col(c0 + 3)}${r1 + 1}` });
  };
  STYLE_NAMES.forEach((s, i) => add(s, OPTIONS[i % OPTIONS.length]));
  add(CUSTOM, [1, 1, 1, 1, 1, 1]);
  // (Own formats: a yellow fill on the header's first cell, a red font in the data, a medium bottom border.)
  add('TableStyleMedium2', [1, 1, 1, 0, 0, 0], { '0,0': 3, '1,1': 2, '2,2': 4 });
  return out;
}

/** The workbook's styles: our fixtures' Calibri, the own formats' xfs, the accent tints' fills and the custom style. */
function styles() {
  const tints = [0.39997558519241921, 0.59999389629810485, 0.79998168889431442];
  const tintFills = [];
  for (let a = 4; a <= 9; a++) for (const t of tints) tintFills.push(`<fill><patternFill patternType="solid"><fgColor theme="${a}" tint="${t}"/><bgColor indexed="64"/></patternFill></fill>`);
  const dxf = (inner) => `<dxf>${inner}</dxf>`;
  const solid = (c) => `<fill><patternFill patternType="solid"><fgColor ${c}/><bgColor ${c}/></patternFill></fill>`;
  const dxfs = [
    dxf(`<font><color theme="1"/></font><border><left style="medium"><color rgb="FFC00000"/></left><right style="medium"><color rgb="FFC00000"/></right><top style="medium"><color rgb="FFC00000"/></top><bottom style="medium"><color rgb="FFC00000"/></bottom><vertical style="thin"><color theme="5"/></vertical><horizontal style="dashed"><color theme="5"/></horizontal></border>`),
    dxf(`<font><b/><i/><color theme="0"/></font>${solid('theme="5"')}<border><bottom style="double"><color theme="1"/></bottom></border>`),
    dxf(`<font><b/><color theme="5" tint="-0.249977111117893"/></font><border><top style="thick"><color theme="5"/></top></border>`),
    dxf(solid('theme="5" tint="0.59999389629810485"')),
    dxf(solid('theme="9" tint="0.79998168889431442"')),
    dxf(`<font><i/></font>${solid('theme="7" tint="0.79998168889431442"')}`),
    dxf('<font><b/><color rgb="FF7030A0"/></font>'),
    dxf(`<font><color rgb="FFFFFF00"/></font>${solid('rgb="FF000000"')}`),
  ];
  const custom = `<tableStyle name="${CUSTOM}" pivot="0" count="8"><tableStyleElement type="wholeTable" dxfId="0"/><tableStyleElement type="headerRow" dxfId="1"/><tableStyleElement type="totalRow" dxfId="2"/><tableStyleElement type="firstColumn" dxfId="6"/><tableStyleElement type="firstRowStripe" size="2" dxfId="3"/><tableStyleElement type="secondRowStripe" dxfId="4"/><tableStyleElement type="firstColumnStripe" dxfId="5"/><tableStyleElement type="firstHeaderCell" dxfId="7"/></tableStyle>`;
  const fillXfs = tintFills.map((_, i) => `<xf numFmtId="0" fontId="0" fillId="${i + 3}" borderId="0" xfId="0" applyFill="1"/>`);
  return `${DECL}<styleSheet xmlns="${X}">` +
    '<fonts count="2"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>' +
    '<font><sz val="11"/><color rgb="FFFF0000"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>' +
    `<fills count="${3 + tintFills.length}"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>` +
    `<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>${tintFills.join('')}</fills>` +
    '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style="medium"><color indexed="64"/></bottom><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    `<cellXfs count="${5 + fillXfs.length}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>' +
    `<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>${fillXfs.join('')}</cellXfs>` +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    `<dxfs count="${dxfs.length}">${dxfs.join('')}</dxfs><tableStyles count="1" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16">${custom}</tableStyles></styleSheet>`;
}

/** A table part as Excel writes one (excel-tables.ps1's cases): the total row's label and SUM in the first and last
 *  columns, the autoFilter over the header and data rows. */
function tablePart(t, id) {
  const [header, totals, rows, cols, first, last] = t.options;
  const head = `${col(t.c0)}${t.r0 + 1}`;
  const dataEnd = `${col(t.c1)}${t.r1 + 1 - totals}`;
  const columns = HEAD.map((h, j) => `<tableColumn id="${j + 1}" name="${h}"${totals && j === 0 ? ' totalsRowLabel="Total"' : ''}${totals && j === 3 ? ' totalsRowFunction="sum"' : ''}/>`).join('');
  return `${DECL}<table xmlns="${X}" id="${id}" name="${t.name}" displayName="${t.name}" ref="${t.ref}"${header ? '' : ' headerRowCount="0"'}${totals ? ' totalsRowCount="1"' : ' totalsRowShown="0"'}>` +
    (header ? `<autoFilter ref="${head}:${dataEnd}"/>` : '') +
    `<tableColumns count="4">${columns}</tableColumns><tableStyleInfo name="${t.style}" showFirstColumn="${first}" showLastColumn="${last}" showRowStripes="${rows}" showColumnStripes="${cols}"/></table>`;
}

export function tables() {
  const list = tableList();
  const cells = new Map();
  const put = (r, c, xml) => { if (!cells.has(r)) cells.set(r, []); cells.get(r).push([c, xml]); };
  const str = (r, c, s, st) => put(r, c, `<c r="${col(c)}${r + 1}"${st ? ` s="${st}"` : ''} t="inlineStr"><is><t>${esc(s)}</t></is></c>`);
  const num = (r, c, v, st) => put(r, c, `<c r="${col(c)}${r + 1}"${st ? ` s="${st}"` : ''}><v>${v}</v></c>`);
  for (const t of list) {
    const [header, totals] = t.options;
    const own = (i, j) => t.own?.[`${i},${j}`];
    let r = t.r0;
    if (header) { HEAD.forEach((h, j) => str(r, t.c0 + j, h, own(0, j))); r++; }
    DATA.forEach((row, i) => { row.forEach((v, j) => (typeof v === 'number' ? num : str)(r, t.c0 + j, v, own(i + header, j))); r++; });
    if (totals) {
      str(r, t.c0, 'Total');
      put(r, t.c1, `<c r="${col(t.c1)}${r + 1}"><f>SUBTOTAL(109,${t.name}[Price])</f><v>7</v></c>`);
    }
  }
  // (The tints: accents 1 to 6, each lightened 40, 60 and 80%, in two rows of nine below the tables.)
  const tintRow = list[list.length - 1].r1 + 3;
  for (let k = 0; k < 18; k++) num(tintRow + Math.floor(k / 9), 1 + (k % 9), k, 5 + k);
  const rowsXml = [...cells.keys()].sort((x, y) => x - y).map((r) => `<row r="${r + 1}">${cells.get(r).sort((x, y) => x[0] - y[0]).map((x) => x[1]).join('')}</row>`);
  const sheet = `${DECL}<worksheet xmlns="${X}" xmlns:r="${R}"><sheetFormatPr defaultRowHeight="15"/><sheetData>${rowsXml.join('')}</sheetData>` +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
    `<tableParts count="${list.length}">${list.map((_, i) => `<tablePart r:id="rId${i + 1}"/>`).join('')}</tableParts></worksheet>`;
  const parts = {};
  list.forEach((t, i) => { parts[`xl/tables/table${i + 1}.xml`] = tablePart(t, i + 1); });
  write('tables.xlsx', {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/styles.xml', CT('styles')], ['xl/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'],
      ...list.map((_, i) => [`xl/tables/table${i + 1}.xml`, CT('table')]), ...PROPS_CT]),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><sheets><sheet name="Tables" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), 'worksheets/sheet1.xml'], ['rId2', RT('styles'), 'styles.xml'], ['rId3', RT('theme'), 'theme/theme1.xml']]),
    'xl/worksheets/sheet1.xml': sheet,
    'xl/worksheets/_rels/sheet1.xml.rels': rels(list.map((_, i) => [`rId${i + 1}`, RT('table'), `../tables/table${i + 1}.xml`])),
    'xl/styles.xml': styles(),
    'xl/theme/theme1.xml': THEME,
    ...parts,
  });
}
