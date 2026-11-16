// A workbook for what moves when rows, columns and cells are inserted, deleted or sorted: the sheet the Excel probe
// builds (tools/excel-compare/refmove-cases.mjs) written as Excel saved it, made-up data:
// refmoves.xlsx  sparklines (a group over O2:O5, a column sparkline at O7, one at Q2 of another sheet's data),
//                data validation rules (two lists, a whole number, a custom formula; the list reading the Data sheet
//                in the Excel 2010 form Excel writes for it), notes at B12 and F20 and a threaded comment at C13
//                (test/refmoves.test.js against test/refmoves-excel.json).
import { write, contentTypes, rels, RT, CT, DECL, X, R, ROOT_RELS, docProps, PROPS_CT, STYLES, esc, col } from './fixture-common.mjs';

const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const XM = 'http://schemas.microsoft.com/office/excel/2006/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const VML_CT = '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>';
const cell = (ref, v, s = 0) => (typeof v === 'string' ? `<c r="${ref}"${s ? ` s="${s}"` : ''} t="inlineStr"><is><t>${esc(v)}</t></is></c>` : `<c r="${ref}"${s ? ` s="${s}"` : ''}><v>${v}</v></c>`);
/** sheetData from { A1: value } (rows in order, cells in order). */
function sheetData(values, styleOf = () => 0) {
  const rows = new Map();
  for (const [ref, v] of Object.entries(values)) {
    const m = /^([A-Z]+)(\d+)$/.exec(ref);
    const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
    if (!rows.has(Number(m[2]))) rows.set(Number(m[2]), []);
    rows.get(Number(m[2])).push([c, cell(ref, v, styleOf(ref))]);
  }
  return [...rows].sort((a, b) => a[0] - b[0]).map(([r, list]) => `<row r="${r}">${list.sort((a, b) => a[0] - b[0]).map((x) => x[1]).join('')}</row>`).join('');
}

/** The probe's numbers and letters (refmove-cases.mjs). */
function refmoveValues() {
  const values = {};
  for (let r = 2; r <= 7; r++) {
    if (r === 6) continue;
    for (let c = 2; c <= 13; c++) values[`${col(c - 1)}${r}`] = ((r * 7 + c * 3) % 11) - 3;
  }
  const column = (c, list) => list.forEach((v, i) => { values[`${c}${10 + i}`] = v; });
  column('B', ['a', 'b', 'x', 'c', 'a', 'b']);
  column('C', [3, 1, 6, 2, 5, 4]);
  column('D', [4, 0, 7, 3, 1, 5]);
  column('E', ['p', 'q', 'zz', 'r', 'p', 'q']);
  column('H', ['a', 'b', 'c']);
  values.I10 = 4;
  return values;
}

export function refmoves() {
  const colors = '<x14:colorSeries rgb="FF376092"/><x14:colorNegative rgb="FFD00000"/><x14:colorAxis rgb="FF000000"/><x14:colorMarkers rgb="FFD00000"/>' +
    '<x14:colorFirst rgb="FFD00000"/><x14:colorLast rgb="FFD00000"/><x14:colorHigh rgb="FFD00000"/><x14:colorLow rgb="FFD00000"/>';
  const spark = (f, at) => `<x14:sparkline><xm:f>${f}</xm:f><xm:sqref>${at}</xm:sqref></x14:sparkline>`;
  const group = (attrs, list) => `<x14:sparklineGroup${attrs} displayEmptyCellsAs="gap">${colors}<x14:sparklines>${list.join('')}</x14:sparklines></x14:sparklineGroup>`;
  // (Groups newest first, as Excel writes them.)
  const groups = group('', [spark('Data!B2:M2', 'Q2')]) + group(' type="column"', [spark('Sheet1!B7:M7', 'O7')]) +
    group('', [2, 3, 4, 5].map((r) => spark(`Sheet1!B${r}:M${r}`, `O${r}`)));
  const rule = (type, sqref, f1, f2) => `<dataValidation type="${type}" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${sqref}"><formula1>${esc(f1)}</formula1>${f2 ? `<formula2>${esc(f2)}</formula2>` : ''}</dataValidation>`;
  const rules = `<dataValidations count="3">${rule('list', 'B10:B15', '$H$10:$H$12')}${rule('whole', 'C10:C15', '1', '$I$10')}${rule('custom', 'D10:D15', 'D10>C10')}</dataValidations>`;
  const ext = `<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="${X14}"><x14:dataValidations count="1" xmlns:xm="${XM}">` +
    '<x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"><x14:formula1><xm:f>Data!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>E10:E15</xm:sqref></x14:dataValidation></x14:dataValidations></ext>' +
    `<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="${X14}"><x14:sparklineGroups xmlns:xm="${XM}">${groups}</x14:sparklineGroups></ext></extLst>`;
  // Notes as Excel 365 writes them: B12 and F20, and C13's threaded comment with its placeholder note.
  const notesAt = [[11, 1], [19, 5], [12, 2]];
  const VML_NS = 'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"';
  const box = ([r, c], n) => `<v:shape id="_x0000_s${1025 + n}" type="#_x0000_t202" style='position:absolute;margin-left:${60 + c * 48}pt;margin-top:${(r - 1) * 15 + 7.5}pt;width:108pt;height:59.25pt;z-index:${n + 1};visibility:hidden' fillcolor="infoBackground [80]" strokecolor="none [81]" o:insetmode="auto">` +
    '<v:fill color2="infoBackground [80]"/><v:shadow color="none [81]" obscured="t"/><v:path o:connecttype="none"/><v:textbox style=\'mso-direction-alt:auto\'><div style=\'text-align:left\'></div></v:textbox>' +
    `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${c + 1}, 15, ${r - 1}, 10, ${c + 3}, 31, ${r + 3}, 9</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${r}</x:Row><x:Column>${c}</x:Column></x:ClientData></v:shape>`;
  const vml = `<xml ${VML_NS}><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>${notesAt.map(box).join('')}</xml>`;
  const TC = '{3C1D2A4B-5E6F-4A7B-8C9D-0E1F2A3B4C5D}';
  const PERSON = '{7A8B9C0D-1E2F-4A3B-9C4D-5E6F7A8B9C0D}';
  const comments = `${DECL}<comments xmlns="${X}"><authors><author>Robin Hale</author><author>tc=${TC}</author></authors><commentList>` +
    '<comment ref="B12" authorId="0" shapeId="0"><text><t>note B12</t></text></comment>' +
    `<comment ref="C13" authorId="1" shapeId="0"><text><t>[Threaded comment]\n\nComment:\n    thread C13</t></text></comment>` +
    '<comment ref="F20" authorId="0" shapeId="0"><text><t>note F20</t></text></comment></commentList></comments>';
  const TN = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';
  const threaded = `${DECL}<ThreadedComments xmlns="${TN}" xmlns:x="${X}"><threadedComment ref="C13" dT="2026-01-01T09:30:00.00" personId="${PERSON}" id="${TC}"><text>thread C13</text></threadedComment></ThreadedComments>`;
  const persons = `${DECL}<personList xmlns="${TN}" xmlns:x="${X}"><person displayName="Robin Hale" id="${PERSON}" userId="robin@example.com" providerId="None"/></personList>`;
  const data = { A1: 'p', A2: 'q', A3: 'r', ...Object.fromEntries(Array.from({ length: 12 }, (_, i) => [`${col(i + 1)}2`, (i * 5) % 7])) };
  write('refmoves.xlsx', {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/worksheets/sheet2.xml', CT('worksheet')], ['xl/styles.xml', CT('styles')],
      ['xl/comments1.xml', CT('comments')], ['xl/threadedComments/threadedComment1.xml', 'application/vnd.ms-excel.threadedcomments+xml'], ['xl/persons/person.xml', 'application/vnd.ms-excel.person+xml'], ...PROPS_CT])
      .replace('<Default Extension="xml"', `${VML_CT}<Default Extension="xml"`),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), 'worksheets/sheet1.xml'], ['rId2', RT('worksheet'), 'worksheets/sheet2.xml'], ['rId3', RT('styles'), 'styles.xml'], ['rId4', 'http://schemas.microsoft.com/office/2017/10/relationships/person', 'persons/person.xml']]),
    'xl/worksheets/sheet1.xml': `${DECL}<worksheet xmlns="${X}" xmlns:r="${R}" xmlns:mc="${MC}" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><dimension ref="B2:M15"/><sheetData>${sheetData(refmoveValues())}</sheetData>${rules}` +
      `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><legacyDrawing r:id="rId1"/>${ext}</worksheet>`,
    'xl/worksheets/_rels/sheet1.xml.rels': rels([['rId1', RT('vmlDrawing'), '../drawings/vmlDrawing1.vml'], ['rId2', RT('comments'), '../comments1.xml'], ['rId3', 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment', '../threadedComments/threadedComment1.xml']]),
    'xl/worksheets/sheet2.xml': `${DECL}<worksheet xmlns="${X}" xmlns:r="${R}"><dimension ref="A1:M3"/><sheetData>${sheetData(data)}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`,
    'xl/drawings/vmlDrawing1.vml': vml,
    'xl/comments1.xml': comments,
    'xl/threadedComments/threadedComment1.xml': threaded,
    'xl/persons/person.xml': persons,
    'xl/styles.xml': STYLES,
  });
}

