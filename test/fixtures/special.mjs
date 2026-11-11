// special.xlsx: a sheet with something for each of Go To Special's choices (Home > Find & Select > Go To Special),
// for test/special.test.js and tools/excel-compare/excel-special.ps1, which asks Excel's SpecialCells the same.
//   A1:E1 headings; A2:A6 text, B2:B6 numbers; C2:C6 formulas giving a number, a number, #DIV/0!, TRUE and "" (text);
//   D2:D5 constants TRUE, #N/A, a date and text; E2 =SUM(B2:B6), E3 =E2+C2 (precedents and dependents);
//   G2:G4 an array formula {=B2:B4*10}; notes on A2 and D5; row 4 and column F hidden; A9:B9 merged (text in A9);
//   H10 formatted only (bold), so the used range runs to it; a whole-number validation on B2:B6 and a list on D2:D3;
//   a conditional format on C2:C6 and another on A2:A3; A12:C16 a block for row and column differences, A17:C17 text
//   differing in case, A20:C21 formulas alike relative to their cells (A21, B21) or not (C21); J12:K13 a block of its
//   own (a current region apart).
// Made-up data; the note's author is made up.
export function special(h) {
  const { write, contentTypes, rels, RT, CT, DECL, X, R, ROOT_RELS, docProps, PROPS_CT, STYLES, esc } = h;
  const n = (ref, v) => `<c r="${ref}"><v>${v}</v></c>`;
  const t = (ref, s) => `<c r="${ref}" t="inlineStr"><is><t>${esc(s)}</t></is></c>`;
  const b = (ref, v) => `<c r="${ref}" t="b"><v>${v ? 1 : 0}</v></c>`;
  const e = (ref, v) => `<c r="${ref}" t="e"><v>${v}</v></c>`;
  // (A formula with its cached result: a number, or { t, v } for another type.)
  const f = (ref, formula, v) => (typeof v === 'object' ? `<c r="${ref}" t="${v.t}"><f>${esc(formula)}</f><v>${esc(v.v)}</v></c>` : `<c r="${ref}"><f>${esc(formula)}</f><v>${v}</v></c>`);
  const rows = {
    1: [t('A1', 'Item'), t('B1', 'Qty'), t('C1', 'Calc'), t('D1', 'Misc'), t('E1', 'Sums')],
    2: [t('A2', 'apples'), n('B2', 4), f('C2', 'B2*2', 8), b('D2', true), f('E2', 'SUM(B2:B6)', 20), '<c r="G2"><f t="array" ref="G2:G4">B2:B4*10</f><v>40</v></c>'],
    3: [t('A3', 'pears'), n('B3', 3), f('C3', 'B3*2', 6), e('D3', '#N/A'), f('E3', 'E2+C2', 28), n('G3', 30)],
    4: [t('A4', 'figs'), n('B4', 5), f('C4', 'B4/0', { t: 'e', v: '#DIV/0!' }), n('D4', 45292), n('G4', 50)],
    5: [t('A5', 'plums'), n('B5', 6), f('C5', 'B5>2', { t: 'b', v: '1' }), t('D5', 'later')],
    6: [t('A6', 'limes'), n('B6', 2), f('C6', '""', { t: 'str', v: '' })],
    9: [t('A9', 'merged')],
    10: ['<c r="H10" s="1"/>'],
    12: [n('A12', 1), n('B12', 1), n('C12', 2), n('J12', 7), n('K12', 8)],
    13: [t('A13', 'x'), t('B13', 'x'), t('C13', 'y'), n('J13', 9)],
    14: [n('A14', 5), f('B14', 'A14', 5), n('C14', 5)],
    15: [f('A15', '1+1', 2), f('B15', '1+1', 2), n('C15', 2)],
    16: [n('A16', 3), n('B16', 4), n('C16', 3)],
    17: [t('A17', 'Abc'), t('B17', 'abc'), t('C17', 'ABD')],
    20: [n('A20', 1), n('B20', 2), n('C20', 3)],
    21: [f('A21', 'A20*2', 2), f('B21', 'B20*2', 4), f('C21', 'A20*2', 2)],
  };
  const sheetData = Object.entries(rows).map(([r, cells]) => `<row r="${r}"${r === '4' ? ' hidden="1"' : ''}>${cells.join('')}</row>`).join('');
  const sheet = `${DECL}<worksheet xmlns="${X}" xmlns:r="${R}"><sheetFormatPr defaultRowHeight="15"/><cols><col min="6" max="6" width="9.140625" hidden="1"/></cols>` +
    `<sheetData>${sheetData}</sheetData><mergeCells count="1"><mergeCell ref="A9:B9"/></mergeCells>` +
    '<conditionalFormatting sqref="C2:C6"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>' +
    '<conditionalFormatting sqref="A2:A3"><cfRule type="containsText" dxfId="0" priority="2" operator="containsText" text="p"><formula>NOT(ISERROR(SEARCH("p",A2)))</formula></cfRule></conditionalFormatting>' +
    '<dataValidations count="2"><dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B2:B6"><formula1>0</formula1><formula2>100</formula2></dataValidation>' +
    '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="D2:D3"><formula1>"yes,no"</formula1></dataValidation></dataValidations>' +
    '<legacyDrawing r:id="rId1"/></worksheet>';
  const VML_NS = 'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"';
  const shape = (k, r, c) => `<v:shape id="_x0000_s${1024 + k}" type="#_x0000_t202" style='position:absolute;margin-left:${60 + c * 48}pt;margin-top:${r * 15}pt;width:108pt;height:59.25pt;z-index:${k};visibility:hidden' fillcolor="infoBackground [80]" strokecolor="none [81]" o:insetmode="auto">` +
    `<v:fill color2="infoBackground [80]"/><v:shadow color="none [81]" obscured="t"/><v:path o:connecttype="none"/><v:textbox style='mso-direction-alt:auto'><div style='text-align:left'></div></v:textbox>` +
    `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${c + 1}, 15, ${r}, 10, ${c + 3}, 31, ${r + 4}, 9</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${r}</x:Row><x:Column>${c}</x:Column></x:ClientData></v:shape>`;
  const vml = `<xml ${VML_NS}><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>` +
    shape(1, 1, 0) + shape(2, 4, 3) + '</xml>';
  const note = (ref, text) => `<comment ref="${ref}" authorId="0" shapeId="0"><text><r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr><t xml:space="preserve">${esc(text)}</t></r></text></comment>`;
  const comments = `${DECL}<comments xmlns="${X}"><authors><author>Robin Park</author></authors><commentList>${note('A2', 'Fresh in')}${note('D5', 'Check later')}</commentList></comments>`;
  write('special.xlsx', {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/styles.xml', CT('styles')], ['xl/comments1.xml', CT('comments')], ...PROPS_CT])
      .replace('<Default Extension="xml"', '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Default Extension="xml"'),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><sheets><sheet name="Special" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), 'worksheets/sheet1.xml'], ['rId2', RT('styles'), 'styles.xml']]),
    'xl/worksheets/sheet1.xml': sheet,
    'xl/worksheets/_rels/sheet1.xml.rels': rels([['rId1', RT('vmlDrawing'), '../drawings/vmlDrawing1.vml'], ['rId2', RT('comments'), '../comments1.xml']]),
    'xl/drawings/vmlDrawing1.vml': vml,
    'xl/comments1.xml': comments,
    'xl/styles.xml': STYLES,
  });
}
