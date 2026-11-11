// Workbooks with drawing and comment parts as Excel writes them: notes, charts, anchors (see make-fixtures.mjs).
import { write, contentTypes, rels, RT, CT, DECL, X, R, ROOT_RELS, docProps, PROPS_CT, STYLES, esc, col } from './fixture-common.mjs';

// Notes and a threaded comment as Excel 365 writes them (its parts and XML read from a workbook Excel saved): a note's
// text in comments1.xml with its box in a VML drawing; a threaded comment in threadedComments with its people in
// persons, and a placeholder note for older versions. The people are made up.
export function notes() {
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
export function charts() {
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
export function anchors(table = false) {
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
