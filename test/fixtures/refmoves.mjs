// Two workbooks for what moves when rows, columns and cells are inserted, deleted or sorted, each the sheet the Excel
// probe builds (tools/excel-compare/refmove-cases.mjs, slicer-cases.mjs) written as Excel saved it, made-up data:
// refmoves.xlsx  sparklines (a group over O2:O5, a column sparkline at O7, one at Q2 of another sheet's data),
//                data validation rules (two lists, a whole number, a custom formula; the list reading the Data sheet
//                in the Excel 2010 form Excel writes for it), notes at B12 and F20 and a threaded comment at C13
//                (test/refmoves.test.js against test/refmoves-excel.json).
// slicers.xlsx   a table Sales on B2:E14 with a slicer on its Region column, a PivotTable of it at L20 with a slicer on
//                Item and a timeline on Date (test/slicers.test.js against test/slicers-excel.json).
import { write, contentTypes, rels, RT, CT, DECL, X, R, ROOT_RELS, docProps, PROPS_CT, STYLES, esc, col } from './fixture-common.mjs';

const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const X15 = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main';
const XM = 'http://schemas.microsoft.com/office/excel/2006/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
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

/** The probe's sales rows (slicer-cases.mjs): dates as serial numbers, a fortnight apart from 2026-01-05. */
const REGIONS = ['North', 'South', 'East', 'West'];
const ITEMS = ['Bolt', 'Nut', 'Pin'];
const SALES = Array.from({ length: 12 }, (_, i) => [46027 + i * 14, REGIONS[i % 4], ITEMS[i % 3], (i * 7) % 11 + 1]);
const isoDate = (serial) => new Date(Date.UTC(1899, 11, 30) + serial * 86400000).toISOString().slice(0, 19);

export function slicerBook() {
  const styles = STYLES.replace('<cellXfs count="2">', '<cellXfs count="3">').replace('</cellXfs>', '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>');
  const values = { B2: 'Date', C2: 'Region', D2: 'Item', E2: 'Units' };
  SALES.forEach((row, i) => row.forEach((v, j) => { values[`${col(j + 1)}${i + 3}`] = v; }));
  // The PivotTable at L20 as Excel lays it out: regions sorted, the sum of units, a grand total.
  const order = [2, 0, 1, 3]; // (East, North, South, West: the regions' cache items in sorted order)
  const sum = (k) => SALES.filter((s) => s[1] === REGIONS[k]).reduce((a, s) => a + s[3], 0);
  values.L20 = 'Row Labels'; values.M20 = 'Sum of Units';
  order.forEach((k, i) => { values[`L${21 + i}`] = REGIONS[k]; values[`M${21 + i}`] = sum(k); });
  values.L25 = 'Grand Total'; values.M25 = SALES.reduce((a, s) => a + s[3], 0);
  const PIVOT_ID = '395703160';
  const cacheFields = '<cacheFields count="4">' +
    `<cacheField name="Date" numFmtId="14"><sharedItems containsSemiMixedTypes="0" containsNonDate="0" containsDate="1" containsString="0" minDate="${isoDate(SALES[0][0])}" maxDate="${isoDate(SALES[11][0] + 1)}" count="12">${SALES.map((s) => `<d v="${isoDate(s[0])}"/>`).join('')}</sharedItems></cacheField>` +
    `<cacheField name="Region" numFmtId="0"><sharedItems count="4">${REGIONS.map((s) => `<s v="${s}"/>`).join('')}</sharedItems></cacheField>` +
    `<cacheField name="Item" numFmtId="0"><sharedItems count="3">${ITEMS.map((s) => `<s v="${s}"/>`).join('')}</sharedItems></cacheField>` +
    '<cacheField name="Units" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="1" maxValue="11"/></cacheField></cacheFields>';
  const cache = `${DECL}<pivotCacheDefinition xmlns="${X}" xmlns:r="${R}" r:id="rId1" refreshedBy="Robin Hale" refreshedDate="46342.0" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="12">` +
    `<cacheSource type="worksheet"><worksheetSource name="Sales"/></cacheSource>${cacheFields}<extLst><ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}" xmlns:x14="${X14}"><x14:pivotCacheDefinition pivotCacheId="${PIVOT_ID}"/></ext></extLst></pivotCacheDefinition>`;
  const records = `${DECL}<pivotCacheRecords xmlns="${X}" xmlns:r="${R}" count="12">${SALES.map((s, i) => `<r><x v="${i}"/><x v="${i % 4}"/><x v="${i % 3}"/><n v="${s[3]}"/></r>`).join('')}</pivotCacheRecords>`;
  const items = (n, sorted = [...Array(n).keys()]) => `<items count="${n + 1}">${sorted.map((x) => `<item x="${x}"/>`).join('')}<item t="default"/></items>`;
  const pivot = `${DECL}<pivotTableDefinition xmlns="${X}" name="Pivot1" cacheId="5" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="8" minRefreshableVersion="5" useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">` +
    '<location ref="L20:M25" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="4">' +
    `<pivotField numFmtId="14" showAll="0">${items(12)}</pivotField><pivotField axis="axisRow" showAll="0">${items(4, order)}</pivotField><pivotField showAll="0">${items(3)}</pivotField><pivotField dataField="1" showAll="0"/></pivotFields>` +
    '<rowFields count="1"><field x="1"/></rowFields><rowItems count="5"><i><x/></i><i><x v="1"/></i><i><x v="2"/></i><i><x v="3"/></i><i t="grand"><x/></i></rowItems><colItems count="1"><i/></colItems>' +
    '<dataFields count="1"><dataField name="Sum of Units" fld="3" baseField="0" baseItem="0"/></dataFields><pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>';
  const table = `${DECL}<table xmlns="${X}" id="1" name="Sales" displayName="Sales" ref="B2:E14" totalsRowShown="0"><autoFilter ref="B2:E14"/><tableColumns count="4"><tableColumn id="1" name="Date" dataDxfId="0"/><tableColumn id="2" name="Region"/><tableColumn id="3" name="Item"/><tableColumn id="4" name="Units"/></tableColumns>` +
    '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>';
  const SL = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
  const slicerCache = (name, source, body) => `${DECL}<slicerCacheDefinition xmlns="${SL}" xmlns:mc="${MC}" mc:Ignorable="x" xmlns:x="${X}" name="${name}" sourceName="${source}">${body}</slicerCacheDefinition>`;
  const itemCache = slicerCache('Slicer_Item', 'Item', `<pivotTables><pivotTable tabId="1" name="Pivot1"/></pivotTables><data><tabular pivotCacheId="${PIVOT_ID}"><items count="3"><i x="0" s="1"/><i x="1" s="1"/><i x="2" s="1"/></items></tabular></data>`);
  const regionCache = slicerCache('Slicer_Region', 'Region', `<extLst><x:ext uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}" xmlns:x15="${X15}"><x15:tableSlicerCache tableId="1" column="2"/></x:ext></extLst>`);
  const slicers = (name) => `${DECL}<slicers xmlns="${SL}" xmlns:mc="${MC}" mc:Ignorable="x" xmlns:x="${X}"><slicer name="${name}" cache="Slicer_${name}" caption="${name}" rowHeight="257175"/></slicers>`;
  const timelineCache = `${DECL}<timelineCacheDefinition xmlns="${X15}" name="NativeTimeline_Date" sourceName="Date"><pivotTables><pivotTable tabId="1" name="Pivot1"/></pivotTables><state minimalRefreshVersion="6" lastRefreshVersion="6" pivotCacheId="${PIVOT_ID}" filterType="unknown"><bounds startDate="2026-01-01T00:00:00" endDate="2027-01-01T00:00:00"/></state></timelineCacheDefinition>`;
  const timeline = `${DECL}<timelines xmlns="${X15}" xmlns:mc="${MC}" mc:Ignorable="x" xmlns:x="${X}"><timeline name="Date" cache="NativeTimeline_Date" caption="Date" level="2" selectionLevel="2" scrollPosition="2026-01-01T00:00:00"/></timelines>`;
  // The drawing: the table slicer kept in place (absolute, as Excel made it over COM), the others moving with cells.
  const at = (c0, r0, c1, r1) => `<xdr:from><xdr:col>${c0}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r0}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${c1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>`;
  const frame = (id, name, requires, uri, el) => `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice ${requires}><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${uri}">${el}</a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice></mc:AlternateContent>`;
  const SLE = 'http://schemas.microsoft.com/office/drawing/2010/slicer';
  const TSLE = 'http://schemas.microsoft.com/office/drawing/2012/timeslicer';
  const drawing = `${DECL}<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}">` +
    `<xdr:twoCellAnchor editAs="absolute">${at(6, 1, 9, 11)}${frame(2, 'Region', 'xmlns:sle15="http://schemas.microsoft.com/office/drawing/2012/slicer" Requires="sle15"', SLE, `<sle:slicer xmlns:sle="${SLE}" name="Region"/>`)}<xdr:clientData/></xdr:twoCellAnchor>` +
    `<xdr:twoCellAnchor editAs="oneCell">${at(15, 1, 18, 11)}${frame(3, 'Item', 'xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14"', SLE, `<sle:slicer xmlns:sle="${SLE}" name="Item"/>`)}<xdr:clientData/></xdr:twoCellAnchor>` +
    `<xdr:twoCellAnchor editAs="oneCell">${at(11, 1, 14, 9)}${frame(4, 'Date', `xmlns:tsle="${TSLE}" Requires="tsle"`, TSLE, `<tsle:timeslicer xmlns:tsle="${TSLE}" name="Date"/>`)}<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;
  const sheetExt = `<extLst><ext uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}" xmlns:x14="${X14}"><x14:slicerList><x14:slicer r:id="rId4"/></x14:slicerList></ext>` +
    `<ext uri="{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}" xmlns:x15="${X15}"><x14:slicerList xmlns:x14="${X14}"><x14:slicer r:id="rId5"/></x14:slicerList></ext>` +
    `<ext uri="{7E03D99C-DC04-49d9-9315-930204A7B6E9}" xmlns:x15="${X15}"><x15:timelineRefs><x15:timelineRef r:id="rId6"/></x15:timelineRefs></ext></extLst>`;
  const bookExt = `<extLst><ext uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}" xmlns:x14="${X14}"><x14:slicerCaches><x14:slicerCache r:id="rId3"/></x14:slicerCaches></ext>` +
    `<ext uri="{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}" xmlns:x15="${X15}"><x15:timelineCacheRefs><x15:timelineCacheRef r:id="rId4"/></x15:timelineCacheRefs></ext>` +
    `<ext uri="{46BE6895-7355-4a93-B00E-2C351335B9C9}" xmlns:x15="${X15}"><x15:slicerCaches xmlns:x14="${X14}"><x14:slicerCache r:id="rId5"/></x15:slicerCaches></ext></extLst>`;
  const ms = (t) => `application/vnd.ms-excel.${t}+xml`;
  write('slicers.xlsx', {
    '[Content_Types].xml': contentTypes([['xl/workbook.xml', CT('sheet.main')], ['xl/worksheets/sheet1.xml', CT('worksheet')], ['xl/styles.xml', CT('styles')], ['xl/tables/table1.xml', CT('table')],
      ['xl/pivotTables/pivotTable1.xml', CT('pivotTable')], ['xl/pivotCache/pivotCacheDefinition1.xml', CT('pivotCacheDefinition')], ['xl/pivotCache/pivotCacheRecords1.xml', CT('pivotCacheRecords')],
      ['xl/slicerCaches/slicerCache1.xml', ms('slicerCache')], ['xl/slicerCaches/slicerCache2.xml', ms('slicerCache')], ['xl/slicers/slicer1.xml', ms('slicer')], ['xl/slicers/slicer2.xml', ms('slicer')],
      ['xl/timelineCaches/timelineCache1.xml', ms('timelineCache')], ['xl/timelines/timeline1.xml', ms('timeline')], ['xl/drawings/drawing1.xml', 'application/vnd.openxmlformats-officedocument.drawing+xml'], ...PROPS_CT]),
    '_rels/.rels': ROOT_RELS,
    ...docProps(),
    'xl/workbook.xml': `${DECL}<workbook xmlns="${X}" xmlns:r="${R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>` +
      '<definedNames><definedName name="NativeTimeline_Date">#N/A</definedName><definedName name="Slicer_Item">#N/A</definedName><definedName name="Slicer_Region">#N/A</definedName></definedNames>' +
      `<calcPr calcId="191029"/><pivotCaches><pivotCache cacheId="5" r:id="rId2"/></pivotCaches>${bookExt}</workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', RT('worksheet'), 'worksheets/sheet1.xml'], ['rId2', RT('pivotCacheDefinition'), 'pivotCache/pivotCacheDefinition1.xml'],
      ['rId3', 'http://schemas.microsoft.com/office/2007/relationships/slicerCache', 'slicerCaches/slicerCache1.xml'], ['rId4', 'http://schemas.microsoft.com/office/2011/relationships/timelineCache', 'timelineCaches/timelineCache1.xml'],
      ['rId5', 'http://schemas.microsoft.com/office/2007/relationships/slicerCache', 'slicerCaches/slicerCache2.xml'], ['rId6', RT('styles'), 'styles.xml']]),
    'xl/worksheets/sheet1.xml': `${DECL}<worksheet xmlns="${X}" xmlns:r="${R}"><dimension ref="B2:M25"/><sheetFormatPr defaultRowHeight="15"/>` +
      '<cols><col min="2" max="2" width="10.42578125" bestFit="1" customWidth="1"/><col min="3" max="3" width="9.42578125" customWidth="1"/><col min="12" max="12" width="13.42578125" bestFit="1" customWidth="1"/><col min="13" max="13" width="12.42578125" bestFit="1" customWidth="1"/></cols>' +
      `<sheetData>${sheetData(values, (ref) => (/^B(\d+)$/.test(ref) && ref !== 'B2' ? 2 : 0))}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><drawing r:id="rId2"/><tableParts count="1"><tablePart r:id="rId3"/></tableParts>${sheetExt}</worksheet>`,
    'xl/worksheets/_rels/sheet1.xml.rels': rels([['rId1', RT('pivotTable'), '../pivotTables/pivotTable1.xml'], ['rId2', RT('drawing'), '../drawings/drawing1.xml'], ['rId3', RT('table'), '../tables/table1.xml'],
      ['rId4', 'http://schemas.microsoft.com/office/2007/relationships/slicer', '../slicers/slicer1.xml'], ['rId5', 'http://schemas.microsoft.com/office/2007/relationships/slicer', '../slicers/slicer2.xml'],
      ['rId6', 'http://schemas.microsoft.com/office/2011/relationships/timeline', '../timelines/timeline1.xml']]),
    'xl/tables/table1.xml': table,
    'xl/pivotTables/pivotTable1.xml': pivot,
    'xl/pivotTables/_rels/pivotTable1.xml.rels': rels([['rId1', RT('pivotCacheDefinition'), '../pivotCache/pivotCacheDefinition1.xml']]),
    'xl/pivotCache/pivotCacheDefinition1.xml': cache,
    'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': rels([['rId1', RT('pivotCacheRecords'), 'pivotCacheRecords1.xml']]),
    'xl/pivotCache/pivotCacheRecords1.xml': records,
    'xl/slicerCaches/slicerCache1.xml': itemCache,
    'xl/slicerCaches/slicerCache2.xml': regionCache,
    'xl/slicers/slicer1.xml': slicers('Item'),
    'xl/slicers/slicer2.xml': slicers('Region'),
    'xl/timelineCaches/timelineCache1.xml': timelineCache,
    'xl/timelines/timeline1.xml': timeline,
    'xl/drawings/drawing1.xml': drawing,
    'xl/styles.xml': styles,
  });
}
