// A new chart's part as Excel writes it when Insert > Charts adds one in its default style (measured over COM:
// tools/excel-compare/excel-chart-insert.ps1, test/chartnew-excel.json): text for text, but for the axis ids and the
// series' unique ids, which Excel makes up afresh each time. Fourteen kinds: column and bar (clustered, stacked, 100%),
// line (with or without markers), pie, doughnut, area (standard, stacked) and scatter (markers, lines, smooth lines).

/** The colour Excel gives series and points: its theme's six accents, in turn. */
const accent = (i) => `accent${(i % 6) + 1}`;
const fill = (clr) => `<a:solidFill><a:schemeClr val="${clr}"/></a:solidFill>`;
const GREY = (mod, off) => `<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="${mod}"/><a:lumOff val="${off}"/></a:schemeClr></a:solidFill>`;
const LINE_15 = `<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">${GREY(15000, 85000)}<a:round/></a:ln>`;
const LINE_25 = `<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">${GREY(25000, 75000)}<a:round/></a:ln>`;
const SP_NONE = '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>';
const SP_GRID = `<c:spPr>${LINE_15}<a:effectLst/></c:spPr>`;
const SP_AXIS = `<c:spPr><a:noFill/>${LINE_15}<a:effectLst/></c:spPr>`;
const SP_AXIS_XY = `<c:spPr><a:noFill/>${LINE_25}<a:effectLst/></c:spPr>`;
const SP_SPACE = `<c:spPr>${fill('bg1')}${LINE_15}<a:effectLst/></c:spPr>`;
const DEF = (sz, extra = '') => `<a:defRPr sz="${sz}" b="0" i="0" u="none" strike="noStrike" kern="1200"${extra} baseline="0">${GREY(65000, 35000)}<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>`;
const BODY = (rot) => `<a:bodyPr rot="${rot}" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>`;
const TX = (rot, def) => `<c:txPr>${BODY(rot)}<a:lstStyle/><a:p><a:pPr>${def}</a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
const TX_TITLE = TX(0, DEF(1400, ' spc="0"'));
const TX_AXIS = TX(-60000000, DEF(900));
const TX_LEGEND = TX(0, DEF(900));
const TX_SPACE = '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>';
const DLBLS = (leader) => `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>${leader ? '<c:showLeaderLines val="1"/>' : ''}</c:dLbls>`;
const EXT_SER = (n, guid) => `<c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{${String(n).padStart(8, '0')}-${guid}}"/></c:ext></c:extLst>`;
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"><c:date1904 val="DATE1904"/><c:lang val="en-US"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent>';
const TAIL = `${SP_SPACE}${TX_SPACE}<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>`;

/** The kinds Insert > Charts offers, as Excel's chart type names them, and how each is written. */
export const CHART_KINDS = {
  columnClustered: { family: 'bar', dir: 'col', grouping: 'clustered' },
  columnStacked: { family: 'bar', dir: 'col', grouping: 'stacked' },
  columnStacked100: { family: 'bar', dir: 'col', grouping: 'percentStacked' },
  barClustered: { family: 'bar', dir: 'bar', grouping: 'clustered' },
  barStacked: { family: 'bar', dir: 'bar', grouping: 'stacked' },
  barStacked100: { family: 'bar', dir: 'bar', grouping: 'percentStacked' },
  line: { family: 'line', grouping: 'standard' },
  lineMarkers: { family: 'line', grouping: 'standard', markers: true },
  pie: { family: 'pie' },
  doughnut: { family: 'doughnut' },
  area: { family: 'area', grouping: 'standard' },
  areaStacked: { family: 'area', grouping: 'stacked' },
  scatter: { family: 'scatter', style: 'lineMarker' },
  scatterLines: { family: 'scatter', style: 'lineMarker', lines: true },
  scatterSmooth: { family: 'scatter', style: 'smoothMarker', lines: true, smooth: true },
};

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
/** A reference with its cached points: strRef for text, numRef (with the first value's format code) for numbers. */
function ref(tag, data) {
  if (!data) return '';
  const pts = data.values.map((v, i) => (v == null || v === '' ? '' : `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`)).join('');
  const body = data.numeric
    ? `<c:numRef><c:f>${esc(data.f)}</c:f><c:numCache><c:formatCode>${esc(data.format || 'General')}</c:formatCode><c:ptCount val="${data.values.length}"/>${pts}</c:numCache></c:numRef>`
    : `<c:strRef><c:f>${esc(data.f)}</c:f><c:strCache><c:ptCount val="${data.values.length}"/>${pts}</c:strCache></c:strRef>`;
  return `<c:${tag}>${body}</c:${tag}>`;
}
const marker = (i) => `<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>${fill(accent(i))}<a:ln w="9525">${fill(accent(i))}</a:ln><a:effectLst/></c:spPr></c:marker>`;

/** One series of `kind` (CHART_KINDS's), i its index: name, categories (or x values) and values. */
function seriesXml(k, s, i, guid) {
  const head = `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${s.name ? ref('tx', { ...s.name, numeric: false }) : ''}`;
  const tail = EXT_SER(i, guid);
  if (k.family === 'bar') return `${head}<c:spPr>${fill(accent(i))}<a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr><c:invertIfNegative val="0"/>${ref('cat', s.cat)}${ref('val', s.val)}${tail}</c:ser>`;
  if (k.family === 'area') return `${head}<c:spPr>${fill(accent(i))}<a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>${ref('cat', s.cat)}${ref('val', s.val)}${tail}</c:ser>`;
  if (k.family === 'line') {
    return `${head}<c:spPr><a:ln w="28575" cap="rnd">${fill(accent(i))}<a:round/></a:ln><a:effectLst/></c:spPr>${k.markers ? marker(i) : '<c:marker><c:symbol val="none"/></c:marker>'}${ref('cat', s.cat)}${ref('val', s.val)}<c:smooth val="0"/>${tail}</c:ser>`;
  }
  if (k.family === 'scatter') {
    const line = k.lines ? `<a:ln w="19050" cap="rnd">${fill(accent(i))}<a:round/></a:ln>` : '<a:ln w="38100" cap="rnd"><a:noFill/><a:round/></a:ln>';
    return `${head}<c:spPr>${line}<a:effectLst/></c:spPr>${marker(i)}${ref('xVal', s.cat)}${ref('yVal', s.val)}<c:smooth val="${k.smooth ? 1 : 0}"/>${tail}</c:ser>`;
  }
  // Pie and doughnut: each point its own accent, outlined in the background colour.
  const points = s.val.values.map((_, p) => `<c:dPt><c:idx val="${p}"/><c:bubble3D val="0"/><c:spPr>${fill(accent(p))}<a:ln w="19050">${fill('lt1')}</a:ln><a:effectLst/></c:spPr></c:dPt>`).join('');
  return `${head}${points}${ref('cat', s.cat)}${ref('val', s.val)}${tail}</c:ser>`;
}

const catAx = (id, cross, pos, tick) => `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="${tick}"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${SP_AXIS}${TX_AXIS}<c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
const valAx = (id, cross, pos, { format = 'General', between = 'between', sp = SP_NONE, numFmt = true } = {}) => `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/><c:majorGridlines>${SP_GRID}</c:majorGridlines>${numFmt ? `<c:numFmt formatCode="${format}" sourceLinked="1"/>` : ''}<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${sp}${TX_AXIS}<c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:crossBetween val="${between}"/></c:valAx>`;

/**
 * The chart part's text for `kind` and its `series` ([{ name, cat, val }], each a reference with its cached values:
 * { f, values, numeric, format }). `ids`: { axes: [two numbers], guid } (Excel's own are made up afresh: random here).
 */
export function chartXml(kind, series, { date1904 = false, ids }) {
  const k = CHART_KINDS[kind];
  const [a1, a2] = ids.axes;
  const sers = series.map((s, i) => seriesXml(k, s, i, ids.guid)).join('');
  const axIds = `<c:axId val="${a1}"/><c:axId val="${a2}"/>`;
  let plot;
  if (k.family === 'bar') {
    const gaps = k.grouping === 'clustered' ? (k.dir === 'col' ? '<c:gapWidth val="219"/><c:overlap val="-27"/>' : '<c:gapWidth val="182"/>') : '<c:gapWidth val="150"/><c:overlap val="100"/>';
    plot = `<c:barChart><c:barDir val="${k.dir}"/><c:grouping val="${k.grouping}"/><c:varyColors val="0"/>${sers}${DLBLS(false)}${gaps}${axIds}</c:barChart>`
      + catAx(a1, a2, k.dir === 'col' ? 'b' : 'l', 'none') + valAx(a2, a1, k.dir === 'col' ? 'l' : 'b', { format: k.grouping === 'percentStacked' ? '0%' : 'General' });
  } else if (k.family === 'line') {
    plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}${DLBLS(false)}${k.markers ? '<c:marker val="1"/>' : ''}<c:smooth val="0"/>${axIds}</c:lineChart>`
      + catAx(a1, a2, 'b', 'none') + valAx(a2, a1, 'l');
  } else if (k.family === 'area') {
    plot = `<c:areaChart><c:grouping val="${k.grouping}"/><c:varyColors val="0"/>${sers}${DLBLS(false)}${axIds}</c:areaChart>`
      + catAx(a1, a2, 'b', 'out') + valAx(a2, a1, 'l', { between: 'midCat' });
  } else if (k.family === 'scatter') {
    plot = `<c:scatterChart><c:scatterStyle val="${k.style}"/><c:varyColors val="0"/>${sers}${DLBLS(false)}${axIds}</c:scatterChart>`
      + valAx(a1, a2, 'b', { between: 'midCat', sp: SP_AXIS_XY, numFmt: false }) + valAx(a2, a1, 'l', { between: 'midCat', sp: SP_AXIS_XY });
  } else {
    plot = `<c:${k.family}Chart><c:varyColors val="1"/>${sers}${DLBLS(true)}<c:firstSliceAng val="0"/>${k.family === 'doughnut' ? '<c:holeSize val="75"/>' : ''}</c:${k.family}Chart>`;
  }
  const blanks = k.family === 'area' ? 'zero' : 'gap';
  const na = '<c:extLst><c:ext uri="{56B9EC1D-385E-4148-901F-78D8002777C0}" xmlns:c16r3="http://schemas.microsoft.com/office/drawing/2017/03/chart"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="1"/></c16r3:dataDisplayOptions16></c:ext></c:extLst>';
  return `${HEAD.replace('DATE1904', date1904 ? '1' : '0')}<c:chart><c:title><c:overlay val="0"/>${SP_NONE}${TX_TITLE}</c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>${plot}${SP_NONE}</c:plotArea>`
    + `<c:legend><c:legendPos val="b"/><c:overlay val="0"/>${SP_NONE}${TX_LEGEND}</c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="${blanks}"/><c:showDLblsOverMax val="0"/>${na}</c:chart>${TAIL}`;
}

/** A series' data labels showing its values, as Excel writes them (Add Chart Element > Data Labels > Show:
 *  excel-chart-edit.ps1). */
const DLBLS_SHOW = `<c:dLbls>${SP_NONE}<c:txPr><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" lIns="38100" tIns="19050" rIns="38100" bIns="19050" anchor="ctr" anchorCtr="1"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr>${DEF(900).replace(GREY(65000, 35000), GREY(75000, 25000))}</a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/><c:showLeaderLines val="0"/><c:extLst><c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"><c15:showLeaderLines val="1"/><c15:leaderLines><c:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">${GREY(35000, 65000)}<a:round/></a:ln><a:effectLst/></c:spPr></c15:leaderLines></c:ext></c:extLst></c:dLbls>`;
/** The pieces chartedit.js puts back into a chart as Excel writes them. */
export const chartLook = { SP_NONE, TX_TITLE, TX_LEGEND, BODY, DEF, DLBLS_SHOW };
