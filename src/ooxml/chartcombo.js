// Combo charts and the secondary axis, as Excel writes them (measured over COM on a clustered column chart of two
// series: tools/excel-compare/excel-chart-format.ps1, test/chartformat-excel.json): a series' chart type changed
// (Series.ChartType: line, line with markers, area, back to columns) and a series put on the secondary axis or back
// (Series.AxisGroup). Each series of a chart lives in a chart group of its kind on one pair of axes; a series moved is
// taken out of its group and put into the group of its new kind and axes, in its order, the group made when there is
// none and dropped when it is left empty. The secondary axes are made with the first series put there and dropped with
// the last one taken off (Excel's part is then as it was).
import { change } from './chartedit.js';
import { kidOf, kidsOf, textOf, valOf } from './charttext.js';
import { plotOf } from './chartformat.js';

const DLBLS = '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>';
/** A new group of each kind, as Excel made it (its series and axes put in where SERS and AXIDS are): a line group
 *  with markers on (whether its series show them or not), an area group bare; a bar group not measured (a new bar
 *  group copies the chart's own bar group's settings, as Excel's secondary one did), Insert's own gap and overlap. */
const NEW = {
  bar: `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>SERS${DLBLS}<c:gapWidth val="219"/><c:overlap val="-27"/>AXIDS</c:barChart>`,
  line: `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>SERS${DLBLS}<c:marker val="1"/><c:smooth val="0"/>AXIDS</c:lineChart>`,
  area: `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>SERS${DLBLS}AXIDS</c:areaChart>`,
};
const FAMILY = { 'c:barChart': 'bar', 'c:lineChart': 'line', 'c:areaChart': 'area' };
// (Excel wrote an area group before the bar group, a line group after it, and the secondary axes' groups after the
// primary ones.)
const RANK = { area: 0, bar: 1, line: 2 };
/** The series types Change Series Chart Type offers here, as Excel's Series.ChartType names them. */
export const SERIES_TYPES = { column: { family: 'bar' }, line: { family: 'line' }, lineMarkers: { family: 'line', markers: true }, area: { family: 'area' } };

/** The plot area's groups as pieces: [{ family, open, head, sers: [{ idx, text }], tail, axIds, ids }]. */
function groupsOf(x) {
  const p = plotOf(x);
  const groups = p.groups.map((g) => {
    const family = FAMILY[g.name];
    const kids = kidsOf(x, g);
    if (!family || (family === 'bar' && valOf(x, kidOf(x, g, 'c:barDir')) === 'bar')) throw new Error("This chart's series types can't be changed here yet.");
    const sers = kids.filter((k) => k.name === 'c:ser');
    const ax = kids.filter((k) => k.name === 'c:axId');
    const first = sers[0] || kids.find((k) => k.name === 'c:dLbls') || ax[0];
    const afterSers = sers.length ? sers[sers.length - 1].end : first.start;
    return {
      family, name: g.name, head: x.slice(g.inner, first.start),
      sers: sers.map((s) => ({ idx: Number(valOf(x, kidOf(x, s, 'c:idx'))), text: textOf(x, s) })),
      tail: x.slice(afterSers, ax.length ? ax[0].start : g.close), axIds: ax.map((a) => textOf(x, a)).join(''), ids: ax.map((a) => valOf(x, a)),
    };
  });
  return { p, groups };
}
const groupText = (g) => `<${g.name}>${g.head}${g.sers.map((s) => s.text).join('')}${g.tail}${g.axIds}</${g.name}>`;

/** A series' text for its new kind, as Excel rewrote it: its name, points' and labels' settings, data and unique id
 *  kept; its look the new kind's default, in the accent of its place (a series given its own red came out a line in
 *  accent 2, its own colour gone: comboOwnColor); a line's markers none or Excel's circle of 5. */
function retyped(ser, type) {
  const to = SERIES_TYPES[type];
  const idx = Number(/<c:idx val="(\d+)"/.exec(ser)[1]);
  const fill = `<a:solidFill><a:schemeClr val="accent${(idx % 6) + 1}"/></a:solidFill>`;
  const take = (tag) => new RegExp(`<c:${tag}>[\\s\\S]*?</c:${tag}>|<c:${tag}\\b[^>]*/>`).exec(ser)?.[0] || '';
  const all = (tag) => [...ser.matchAll(new RegExp(`<c:${tag}>[\\s\\S]*?</c:${tag}>`, 'g'))].map((m) => m[0]).join('');
  const head = `<c:ser>${take('idx')}${take('order')}${take('tx')}`;
  const body = `${all('dPt')}${take('dLbls')}${all('trendline')}${all('errBars')}${take('cat')}${take('val')}`;
  const ext = /<c:extLst>(?:(?!<c:extLst>)[\s\S])*?<\/c:extLst><\/c:ser>$/.exec(ser)?.[0].replace(/<\/c:ser>$/, '') || '';
  if (to.family === 'line') {
    const marker = to.markers ? `<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>${fill}<a:ln w="9525">${fill}</a:ln><a:effectLst/></c:spPr></c:marker>` : '<c:marker><c:symbol val="none"/></c:marker>';
    return `${head}<c:spPr><a:ln w="28575" cap="rnd">${fill}<a:round/></a:ln><a:effectLst/></c:spPr>${marker}${body}<c:smooth val="0"/>${ext}</c:ser>`;
  }
  const sp = `<c:spPr>${fill}<a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>`;
  return `${head}${sp}${to.family === 'bar' ? '<c:invertIfNegative val="0"/>' : ''}${body}${ext}</c:ser>`;
}

const freshId = () => String(Math.floor(1e8 + Math.random() * 2e9));

/** The secondary axes Excel made: a value axis on the right (no gridlines, ticks out, crossing at the maximum, the
 *  primary value axis's number format, text and crossing between) and a hidden category axis under it. */
function secondaryAxes(x, p, ids) {
  const [c2, v2] = ids;
  const val = p.axes.find((a) => p.primary.includes(a.id) && a.kind === 'valAx');
  const valText = val ? textOf(x, val.el) : '';
  const pick = (re, d) => re.exec(valText)?.[0] || d;
  const numFmt = pick(/<c:numFmt\b[^>]*\/>/, '<c:numFmt formatCode="General" sourceLinked="1"/>');
  const txPr = pick(/<c:txPr>[\s\S]*?<\/c:txPr>/, '');
  const between = pick(/<c:crossBetween val="\w+"\/>/, '<c:crossBetween val="between"/>');
  const head = (id, del, pos) => `<c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${del}"/><c:axPos val="${pos}"/>`;
  return `<c:valAx>${head(v2, 0, 'r')}${numFmt}<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>${txPr}<c:crossAx val="${c2}"/><c:crosses val="max"/>${between}</c:valAx>`
    + `<c:catAx>${head(c2, 1, 'b')}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${v2}"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
}

/** Series `idx` moved to `type` (SERIES_TYPES) and to the secondary axes or the primary ones, the part rewritten from
 *  the first group to the last axis. */
function move(x, idx, type, secondary) {
  const { p, groups } = groupsOf(x);
  const from = groups.find((g) => g.sers.some((s) => s.idx === idx));
  if (!from) throw new Error('No such series');
  const fromSecondary = !!p.secondary && from.ids.join() === p.secondary.join();
  const family = type ? SERIES_TYPES[type].family : from.family;
  const onSecondary = secondary ?? fromSecondary;
  let ser = from.sers.find((s) => s.idx === idx).text;
  const markers = /<c:marker><c:symbol val="(?!none)/.test(ser);
  if (type && (family !== from.family || (family === 'line' && markers !== !!SERIES_TYPES[type].markers))) ser = retyped(ser, type);
  if (family === from.family && onSecondary === fromSecondary && ser === from.sers.find((s) => s.idx === idx).text) return x;
  from.sers = from.sers.filter((s) => s.idx !== idx);
  let ids = onSecondary ? p.secondary : p.primary;
  let axesText = null;
  if (onSecondary && !ids) { ids = [freshId(), freshId()]; axesText = secondaryAxes(x, p, ids); }
  let to = groups.find((g) => g.family === family && g.ids.join() === ids.join());
  if (!to) {
    // (A new group: a copy of the series' own group's settings when only the axes change, else the kind's own.)
    const axIds = ids.map((id) => `<c:axId val="${id}"/>`).join('');
    const m = /^<(c:\w+)>([\s\S]*?)SERS([\s\S]*?)AXIDS<\/c:\w+>$/.exec(NEW[family]);
    to = family === from.family ? { ...from, sers: [], axIds, ids } : { family, name: m[1], head: m[2], sers: [], tail: m[3], axIds, ids };
    const rank = (g) => (g.ids.join() === p.primary.join() ? 0 : 10) + RANK[g.family];
    const at = groups.findIndex((g) => rank(g) > rank(to));
    groups.splice(at < 0 ? groups.length : at, 0, to);
  } else if (to.family === 'bar' && from.family !== 'bar') {
    // (A series of another kind joining a bar group: Excel dropped the group's overlap, keeping its gap width.)
    to.tail = to.tail.replace(/<c:overlap val="-?\d+"\/>/, '');
  }
  const at = to.sers.findIndex((s) => s.idx > idx);
  to.sers.splice(at < 0 ? to.sers.length : at, 0, { idx, text: ser });
  const kept = groups.filter((g) => g.sers.length);
  const usesSecondary = kept.some((g) => g.ids.join() === ids.join() && onSecondary) || (!!p.secondary && kept.some((g) => g.ids.join() === p.secondary.join()));
  let axes = p.axes.filter((a) => usesSecondary || !p.secondary || !p.secondary.includes(a.id)).map((a) => textOf(x, a.el)).join('');
  if (axesText) axes += axesText;
  const start = p.groups[0].start;
  const end = p.axes.length ? p.axes[p.axes.length - 1].el.end : p.groups[p.groups.length - 1].end;
  return x.slice(0, start) + kept.map(groupText).join('') + axes + x.slice(end);
}

/** Series `idx` (its c:idx) drawn as `type`: Excel's Change Series Chart Type (Series.ChartType). */
export function setSeriesType(wb, si, k, idx, type) {
  if (!SERIES_TYPES[type]) throw new Error(`No series type ${type}`);
  return change(wb, si, k, (x) => move(x, idx, type, null));
}

/** Series `idx` on the secondary value axis or back on the primary one (Series.AxisGroup 2 or 1). */
export function setSeriesAxis(wb, si, k, idx, secondary) {
  return change(wb, si, k, (x) => move(x, idx, null, !!secondary));
}
