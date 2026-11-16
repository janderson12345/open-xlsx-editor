// A chart's data, as Excel's Select Data changes it (measured over COM on a clustered column chart:
// tools/excel-compare/excel-chart-data.ps1, test/chartdata-excel.json): the chart's data range (SetSourceData) and
// Switch Row/Column (PlotBy), and its series one at a time: added (NewSeries), edited (Name, Values, XValues), removed
// (Delete) and moved (PlotOrder); and how hidden and empty cells show. Each changes the chart part's series (or its
// dispBlanksAs and plotVisOnly) and keeps the rest of the part as it was: the title, legend, axes, gap width.
//
// What Excel wrote, measured:
// - A series' place is its c:order; Excel writes the c:ser elements in that order. Its c:idx stays with it (and with
//   it its colour, the theme accent of its idx): moved or another removed, only the orders are numbered again.
// - A new series gets the lowest idx no series has, the next order, a unique id one past the highest, the look a new
//   chart of its kind gives series idx, and data labels when the chart shows them. No name given: no c:tx; no
//   categories: no c:cat.
// - A new data range the same way round: series k keeps the k-th series' idx, id and look (a colour of its own kept),
//   with the new references; more series are added as above, fewer drop the last. The other way round (Switch
//   Row/Column, or a range given by rows): every series made afresh, idx and order from 0, in the kind's own look.
//   With no way given, Excel keeps the chart's own, whatever the range's shape.
// - A name typed as text: <c:tx><c:v>; emptied: an empty c:v. Categories emptied: no c:cat. References are written
//   absolute ($), the sheet's name quoted only when it must be, with each cache taken from the cells.
// Not measured: a new data range after a series was removed and its colour changed (Excel gave the colour to another
// series there; kept with its series here).
import { change, chartParts } from './chartedit.js';
import { seriesXml, ref, CHART_KINDS, chartLook } from './chartxml.js';
import { chartSeries, freshIds } from './chartnew.js';
import { parseRange, colName } from './refs.js';
import { formatCodeOf } from './styles.js';
import { sheetPrefix } from '../calc/tokens.js';
import { maxOf, minOf } from './listmath.js';

const GROUP = /<c:(barChart|bar3DChart|lineChart|line3DChart|areaChart|area3DChart|pieChart|pie3DChart|doughnutChart|scatterChart)>[\s\S]*?<\/c:\1>/;
const SER = /<c:ser>[\s\S]*?<\/c:ser>/g;
const unesc = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&');
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const num = (ser, tag) => Number(new RegExp(`<c:${tag} val="(\\d+)"`).exec(ser)?.[1] ?? 0);
const inner = (ser, tag) => new RegExp(`<c:${tag}>([\\s\\S]*?)</c:${tag}>`).exec(ser)?.[1] ?? null;
const formulaIn = (el) => { const f = el == null ? null : /<c:f>([\s\S]*?)<\/c:f>/.exec(el)?.[1]; return f == null ? null : unesc(f); };

/** The CHART_KINDS name of a chart part's kind, or null for one Insert > Charts doesn't make. */
export function kindIn(text) {
  const g = GROUP.exec(text);
  if (!g) return null;
  const grouping = /<c:grouping val="(\w+)"/.exec(g[0])?.[1] || 'clustered';
  if (g[1] === 'barChart') {
    const dir = /<c:barDir val="bar"/.test(g[0]) ? 'bar' : 'column';
    return `${dir}${grouping === 'stacked' ? 'Stacked' : grouping === 'percentStacked' ? 'Stacked100' : 'Clustered'}`;
  }
  if (g[1] === 'lineChart') return /<\/c:ser>[\s\S]*<c:marker val="1"\/>/.test(g[0]) || /<c:symbol val="(?!none)/.test(g[0]) ? 'lineMarkers' : 'line';
  if (g[1] === 'areaChart') return grouping === 'stacked' ? 'areaStacked' : grouping === 'standard' ? 'area' : null;
  if (g[1] === 'pieChart' || g[1] === 'doughnutChart') return g[1].replace('Chart', '');
  if (g[1] === 'scatterChart') {
    if (/<c:scatterStyle val="smoothMarker"/.test(g[0])) return 'scatterSmooth';
    return /<c:ser>(?:(?!<\/c:ser>)[\s\S])*?<c:spPr><a:ln\b[^>]*>(?!<a:noFill)/.test(g[0]) ? 'scatterLines' : 'scatter';
  }
  return null;
}

/** A reference's sheet and cells: { si, sheet (its name), area } or null when it isn't one of this workbook's. */
function refArea(wb, f) {
  const m = /^=?\s*(?:'((?:[^']|'')+)'|([^!'=]+))!(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)\s*$/.exec(f || '');
  if (!m) return null;
  const name = (m[1] ?? m[2]).replace(/''/g, "'");
  const si = wb.sheetIndexOf(name);
  const area = parseRange(m[3].replace(/\$/g, '').toUpperCase());
  return si < 0 || !area ? null : { si, sheet: wb.sheets[si].name, area };
}

/** A range as a series formula writes it: Sheet1!$A$2:$A$5 (absolute, as Excel makes even a relative one). */
const absolute = (name, { r0, c0, r1, c1 }) => `${sheetPrefix(name)}$${colName(c0)}$${r0 + 1}${r0 === r1 && c0 === c1 ? '' : `:$${colName(c1)}$${r1 + 1}`}`;

/**
 * A reference typed for a series (=Sheet1!$B$2:$B$5; with no sheet, `own`'s) with its cells' values, as chartxml.js
 * writes it: { f, values, numeric, format }. `numeric` null: numbers when the first cell with a value has a number.
 * Throws Excel's complaint when it isn't a reference to cells of this workbook.
 */
export function refData(wb, text, numeric, own) {
  let t = String(text).trim().replace(/^=/, '');
  if (!t.includes('!')) t = `${sheetPrefix(own)}${t}`;
  const at = refArea(wb, t);
  if (!at) throw new Error('The reference is not valid. References must be to cells on a sheet of this workbook.');
  const sheet = wb.sheet(at.si);
  const cells = [];
  for (let r = at.area.r0; r <= at.area.r1; r++) for (let c = at.area.c0; c <= at.area.c1; c++) cells.push(sheet.cells.get(r)?.get(c) || null);
  const first = cells.find((cell) => cell?.v != null && cell.v !== '');
  const isNum = numeric ?? (first?.t === 'n');
  const shown = (cell) => (cell?.v == null ? null : cell.t === 'b' ? (cell.v ? 'TRUE' : 'FALSE') : cell.v);
  const values = cells.map((cell) => (isNum ? (cell?.t === 'n' && typeof cell.v === 'number' ? cell.v : null) : shown(cell) == null ? null : String(shown(cell))));
  return { f: absolute(at.sheet, at.area), values, numeric: isNum, format: isNum && first ? formatCodeOf(wb.styles, first.s || 0) : 'General' };
}

/** Chart k's series in their order: [{ ser (its text), idx, order, name: { f } | { text } | null, cat, val }]. */
function seriesIn(text) {
  const g = GROUP.exec(text);
  if (!g) throw new Error("This chart's data can't be changed here yet.");
  const list = [...g[0].matchAll(SER)].map((m) => {
    const ser = m[0];
    const tx = inner(ser, 'tx');
    const v = tx != null && /^<c:v>/.test(tx) ? unesc(/<c:v>([\s\S]*?)<\/c:v>/.exec(tx)[1]) : null;
    return { ser, idx: num(ser, 'idx'), order: num(ser, 'order'), name: tx == null ? null : v != null ? { text: v } : { f: formulaIn(tx) },
      cat: formulaIn(inner(ser, 'cat') ?? inner(ser, 'xVal')), val: formulaIn(inner(ser, 'val') ?? inner(ser, 'yVal')) };
  });
  return { group: g[0], list: list.sort((a, b) => a.order - b.order) };
}

/** Whether the series run along rows: their values across a row (else, for one cell each, their categories). */
function alongRows(wb, list) {
  for (const s of list) {
    const v = refArea(wb, s.val)?.area;
    if (v && v.c1 > v.c0 && v.r0 === v.r1) return true;
    if (v && v.r1 > v.r0) return false;
    const c = refArea(wb, s.cat)?.area;
    if (c && c.c1 > c.c0) return true;
    if (c && c.r1 > c.r0) return false;
  }
  return false;
}

/** The range the series were taken from, as Excel's Chart data range shows it (=Sheet1!$A$1:$C$5), or null when
 *  they are not the series that range gives (Excel: "too complex to be displayed"). */
function rangeOf(wb, list, byRows) {
  const refs = list.flatMap((s) => [s.name?.f, s.cat, s.val]).filter(Boolean).map((f) => refArea(wb, f));
  if (!refs.length || refs.some((x) => !x || x.si !== refs[0].si)) return null;
  const box = { r0: minOf(refs.map((x) => x.area.r0)), c0: minOf(refs.map((x) => x.area.c0)), r1: maxOf(refs.map((x) => x.area.r1)), c1: maxOf(refs.map((x) => x.area.c1)) };
  const found = chartSeries(wb, refs[0].si, box, byRows);
  const same = found && found.series.length === list.length && found.series.every((s, k) => (s.name?.f ?? null) === (list[k].name?.f ?? null)
    && (s.cat?.f ?? null) === list[k].cat && s.val.f === list[k].val);
  return same ? `=${absolute(refs[0].sheet, box)}` : null;
}

/**
 * Chart k of sheet si's data as Select Data shows it: { range (or null), byRows, series: [{ idx, name, cat, val }]
 * in their order, blanks ('gap', 'zero' or 'span'), hidden (whether hidden rows' and columns' data is drawn), kind }.
 */
export function chartData(wb, si, k) {
  const part = chartParts(wb, si)[k];
  if (!part) throw new Error('No such chart');
  const text = wb.pkg.text(part.path);
  const { list } = seriesIn(text);
  const byRows = alongRows(wb, list);
  return {
    range: rangeOf(wb, list, byRows), byRows, kind: kindIn(text),
    series: list.map(({ idx, name, cat, val }) => ({ idx, name, cat, val })),
    blanks: /<c:dispBlanksAs val="(\w+)"/.exec(text)?.[1] || (/<c:dispBlanksAs\b/.test(text) ? 'zero' : 'gap'),
    hidden: /<c:plotVisOnly val="0"/.test(text),
  };
}

/** The group's series put back in place of the old ones, in their order. (Excel's category axis takes its linked
 *  number format from the first series' categories: with none there, it has no c:numFmt: xvaluesEmpty.) */
function withSeries(text, group, sers) {
  const first = group.search(/<c:ser>/);
  const rest = group.replace(SER, '');
  const at = first < 0 ? rest.search(/<c:(dLbls|gapWidth|overlap|marker|smooth|firstSliceAng|holeSize|axId)\b/) : first;
  const next = rest.slice(0, at) + sers.join('') + rest.slice(at);
  const out = text.replace(group, () => next);
  if (!sers.length || /<c:cat>/.test(sers[0])) return out;
  return out.replace(/<c:catAx>[\s\S]*?<\/c:catAx>/, (ax) => ax.replace(/<c:numFmt formatCode="[^"]*" sourceLinked="1"\/>/, ''));
}

/** A series' name as written: a reference with its cache, text of its own, or none. */
const nameXml = (name) => (name == null ? '' : 'text' in name ? `<c:tx><c:v>${esc(name.text)}</c:v></c:tx>` : ref('tx', { ...name, numeric: false }));

/** A series with new references (a name, categories, values: given ones replaced, null for categories none). */
function rewire(ser, scatter, { name, cat, val }) {
  const [catTag, valTag] = scatter ? ['xVal', 'yVal'] : ['cat', 'val'];
  let out = ser;
  if (name !== undefined) out = out.replace(/<c:tx>[\s\S]*?<\/c:tx>/, '').replace(/(<c:order val="\d+"\/>)/, (m) => m + nameXml(name));
  if (val) out = out.replace(new RegExp(`<c:${valTag}>[\\s\\S]*?</c:${valTag}>`), () => ref(valTag, val));
  if (cat !== undefined) {
    out = out.replace(new RegExp(`<c:${catTag}>[\\s\\S]*?</c:${catTag}>`), '');
    if (cat) out = out.replace(`<c:${valTag}>`, () => `${ref(catTag, cat)}<c:${valTag}>`);
  }
  return out;
}

/** The chart shows its values as data labels (Excel gives a new series them too). */
const labelled = (list) => list.some((s) => /<c:dLbls>(?:(?!<\/c:dLbls>)[\s\S])*?<c:showVal val="1"/.test(s.ser));
const uidOf = (ser) => /uniqueId val="\{([0-9A-F]{8})-([0-9A-F-]{27})\}"/i.exec(ser);

/** A new series of `kind`: idx, order, unique id number `uid`, the chart's id tail `guid`, data labels when `labels`. */
function freshSeries(kind, s, { idx, order, uid, guid, labels, outline, hiddenLine }) {
  const k = CHART_KINDS[kind];
  let x = seriesXml(k, { ...s, name: null }, idx, guid)
    .replace(/<c:order val="\d+"\/>/, `<c:order val="${order}"/>${nameXml(s.name)}`)
    .replace(/(uniqueId val="\{)\d{8}/, `$1${String(uid).padStart(8, '0')}`);
  if (labels) x = x.replace(/(<c:(?:trendline|errBars|cat|val|xVal|yVal)>)/, (m) => chartLook.DLBLS_SHOW + m);
  // (In a pie whose slices have no outlines, or a scatter chart whose hidden lines are another width, as Change Chart
  // Type leaves them, Excel draws a new series as the others: pieSource, scatterSource.)
  if (outline === false) x = x.replace(/<a:ln w="19050"><a:solidFill><a:schemeClr val="lt1"\/><\/a:solidFill><\/a:ln>/g, '<a:ln><a:noFill/></a:ln>');
  if (hiddenLine) x = x.replace(/<a:ln w="\d+" cap="rnd"><a:noFill\/>/, `<a:ln w="${hiddenLine}" cap="rnd"><a:noFill/>`);
  return x;
}

/** What a new series needs of the chart: its kind (refused when Insert > Charts doesn't make it), labels, ids. */
function freshContext(text, list) {
  const kind = kindIn(text);
  if (!kind) throw new Error("Series can't be added to this kind of chart here yet.");
  const uids = list.map((s) => uidOf(s.ser)).filter(Boolean);
  const one = list[0]?.ser || '';
  return { kind, labels: labelled(list), guid: uids[0]?.[2] || freshIds().guid, nextUid: uids.length ? maxOf(uids.map((u) => parseInt(u[1], 16))) + 1 : list.length,
    outline: /<c:dPt>/.test(one) ? !/<c:dPt>(?:(?!<\/c:dPt>)[\s\S])*?<a:ln><a:noFill\/><\/a:ln>/.test(one) : undefined,
    hiddenLine: /<c:spPr><a:ln w="(\d+)" cap="rnd"><a:noFill\/>/.exec(one)?.[1] };
}
const freeIdx = (used) => { let i = 0; while (used.has(i)) i++; used.add(i); return i; };

/**
 * The chart's data range (Excel's Chart data range, SetSourceData): `range` a reference (=Sheet1!$A$1:$C$5), its
 * series taken from it as Insert > Charts takes them, along the rows when `byRows`, down the columns when false, the
 * chart's own way when null (as Excel does).
 */
export function setChartData(wb, si, k, range, byRows = null) {
  const at = refArea(wb, String(range).trim().includes('!') ? range : `${sheetPrefix(wb.sheets[si].name)}${String(range).trim().replace(/^=/, '')}`);
  if (!at) throw new Error('The reference is not valid. References must be to cells on a sheet of this workbook.');
  return change(wb, si, k, (text) => {
    const { group, list } = seriesIn(text);
    const was = alongRows(wb, list);
    const way = byRows ?? was;
    const found = chartSeries(wb, at.si, at.area, way);
    if (!found || !found.series.some((s) => s.val.values.some((v) => v != null))) throw new Error('Select a range with numbers in it to chart.');
    const ctx = freshContext(text, list);
    const scatter = CHART_KINDS[ctx.kind].family === 'scatter';
    let sers;
    if (way === was) {
      const used = new Set(found.series.slice(0, list.length).map((_, n) => list[n].idx));
      let uid = ctx.nextUid;
      sers = found.series.map((s, n) => (n < list.length
        ? rewire(list[n].ser, scatter, { name: s.name, cat: s.cat, val: s.val }).replace(/<c:order val="\d+"\/>/, `<c:order val="${n}"/>`)
        : freshSeries(ctx.kind, s, { ...ctx, idx: freeIdx(used), order: n, uid: uid++ })));
    } else {
      sers = found.series.map((s, n) => freshSeries(ctx.kind, s, { ...ctx, idx: n, order: n, uid: n }));
    }
    return withSeries(text, group, sers);
  });
}

/** The series' parts typed in the Edit Series box: a name (text, or =reference; '' none), values and categories
 *  (=references; categories '' none) as chartxml.js writes them. */
function parts(wb, si, { name, val, cat }) {
  const own = wb.sheets[si].name;
  const out = {};
  if (name !== undefined) out.name = name == null ? null : String(name).trim().startsWith('=') ? refData(wb, name, false, own) : { text: String(name) };
  if (val !== undefined) out.val = refData(wb, val, true, own);
  if (cat !== undefined) out.cat = cat == null || String(cat).trim() === '' ? null : refData(wb, cat, null, own);
  return out;
}

/** A series added (Select Data's Add, NewSeries): { name, val, cat } as the Edit Series box takes them. */
export function addChartSeries(wb, si, k, series) {
  const p = parts(wb, si, { cat: '', name: null, ...series });
  if (!p.val) throw new Error('A series needs its values.');
  return change(wb, si, k, (text) => {
    const { group, list } = seriesIn(text);
    const ctx = freshContext(text, list);
    const idx = freeIdx(new Set(list.map((s) => s.idx)));
    const add = freshSeries(ctx.kind, { name: p.name, cat: p.cat, val: p.val }, { ...ctx, idx, order: list.length, uid: ctx.nextUid });
    return withSeries(text, group, [...list.map((s) => s.ser), add]);
  });
}

/** Series n (by its place) edited: any of { name, val, cat } as the Edit Series box takes them. */
export function editChartSeries(wb, si, k, n, series) {
  const p = parts(wb, si, series);
  return change(wb, si, k, (text) => {
    const { group, list } = seriesIn(text);
    if (!list[n]) throw new Error('No such series');
    const scatter = /^<c:scatterChart>/.test(group);
    return withSeries(text, group, list.map((s, m) => (m === n ? rewire(s.ser, scatter, p) : s.ser)));
  });
}

/** The series in a new order: each keeps its idx (and colour), the orders numbered again. */
const reordered = (list) => list.map((s, m) => s.ser.replace(/<c:order val="\d+"\/>/, `<c:order val="${m}"/>`));

/** Series n removed (Select Data's Remove, Delete). */
export function removeChartSeries(wb, si, k, n) {
  return change(wb, si, k, (text) => {
    const { group, list } = seriesIn(text);
    if (!list[n] || list.length < 2) throw new Error(list[n] ? 'A chart keeps at least one series.' : 'No such series');
    return withSeries(text, group, reordered(list.filter((_, m) => m !== n)));
  });
}

/** Series n moved up (-1) or down (+1) a place (Select Data's Move Up and Move Down, PlotOrder). */
export function moveChartSeries(wb, si, k, n, by) {
  return change(wb, si, k, (text) => {
    const { group, list } = seriesIn(text);
    const to = n + by;
    if (!list[n] || !list[to]) throw new Error('No such series');
    [list[n], list[to]] = [list[to], list[n]];
    return withSeries(text, group, reordered(list));
  });
}

/** Hidden and Empty Cells: empty cells as gaps ('gap'), zero ('zero') or the line carried over them ('span'), and
 *  whether data in hidden rows and columns is drawn (DisplayBlanksAs and PlotVisibleOnly). */
export function setChartBlanks(wb, si, k, { blanks, hidden }) {
  return change(wb, si, k, (text) => {
    let out = text;
    if (blanks) out = /<c:dispBlanksAs\b/.test(out) ? out.replace(/<c:dispBlanksAs\b[^>]*\/>/, `<c:dispBlanksAs val="${blanks}"/>`) : out.replace(/(<c:plotVisOnly\b[^>]*\/>)/, `$1<c:dispBlanksAs val="${blanks}"/>`);
    if (hidden != null) out = out.replace(/<c:plotVisOnly\b[^>]*\/>/, `<c:plotVisOnly val="${hidden ? 0 : 1}"/>`);
    return out;
  });
}
