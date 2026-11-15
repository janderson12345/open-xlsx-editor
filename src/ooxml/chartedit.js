// A chart's own edits, as Excel writes them (measured over COM on a clustered column chart:
// tools/excel-compare/excel-chart-edit.ps1, test/chartedit-excel.json): its title's text, or none; its legend moved or
// gone; data labels showing values; its kind changed; and the chart moved, resized or deleted. Each changes the chart
// part (or its anchor in the sheet's drawing) and nothing else. Charts are found as readCharts lists them (charts.js).
// Edited here: chart parts written with the c: and a: prefixes, as Excel and the other big editors write them.
import { chartXml, CHART_KINDS, chartLook } from './chartxml.js';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const unesc = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&');

/** Sheet si's charts, in readCharts's order: [{ drawing, anchor (its text), rid, path }]. */
export function chartParts(wb, si) {
  const pkg = wb.pkg;
  const out = [];
  for (const rel of pkg.readRels(wb.sheets[si].path).filter((r) => r.type === `${REL}/drawing`)) {
    const text = pkg.text(rel.target);
    if (!text) continue;
    const rels = pkg.readRels(rel.target);
    for (const m of text.matchAll(/<(\w+:)?twoCellAnchor\b[^>]*>[\s\S]*?<\/\1?twoCellAnchor>/g)) {
      const rid = /<(?:\w+:)?chart\b[^>]*\br:id="([^"]+)"/.exec(m[0])?.[1];
      const path = rid && rels.find((r) => r.id === rid)?.target;
      if (path && pkg.text(path)) out.push({ drawing: rel.target, anchor: m[0], rid, path });
    }
  }
  return out;
}

/** Chart k of sheet si: its part's text changed by `fn(text)`. Refused for a part not written with the c: prefix. */
export function change(wb, si, k, fn) {
  const part = chartParts(wb, si)[k];
  if (!part) throw new Error('No such chart');
  const text = wb.pkg.text(part.path);
  if (!/<c:chartSpace\b/.test(text)) throw new Error("This chart can't be edited here yet.");
  wb.pkg.setText(part.path, fn(text));
  return part;
}

const TITLE = (body) => `<c:title>${body}<c:overlay val="0"/>${chartLook.SP_NONE}${chartLook.TX_TITLE}</c:title>`;
/** The chart's title: `text` its own, null none at all ('Chart Title' or the series' name are Excel's own: 'auto'). */
export function setChartTitle(wb, si, k, text) {
  return change(wb, si, k, (x) => {
    const rich = text == null || text === 'auto' ? '' : `<c:tx><c:rich>${chartLook.BODY(0)}<a:lstStyle/><a:p><a:pPr>${chartLook.DEF(1400, ' spc="0"')}</a:pPr><a:r><a:rPr lang="en-US"/><a:t>${esc(text)}</a:t></a:r></a:p></c:rich></c:tx>`;
    let out = x;
    if (text == null) {
      out = out.replace(/<c:title>[\s\S]*?<\/c:title>/, '');
      return out.replace(/<c:autoTitleDeleted val="0"\/>/, '<c:autoTitleDeleted val="1"/>');
    }
    if (/<c:title>/.test(out)) {
      out = out.replace(/(<c:title>)(<c:tx>[\s\S]*?<\/c:tx>)?/, `$1${rich.replace(/\$/g, '$$$$')}`);
    } else {
      out = out.replace(/<c:chart>/, `<c:chart>${TITLE(rich).replace(/\$/g, '$$$$')}`);
    }
    return out.replace(/<c:autoTitleDeleted val="1"\/>/, '<c:autoTitleDeleted val="0"/>');
  });
}

/** The legend: at 'b' (bottom), 't', 'l', 'r', or null for none. */
export function setChartLegend(wb, si, k, pos) {
  return change(wb, si, k, (x) => {
    if (!pos) return x.replace(/<c:legend>[\s\S]*?<\/c:legend>/, '');
    if (/<c:legend>/.test(x)) return x.replace(/(<c:legend><c:legendPos val=")\w+/, `$1${pos}`);
    const legend = `<c:legend><c:legendPos val="${pos}"/><c:overlay val="0"/>${chartLook.SP_NONE}${chartLook.TX_LEGEND}</c:legend>`;
    return x.replace(/<\/c:plotArea>/, `</c:plotArea>${legend}`);
  });
}

/** Data labels showing each point's value (Excel's Add Chart Element > Data Labels > Show), or none. */
export function setChartLabels(wb, si, k, on) {
  return change(wb, si, k, (x) => {
    let out = x.replace(/(<c:ser>[\s\S]*?)<c:dLbls>[\s\S]*?<\/c:dLbls>(?=[\s\S]*?<\/c:ser>)/g, '$1');
    if (on) out = out.replace(/(<c:ser>[\s\S]*?)(<c:(?:trendline|errBars|cat|val|xVal|yVal)>)/g, (m, a, b) => (a.includes('<c:dLbls>') ? m : `${a}${chartLook.DLBLS_SHOW}${b}`));
    // (The chart's own: showVal on or off.)
    return out.replace(/(<\/c:ser><c:dLbls>(?:(?!<\/c:dLbls>)[\s\S])*?<c:showVal val=")[01]/, `$1${on ? 1 : 0}`);
  });
}

/** A cache's points, by index. */
const points = (cache, n) => {
  const out = new Array(n).fill(null);
  for (const m of cache.matchAll(/<c:pt idx="(\d+)"[^>]*><c:v>([\s\S]*?)<\/c:v><\/c:pt>/g)) out[Number(m[1])] = unesc(m[2]);
  return out;
};
/** A series' reference (tx, cat, val, xVal or yVal) as chartxml.js takes it: { f, values, numeric, format }. */
function refOf(ser, tag) {
  const el = new RegExp(`<c:${tag}>([\\s\\S]*?)</c:${tag}>`).exec(ser)?.[1];
  if (!el) return null;
  const f = /<c:f>([\s\S]*?)<\/c:f>/.exec(el)?.[1];
  if (f == null) return null;
  const numeric = /<c:numRef>/.test(el);
  const n = Number(/<c:ptCount val="(\d+)"/.exec(el)?.[1] || 0);
  const values = points(el, n);
  return { f: unesc(f), numeric, values: numeric ? values.map((v) => (v == null ? null : Number(v))) : values, format: unesc(/<c:formatCode>([\s\S]*?)<\/c:formatCode>/.exec(el)?.[1] || 'General') };
}

const GROUP = /<c:(barChart|bar3DChart|lineChart|line3DChart|areaChart|area3DChart|pieChart|pie3DChart|doughnutChart|scatterChart)>[\s\S]*?<\/c:\1>/;
const AXES = /<c:(catAx|valAx|dateAx|serAx)>[\s\S]*?<\/c:\1>/g;

/** An axis of the old chart turned to what `kind` wants: its position, and how the value axis crosses (Excel keeps
 *  the rest: its tick marks, number format, lines, text). A category axis becomes a scatter chart's x value axis. */
function adapt(axis, role, k) {
  let out = axis;
  const pos = k.family === 'bar' && k.dir === 'bar' ? (role === 'cat' ? 'l' : 'b') : role === 'cat' ? 'b' : 'l';
  out = out.replace(/<c:axPos val="\w+"\/>/, `<c:axPos val="${pos}"/>`);
  if (role === 'cat' && k.family === 'scatter' && /^<c:catAx>/.test(out)) {
    out = out.replace(/^<c:catAx>/, '<c:valAx>').replace(/<\/c:catAx>$/, '</c:valAx>')
      .replace(/<c:auto val="\d"\/>|<c:lblAlgn val="\w+"\/>|<c:lblOffset val="\d+"\/>|<c:noMultiLvlLbl val="\d"\/>/g, '')
      .replace(/(<c:crosses val="\w+"\/>)/, '$1<c:crossBetween val="midCat"/>');
  }
  if (role === 'val') out = out.replace(/<c:crossBetween val="\w+"\/>/, `<c:crossBetween val="${k.family === 'area' || k.family === 'scatter' ? 'midCat' : 'between'}"/>`);
  return out;
}

/**
 * The chart's kind changed (chartxml.js CHART_KINDS), as Excel's Change Chart Type does it: the series kept (their
 * references, cached points and ids) and drawn as a new chart of that kind is, but the old gap width kept, pie and
 * doughnut slices without outlines, a doughnut's hole 50%, a scatter chart's hidden line 2pt; the old axes kept and turned (adapt), and the
 * chart's other parts (title, legend, how blanks show) as they were.
 */
export function setChartKind(wb, si, k, kind) {
  const target = CHART_KINDS[kind];
  if (!target) throw new Error(`No chart kind ${kind}`);
  return change(wb, si, k, (x) => {
    const group = GROUP.exec(x);
    if (!group) throw new Error("This chart can't be changed here yet.");
    const sers = [...group[0].matchAll(/<c:ser>[\s\S]*?<\/c:ser>/g)].map((m) => m[0]);
    const series = sers.map((s) => ({ name: refOf(s, 'tx'), cat: refOf(s, 'cat') || refOf(s, 'xVal'), val: refOf(s, 'val') || refOf(s, 'yVal') }));
    if (series.some((s) => !s.val)) throw new Error("This chart's series can't be changed here yet.");
    const uid = /uniqueId val="\{\d{8}-([0-9A-F-]{27})\}"/i.exec(group[0])?.[1] || 'E1C2-4795-8E4A-B8BD83E2BD6E';
    const oldAxes = [...x.matchAll(AXES)].map((m) => m[0]);
    const ids = oldAxes.length >= 2 ? oldAxes.map((a) => Number(/<c:axId val="(\d+)"/.exec(a)[1])) : [Math.floor(1e8 + Math.random() * 2e9), Math.floor(1e8 + Math.random() * 2e9)];
    const fresh = chartXml(kind, series, { ids: { axes: ids.slice(0, 2), guid: uid } });
    let newGroup = GROUP.exec(fresh)[0];
    if (target.family === 'pie' || target.family === 'doughnut') newGroup = newGroup.replace(/<a:ln w="19050"><a:solidFill><a:schemeClr val="lt1"\/><\/a:solidFill><\/a:ln>/g, '<a:ln><a:noFill/></a:ln>');
    // (A doughnut made by changing the kind has Excel's older hole of 50%; one inserted, 75%.)
    if (target.family === 'doughnut') newGroup = newGroup.replace('<c:holeSize val="75"/>', '<c:holeSize val="50"/>');
    if (target.family === 'scatter') newGroup = newGroup.replace(/<a:ln w="38100" cap="rnd">/g, '<a:ln w="25400" cap="rnd">');
    const oldGap = /<c:gapWidth val="(\d+)"\/>/.exec(group[0])?.[1];
    if (oldGap && target.family === 'bar') newGroup = newGroup.replace(/<c:gapWidth val="\d+"\/>/, `<c:gapWidth val="${oldGap}"/>`);
    let axes = '';
    if (target.family !== 'pie' && target.family !== 'doughnut') {
      if (oldAxes.length >= 2) {
        const cat = oldAxes.find((a) => /^<c:(catAx|dateAx)>/.test(a) || /<c:axPos val="b"\/>/.test(a) && target.family === 'scatter') || oldAxes[0];
        const val = oldAxes.find((a) => a !== cat && /^<c:valAx>/.test(a)) || oldAxes[1];
        axes = adapt(cat, 'cat', target) + adapt(val, 'val', target);
      } else axes = [...fresh.matchAll(AXES)].map((m) => m[0]).join('');
    }
    const start = x.indexOf(group[0]);
    const tailStart = (() => { let at = start + group[0].length; for (const a of oldAxes) { const p = x.indexOf(a, at); if (p === at) at += a.length; } return at; })();
    return x.slice(0, start) + newGroup + axes + x.slice(tailStart);
  });
}

/** Chart k moved or sized: its anchor's corners ({ col, colOff, row, rowOff }, offsets in EMU). */
export function setChartAnchor(wb, si, k, from, to) {
  const part = chartParts(wb, si)[k];
  if (!part) throw new Error('No such chart');
  const corner = (tag, p, a) => a.replace(new RegExp(`<(\\w+:)?${tag}>[\\s\\S]*?</\\1?${tag}>`), (m, pre = '') => `<${pre}${tag}><${pre}col>${p.col}</${pre}col><${pre}colOff>${p.colOff}</${pre}colOff><${pre}row>${p.row}</${pre}row><${pre}rowOff>${p.rowOff}</${pre}rowOff></${pre}${tag}>`);
  const next = corner('to', to, corner('from', from, part.anchor));
  wb.pkg.setText(part.drawing, wb.pkg.text(part.drawing).replace(part.anchor, () => next));
  return part;
}

/** Chart k deleted: its anchor, its relationship from the drawing, and its part with what only it uses. */
export function deleteChart(wb, si, k) {
  const part = chartParts(wb, si)[k];
  if (!part) throw new Error('No such chart');
  const pkg = wb.pkg;
  pkg.setText(part.drawing, pkg.text(part.drawing).replace(part.anchor, ''));
  const dir = part.drawing.slice(0, part.drawing.lastIndexOf('/') + 1);
  const relsPath = `${dir}_rels/${part.drawing.slice(dir.length)}.rels`;
  pkg.setText(relsPath, pkg.text(relsPath).replace(new RegExp(`<Relationship\\b[^>]*\\sId="${part.rid}"[^>]*/>`), ''));
  pkg.removeTree(part.path);
  return part;
}
