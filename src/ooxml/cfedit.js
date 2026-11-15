// Conditional formatting rules added as Excel's Conditional Formatting menu adds them (Home > Conditional Formatting:
// Highlight Cells Rules, Top/Bottom Rules, Color Scales; Clear Rules), measured: tools/excel-compare/excel-cf-write.ps1
// had Excel add each rule with the menu's preset formats (test/cf-write-excel.json); this writes the same cfRule
// (attributes in Excel's order, its formulas) and dxf (the preset's font colour, its fill as the pattern's background).
// Two things COM does its own way and the menu doesn't: COM's rules Stop If True, and COM gives a later rule the lower
// priority. Here, as the menu does, a new rule doesn't stop and comes first (priority 1, the others after it: the
// Rules Manager lists the newest first).
import { kids, kid, parseXml } from './xml.js';
import { parseRange, rangeName, MAX_ROWS, MAX_COLS } from './refs.js';
import { readDxf } from './styles.js';
import { parseInput } from './input.js';
import { subtract } from './dvedit.js';
import { barRuleXml, addX14Rule, x14Rules } from './cfbarsedit.js';
import { X14, XM } from './cfext.js';
import { shiftFormula } from '../calc/tokens.js';

/** The menu's preset formats: [font colour, fill colour] (ARGB, null for none). */
export const CF_PRESETS = {
  lightRed: { label: 'Light Red Fill with Dark Red Text', font: 'FF9C0006', fill: 'FFFFC7CE' },
  yellow: { label: 'Yellow Fill with Dark Yellow Text', font: 'FF9C5700', fill: 'FFFFEB9C' },
  green: { label: 'Green Fill with Dark Green Text', font: 'FF006100', fill: 'FFC6EFCE' },
  redFill: { label: 'Light Red Fill', font: null, fill: 'FFFFC7CE' },
  redText: { label: 'Red Text', font: 'FF9C0006', fill: null },
};
/** Excel's Color Scales gallery: each preset's colours from the lowest value to the highest, named as Excel names
 *  them, the highest's colour first (Excel's standard scale colours, as its files have them; the gallery isn't
 *  reachable over COM, so they aren't measured here). */
const R = 'FFF8696B'; const Y = 'FFFFEB84'; const G = 'FF63BE7B'; const W = 'FFFCFCFF'; const B = 'FF5A8AC6'; const Y2 = 'FFFFEF9C';
export const CF_GALLERY = [
  ['Green - Yellow - Red', [R, Y, G]], ['Red - Yellow - Green', [G, Y, R]], ['Green - White - Red', [R, W, G]], ['Red - White - Green', [G, W, R]],
  ['Blue - White - Red', [R, W, B]], ['Red - White - Blue', [B, W, R]], ['White - Red', [R, W]], ['Red - White', [W, R]],
  ['Green - White', [W, G]], ['White - Green', [G, W]], ['Green - Yellow', [Y2, G]], ['Yellow - Green', [G, Y2]],
];
/** The menu's colour scales (Excel's own defaults for three and two colours): min, (50th percentile,) max. */
export const CF_SCALES = {
  3: ['FFF8696B', 'FFFFEB84', 'FF63BE7B'],
  2: ['FFFF7128', 'FFFFEF9C'],
};
const OPS = { greaterThan: 'greaterThan', lessThan: 'lessThan', between: 'between', equal: 'equal', notBetween: 'notBetween', notEqual: 'notEqual',
  greaterThanOrEqual: 'greaterThanOrEqual', lessThanOrEqual: 'lessThanOrEqual' };
/** A Date Occurring's periods, in the menu's order, with their labels and the formulas Excel writes for them (`@` the
 *  range's first cell; measured: tools/excel-compare/excel-cf-dates.ps1, test/cf-dates-excel.json). */
export const CF_PERIODS = {
  yesterday: ['Yesterday', 'FLOOR(@,1)=TODAY()-1'],
  today: ['Today', 'FLOOR(@,1)=TODAY()'],
  tomorrow: ['Tomorrow', 'FLOOR(@,1)=TODAY()+1'],
  last7Days: ['In the last 7 days', 'AND(TODAY()-FLOOR(@,1)<=6,FLOOR(@,1)<=TODAY())'],
  lastWeek: ['Last week', 'AND(TODAY()-ROUNDDOWN(@,0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(@,0)<(WEEKDAY(TODAY())+7))'],
  thisWeek: ['This week', 'AND(TODAY()-ROUNDDOWN(@,0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(@,0)-TODAY()<=7-WEEKDAY(TODAY()))'],
  nextWeek: ['Next week', 'AND(ROUNDDOWN(@,0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(@,0)-TODAY()<(15-WEEKDAY(TODAY())))'],
  lastMonth: ['Last month', 'AND(MONTH(@)=MONTH(EDATE(TODAY(),0-1)),YEAR(@)=YEAR(EDATE(TODAY(),0-1)))'],
  thisMonth: ['This month', 'AND(MONTH(@)=MONTH(TODAY()),YEAR(@)=YEAR(TODAY()))'],
  nextMonth: ['Next month', 'AND(MONTH(@)=MONTH(EDATE(TODAY(),0+1)),YEAR(@)=YEAR(EDATE(TODAY(),0+1)))'],
};
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** A dxf for a preset, as Excel writes it. */
export function presetDxf(name) {
  const p = CF_PRESETS[name] || CF_PRESETS.lightRed;
  return `<dxf>${p.font ? `<font><color rgb="${p.font}"/></font>` : ''}${p.fill ? `<fill><patternFill><bgColor rgb="${p.fill}"/></patternFill></fill>` : ''}</dxf>`;
}

/** A value as typed into the menu's boxes, as its rule's formula: a number as it is, "=..." as a formula, text quoted. */
export function cfFormula(typed) {
  const t = String(typed ?? '').trim();
  if (t.startsWith('=')) return t.slice(1);
  const v = parseInput(t).value;
  return typeof v === 'number' ? String(v) : `"${t.replace(/"/g, '""')}"`;
}

/**
 * The cfRule for `rule` (dxfId the preset's format, priority as given, `at` the range's first cell for formulas that
 * read it), as Excel writes it. `rule`: { kind: 'cellIs', operator, f1, f2 } | { kind: 'containsText', text } |
 * { kind: 'timePeriod', period } (CF_PERIODS) | { kind: 'duplicate', unique } | { kind: 'top10', rank, percent, bottom } | { kind: 'average', below } |
 * { kind: 'expression', formula } | { kind: 'colorScale', points } (each with `preset`, but the colour scale), or
 * cfbarsedit.js's data bars and icon sets.
 */
export function cfRuleXml(rule, dxfId, priority, at) {
  const a = [`type="${{ cellIs: 'cellIs', containsText: 'containsText', duplicate: rule.unique ? 'uniqueValues' : 'duplicateValues', top10: 'top10', timePeriod: 'timePeriod', average: 'aboveAverage', colorScale: 'colorScale', expression: 'expression' }[rule.kind]}"`];
  if (rule.kind !== 'colorScale') a.push(`dxfId="${dxfId}"`);
  a.push(`priority="${priority}"`);
  if (rule.kind === 'average' && rule.below) a.push('aboveAverage="0"');
  if (rule.kind === 'top10' && rule.percent) a.push('percent="1"');
  if (rule.kind === 'top10' && rule.bottom) a.push('bottom="1"');
  if (rule.kind === 'cellIs') a.push(`operator="${OPS[rule.operator] || 'greaterThan'}"`);
  if (rule.kind === 'containsText') a.push('operator="containsText"', `text="${esc(rule.text)}"`);
  if (rule.kind === 'top10') a.push(`rank="${rule.rank}"`);
  const period = CF_PERIODS[rule.period] ? rule.period : 'yesterday';
  if (rule.kind === 'timePeriod') a.push(`timePeriod="${period}"`);
  let body = '';
  if (rule.kind === 'cellIs') body = `<formula>${esc(cfFormula(rule.f1))}</formula>${rule.operator === 'between' || rule.operator === 'notBetween' ? `<formula>${esc(cfFormula(rule.f2))}</formula>` : ''}`;
  // (Use a formula: the formula as typed, without its =, as Excel's New Rule writes it.)
  if (rule.kind === 'expression') body = `<formula>${esc(String(rule.formula || '').trim().replace(/^=/, ''))}</formula>`;
  if (rule.kind === 'timePeriod') body = `<formula>${esc(CF_PERIODS[period][1].replaceAll('@', at))}</formula>`;
  if (rule.kind === 'containsText') body = `<formula>${esc(`NOT(ISERROR(SEARCH("${String(rule.text).replace(/"/g, '""')}",${at})))`)}</formula>`;
  if (rule.kind === 'colorScale') {
    // (Thresholds and colours of its own, `stops` [{ type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula',
    // val, color }], as Excel writes them (excel-cf-scale.ps1); else a gallery's colours (`colors`) or Excel's default.)
    if (rule.stops?.length) {
      const cfvo = rule.stops.map((s) => (s.type === 'min' || s.type === 'max' ? `<cfvo type="${s.type}"/>` : `<cfvo type="${s.type}" val="${esc(String(s.val ?? '').replace(/^=/, ''))}"/>`)).join('');
      body = `<colorScale>${cfvo}${rule.stops.map((s) => `<color rgb="${s.color}"/>`).join('')}</colorScale>`;
    } else {
      const colors = rule.colors?.length ? rule.colors : CF_SCALES[rule.points === 2 ? 2 : 3];
      const cfvo = colors.length === 3 ? '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>' : '<cfvo type="min"/><cfvo type="max"/>';
      body = `<colorScale>${cfvo}${colors.map((c) => `<color rgb="${c}"/>`).join('')}</colorScale>`;
    }
  }
  return body ? `<cfRule ${a.join(' ')}>${body}</cfRule>` : `<cfRule ${a.join(' ')}/>`;
}

/** CT_Worksheet's elements after conditionalFormatting: a new one goes before the first of them. */
export const AFTER_CF = ['dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

/**
 * Add `rule` for cells `ranges` of sheet si, as the menu does: its preset's format added to the workbook's dxfs, the
 * rule first (priority 1, every other rule of the sheet one after), in the conditionalFormatting of the same ranges or a
 * new one. Edits the sheet's XML and the styles; the caller reads the sheet again.
 */
/**
 * A rule's own format (New Rule's Format...) as a dxf, as Excel writes one (tools/excel-compare/excel-cf-format.ps1):
 * font (bold, italic, underline, strikethrough, colour), then the number format (numbered `numFmtId` from 164, as
 * Excel numbers its own), the fill as the pattern's background, and the border's sides in Excel's order, each with its
 * colour or automatic. `f`: { bold, italic, underline, strike, color, numFmt, fill, border: { left, right, top,
 * bottom: { style, color } } } (colours ARGB).
 */
export function formatDxf(f, numFmtId = 164) {
  const font = [f.bold && '<b/>', f.italic && '<i/>', f.underline && '<u/>', f.strike && '<strike/>', f.color && `<color rgb="${f.color}"/>`].filter(Boolean).join('');
  const num = f.numFmt ? `<numFmt numFmtId="${numFmtId}" formatCode="${esc(f.numFmt)}"/>` : '';
  const fill = f.fill ? `<fill><patternFill><bgColor rgb="${f.fill}"/></patternFill></fill>` : '';
  const sides = ['left', 'right', 'top', 'bottom'].filter((k) => f.border?.[k]?.style)
    .map((k) => `<${k} style="${f.border[k].style}">${f.border[k].color ? `<color rgb="${f.border[k].color}"/>` : '<color auto="1"/>'}</${k}>`).join('');
  return `<dxf>${font ? `<font>${font}</font>` : ''}${num}${fill}${sides ? `<border>${sides}</border>` : ''}</dxf>`;
}

/** A preset's format (or a rule's own, an object: formatDxf) added to the workbook's dxfs (made when it has none): its
 *  dxfId. */
export function addPresetDxf(wb, preset) {
  const sroot = wb.stylesDoc.documentElement;
  const sns = sroot.namespaceURI;
  let dxfs = kid(sroot, 'dxfs', sns);
  if (!dxfs) {
    dxfs = wb.stylesDoc.createElementNS(sns, sroot.prefix ? `${sroot.prefix}:dxfs` : 'dxfs');
    const after = [...sroot.childNodes].find((n) => n.nodeType === 1 && ['tableStyles', 'colors', 'extLst'].includes(n.localName));
    sroot.insertBefore(dxfs, after || null);
  }
  // (A number format of its own takes the next id from 164 not yet used in the styles.)
  const used = [...wb.stylesDoc.getElementsByTagNameNS(sns, 'numFmt')].map((n) => Number(n.getAttribute('numFmtId')) || 0);
  const xml = preset && typeof preset === 'object' ? formatDxf(preset, Math.max(163, ...used) + 1) : presetDxf(preset);
  const el = wb.stylesDoc.importNode(parseXml(`<w xmlns="${sns}">${xml}</w>`).documentElement.firstChild, true);
  dxfs.appendChild(el);
  const dxfId = kids(dxfs, 'dxf', sns).length - 1;
  dxfs.setAttribute('count', String(dxfId + 1));
  wb.styles.dxfs.push(readDxf(el, sns));
  wb.dirty.add(wb.stylesPath);
  return dxfId;
}

/** A cfRule's XML (cfRuleXml's, or barRuleXml's base) as an element of the sheet's part, in the part's prefix. */
export function ruleElement(doc, xml) {
  const root = doc.documentElement;
  const p = root.prefix ? `${root.prefix}:` : '';
  const prefixed = xml.replace(/<(\/?)(cfRule|formula|colorScale|cfvo|color|dataBar|iconSet|extLst|ext)\b/g, `<$1${p}$2`);
  return doc.importNode(parseXml(`<w xmlns${root.prefix ? `:${root.prefix}` : ''}="${root.namespaceURI}">${prefixed}</w>`).documentElement.firstChild, true);
}

export function addConditional(wb, si, ranges, rule) {
  const doc = wb._part(si).doc;
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const p = root.prefix ? `${root.prefix}:` : '';
  const bars = rule.kind === 'dataBar' || rule.kind === 'iconSet';
  const dxfId = rule.kind !== 'colorScale' && !bars ? addPresetDxf(wb, rule.preset) : null;
  // Every rule of the sheet one later: the new one first.
  // (Excel 2010's rules in the extLst have priorities among the others: cfbarsedit.js.)
  for (const r of [...doc.getElementsByTagNameNS(ns, 'cfRule'), ...x14Rules(doc).filter((x) => x.hasAttribute('priority'))]) r.setAttribute('priority', String(Number(r.getAttribute('priority') || 0) + 1));
  const sqref = ranges.map(rangeName).join(' ');
  let base = null;
  if (bars) {
    const made = barRuleXml(rule, 1);
    if (made.x14) addX14Rule(doc, made.x14, sqref);
    base = made.base;
    if (!base) return;
  }
  const ruleEl = ruleElement(doc, base ?? cfRuleXml(rule, dxfId, 1, rangeName({ r0: ranges[0].r0, c0: ranges[0].c0, r1: ranges[0].r0, c1: ranges[0].c0 })));
  let cf = kids(root, 'conditionalFormatting', ns).find((e) => e.getAttribute('sqref') === sqref);
  if (!cf) {
    cf = doc.createElementNS(ns, `${p}conditionalFormatting`);
    cf.setAttribute('sqref', sqref);
    root.insertBefore(cf, [...root.childNodes].find((n) => n.nodeType === 1 && AFTER_CF.includes(n.localName)) || null);
  }
  cf.appendChild(ruleEl);
}

/** Clear Rules: from the cells `ranges` of sheet si (their cells taken out of every rule's ranges), or from the whole
 *  sheet (`ranges` null). A rule left without its first cell has its formulas written for its new first cell, as
 *  Excel's do (measured: =E1>F1 over E1:E4, E1:E2 cleared or pasted over, is =E3>F3 over E3:E4). */
export function clearConditional(wb, si, ranges = null) {
  const doc = wb._part(si).doc;
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const trim = (sqref) => {
    let parts = (sqref || '').split(/\s+/).map(parseRange).filter(Boolean);
    for (const s of ranges) parts = parts.flatMap((x) => subtract(x, { r0: s.r0, c0: s.c0, r1: Math.min(s.r1, MAX_ROWS - 1), c1: Math.min(s.c1, MAX_COLS - 1) }));
    return parts.map(rangeName).join(' ');
  };
  for (const cf of kids(root, 'conditionalFormatting', ns)) {
    const left = ranges ? trim(cf.getAttribute('sqref')) : '';
    if (!left) { root.removeChild(cf); continue; }
    const was = parseRange((cf.getAttribute('sqref') || '').split(/\s+/)[0]);
    const now = parseRange(left.split(/\s+/)[0]);
    cf.setAttribute('sqref', left);
    if (was && now && (was.r0 !== now.r0 || was.c0 !== now.c0)) {
      for (const f of cf.getElementsByTagNameNS(ns, 'formula')) f.textContent = shiftFormula(f.textContent, now.r0 - was.r0, now.c0 - was.c0);
    }
  }
  // (The extLst's rules, data bars' settings and Excel 2010's icon sets, likewise.)
  for (const cf of [...doc.getElementsByTagNameNS(X14, 'conditionalFormatting')]) {
    const sq = kid(cf, 'sqref', XM);
    const left = ranges && sq ? trim(sq.textContent) : '';
    if (left) sq.textContent = left; else cf.parentNode.removeChild(cf);
  }
}
