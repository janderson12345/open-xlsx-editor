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

/** The menu's preset formats: [font colour, fill colour] (ARGB, null for none). */
export const CF_PRESETS = {
  lightRed: { label: 'Light Red Fill with Dark Red Text', font: 'FF9C0006', fill: 'FFFFC7CE' },
  yellow: { label: 'Yellow Fill with Dark Yellow Text', font: 'FF9C5700', fill: 'FFFFEB9C' },
  green: { label: 'Green Fill with Dark Green Text', font: 'FF006100', fill: 'FFC6EFCE' },
  redFill: { label: 'Light Red Fill', font: null, fill: 'FFFFC7CE' },
  redText: { label: 'Red Text', font: 'FF9C0006', fill: null },
};
/** The menu's colour scales (Excel's own defaults for three and two colours): min, (50th percentile,) max. */
export const CF_SCALES = {
  3: ['FFF8696B', 'FFFFEB84', 'FF63BE7B'],
  2: ['FFFF7128', 'FFFFEF9C'],
};
const OPS = { greaterThan: 'greaterThan', lessThan: 'lessThan', between: 'between', equal: 'equal', notBetween: 'notBetween', notEqual: 'notEqual',
  greaterThanOrEqual: 'greaterThanOrEqual', lessThanOrEqual: 'lessThanOrEqual' };
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
 * { kind: 'duplicate', unique } | { kind: 'top10', rank, percent, bottom } | { kind: 'average', below } |
 * { kind: 'expression', formula } | { kind: 'colorScale', points } (each with `preset`, but the colour scale), or
 * cfbarsedit.js's data bars and icon sets.
 */
export function cfRuleXml(rule, dxfId, priority, at) {
  const a = [`type="${{ cellIs: 'cellIs', containsText: 'containsText', duplicate: rule.unique ? 'uniqueValues' : 'duplicateValues', top10: 'top10', average: 'aboveAverage', colorScale: 'colorScale', expression: 'expression' }[rule.kind]}"`];
  if (rule.kind !== 'colorScale') a.push(`dxfId="${dxfId}"`);
  a.push(`priority="${priority}"`);
  if (rule.kind === 'average' && rule.below) a.push('aboveAverage="0"');
  if (rule.kind === 'top10' && rule.percent) a.push('percent="1"');
  if (rule.kind === 'top10' && rule.bottom) a.push('bottom="1"');
  if (rule.kind === 'cellIs') a.push(`operator="${OPS[rule.operator] || 'greaterThan'}"`);
  if (rule.kind === 'containsText') a.push('operator="containsText"', `text="${esc(rule.text)}"`);
  if (rule.kind === 'top10') a.push(`rank="${rule.rank}"`);
  let body = '';
  if (rule.kind === 'cellIs') body = `<formula>${esc(cfFormula(rule.f1))}</formula>${rule.operator === 'between' || rule.operator === 'notBetween' ? `<formula>${esc(cfFormula(rule.f2))}</formula>` : ''}`;
  // (Use a formula: the formula as typed, without its =, as Excel's New Rule writes it.)
  if (rule.kind === 'expression') body = `<formula>${esc(String(rule.formula || '').trim().replace(/^=/, ''))}</formula>`;
  if (rule.kind === 'containsText') body = `<formula>${esc(`NOT(ISERROR(SEARCH("${String(rule.text).replace(/"/g, '""')}",${at})))`)}</formula>`;
  if (rule.kind === 'colorScale') {
    const colors = CF_SCALES[rule.points === 2 ? 2 : 3];
    const cfvo = colors.length === 3 ? '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>' : '<cfvo type="min"/><cfvo type="max"/>';
    body = `<colorScale>${cfvo}${colors.map((c) => `<color rgb="${c}"/>`).join('')}</colorScale>`;
  }
  return body ? `<cfRule ${a.join(' ')}>${body}</cfRule>` : `<cfRule ${a.join(' ')}/>`;
}

/** CT_Worksheet's elements after conditionalFormatting: a new one goes before the first of them. */
const AFTER_CF = ['dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

/**
 * Add `rule` for cells `ranges` of sheet si, as the menu does: its preset's format added to the workbook's dxfs, the
 * rule first (priority 1, every other rule of the sheet one after), in the conditionalFormatting of the same ranges or a
 * new one. Edits the sheet's XML and the styles; the caller reads the sheet again.
 */
export function addConditional(wb, si, ranges, rule) {
  const doc = wb._part(si).doc;
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const p = root.prefix ? `${root.prefix}:` : '';
  let dxfId = null;
  const bars = rule.kind === 'dataBar' || rule.kind === 'iconSet';
  if (rule.kind !== 'colorScale' && !bars) {
    const sroot = wb.stylesDoc.documentElement;
    const sns = sroot.namespaceURI;
    let dxfs = kid(sroot, 'dxfs', sns);
    if (!dxfs) {
      dxfs = wb.stylesDoc.createElementNS(sns, sroot.prefix ? `${sroot.prefix}:dxfs` : 'dxfs');
      const after = [...sroot.childNodes].find((n) => n.nodeType === 1 && ['tableStyles', 'colors', 'extLst'].includes(n.localName));
      sroot.insertBefore(dxfs, after || null);
    }
    const el = wb.stylesDoc.importNode(parseXml(`<w xmlns="${sns}">${presetDxf(rule.preset)}</w>`).documentElement.firstChild, true);
    dxfs.appendChild(el);
    dxfId = kids(dxfs, 'dxf', sns).length - 1;
    dxfs.setAttribute('count', String(dxfId + 1));
    wb.styles.dxfs.push(readDxf(el, sns));
    wb.dirty.add(wb.stylesPath);
  }
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
  const xml = (base ?? cfRuleXml(rule, dxfId, 1, rangeName({ r0: ranges[0].r0, c0: ranges[0].c0, r1: ranges[0].r0, c1: ranges[0].c0 }))).replace(/<(\/?)(cfRule|formula|colorScale|cfvo|color|dataBar|iconSet|extLst|ext)\b/g, `<$1${p}$2`);
  const ruleEl = doc.importNode(parseXml(`<w xmlns${root.prefix ? `:${root.prefix}` : ''}="${ns}">${xml}</w>`).documentElement.firstChild, true);
  let cf = kids(root, 'conditionalFormatting', ns).find((e) => e.getAttribute('sqref') === sqref);
  if (!cf) {
    cf = doc.createElementNS(ns, `${p}conditionalFormatting`);
    cf.setAttribute('sqref', sqref);
    root.insertBefore(cf, [...root.childNodes].find((n) => n.nodeType === 1 && AFTER_CF.includes(n.localName)) || null);
  }
  cf.appendChild(ruleEl);
}

/** Clear Rules: from the cells `ranges` of sheet si (their cells taken out of every rule's ranges), or from the whole
 *  sheet (`ranges` null). */
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
    if (left) cf.setAttribute('sqref', left); else root.removeChild(cf);
  }
  // (The extLst's rules, data bars' settings and Excel 2010's icon sets, likewise.)
  for (const cf of [...doc.getElementsByTagNameNS(X14, 'conditionalFormatting')]) {
    const sq = kid(cf, 'sqref', XM);
    const left = ranges && sq ? trim(sq.textContent) : '';
    if (left) sq.textContent = left; else cf.parentNode.removeChild(cf);
  }
}
