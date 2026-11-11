// Data validation rules written as Excel's Data Validation box writes them (measured: test/dv-excel.json holds the
// dataValidation element Excel saved for each of 15 rules; this writes the same, but for Excel's xr:uid). A rule for
// the selection takes its cells out of any rule they had (those rules' ranges split around them), then is added after
// the others; Clear All only takes them out.
import { parseInput } from './input.js';
import { parseRange, rangeName, MAX_ROWS, MAX_COLS } from './refs.js';
import { parseXml } from './xml.js';

/** Excel's names for the box's choices. */
const TYPES = { any: null, whole: 'whole', decimal: 'decimal', list: 'list', date: 'date', time: 'time', length: 'textLength', custom: 'custom' };
const OPERATORS = { between: 'between', notBetween: 'notBetween', equal: 'equal', notEqual: 'notEqual', greater: 'greaterThan', less: 'lessThan', greaterEqual: 'greaterThanOrEqual', lessEqual: 'lessThanOrEqual' };
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** A number as Excel writes it in a formula: 15 significant digits at most, no trailing zeros. */
const number = (n) => String(Number(n.toPrecision(15)));

/**
 * A rule's formula as Excel stores it, from what was typed in the box: a reference or formula without its "=", a
 * typed list in quotes, a date or time as its serial number, a number as it is. `date1904` for dates.
 */
export function storedFormula(type, typed, date1904 = false) {
  const t = String(typed ?? '').trim();
  if (t.startsWith('=')) return t.slice(1);
  if (type === 'list') return `"${t.replace(/"/g, '""')}"`;
  if (type === 'custom') return t;
  const v = parseInput(t, { date1904 }).value;
  return typeof v === 'number' ? number(v) : t;
}

/** The dataValidation element for `rule` on `ranges` ([{ r0, c0, r1, c1 }]), as Excel writes it. `rule`: { type
 *  ('any', 'whole', 'decimal', 'list', 'date', 'time', 'length', 'custom'), op, f1, f2 (as typed), ignoreBlank, dropDown,
 *  prompt: { show, title, text }, error: { show, style ('stop', 'warning', 'information'), title, text } }. */
export function validationXml(rule, ranges, date1904 = false) {
  const type = TYPES[rule.type] ?? null;
  const op = OPERATORS[rule.op] || 'between';
  const a = [];
  if (type) a.push(`type="${type}"`);
  if (rule.error?.style && rule.error.style !== 'stop') a.push(`errorStyle="${rule.error.style}"`);
  const ranged = type && type !== 'list' && type !== 'custom';
  if (ranged && op !== 'between') a.push(`operator="${op}"`);
  if (rule.ignoreBlank !== false) a.push('allowBlank="1"');
  if (type === 'list' && rule.dropDown === false) a.push('showDropDown="1"');
  if (rule.prompt?.show !== false) a.push('showInputMessage="1"');
  if (rule.error?.show !== false) a.push('showErrorMessage="1"');
  if (rule.error?.title) a.push(`errorTitle="${esc(rule.error.title)}"`);
  if (rule.error?.text) a.push(`error="${esc(rule.error.text)}"`);
  if (rule.prompt?.title) a.push(`promptTitle="${esc(rule.prompt.title)}"`);
  if (rule.prompt?.text) a.push(`prompt="${esc(rule.prompt.text)}"`);
  a.push(`sqref="${ranges.map(rangeName).join(' ')}"`);
  const two = ranged && (op === 'between' || op === 'notBetween');
  const formulas = type ? `<formula1>${esc(storedFormula(rule.type, rule.f1, date1904))}</formula1>${two ? `<formula2>${esc(storedFormula(rule.type, rule.f2, date1904))}</formula2>` : ''}` : '';
  return formulas ? `<dataValidation ${a.join(' ')}>${formulas}</dataValidation>` : `<dataValidation ${a.join(' ')}/>`;
}

/** Range `a` less range `b`: the rectangles left (bands above and below, then left and right between them). */
export function subtract(a, b) {
  if (b.r1 < a.r0 || b.r0 > a.r1 || b.c1 < a.c0 || b.c0 > a.c1) return [a];
  const out = [];
  if (b.r0 > a.r0) out.push({ r0: a.r0, c0: a.c0, r1: b.r0 - 1, c1: a.c1 });
  if (b.r1 < a.r1) out.push({ r0: b.r1 + 1, c0: a.c0, r1: a.r1, c1: a.c1 });
  const r0 = Math.max(a.r0, b.r0);
  const r1 = Math.min(a.r1, b.r1);
  if (b.c0 > a.c0) out.push({ r0, c0: a.c0, r1, c1: b.c0 - 1 });
  if (b.c1 < a.c1) out.push({ r0, c0: b.c1 + 1, r1, c1: a.c1 });
  return out;
}

/**
 * Give sheet si's cells `ranges` the validation `rule` (null: Clear All), as Excel's box does. Edits the sheet's XML
 * (`doc`); the caller reads the sheet again.
 */
export function setValidation(doc, ranges, rule, date1904 = false) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const sel = ranges.map((r) => ({ r0: r.r0, c0: r.c0, r1: Math.min(r.r1, MAX_ROWS - 1), c1: Math.min(r.c1, MAX_COLS - 1) }));
  let wrap = [...root.childNodes].find((n) => n.localName === 'dataValidations');
  // The selection's cells out of the rules they were in; a rule left with none goes.
  for (const el of wrap ? [...wrap.childNodes].filter((n) => n.localName === 'dataValidation') : []) {
    let parts = (el.getAttribute('sqref') || '').split(/\s+/).map(parseRange).filter(Boolean);
    for (const s of sel) parts = parts.flatMap((p) => subtract(p, s));
    if (!parts.length) wrap.removeChild(el);
    else el.setAttribute('sqref', parts.map(rangeName).join(' '));
  }
  if (rule) {
    if (!wrap) {
      wrap = doc.createElementNS(ns, root.prefix ? `${root.prefix}:dataValidations` : 'dataValidations');
      // (After conditionalFormatting, before hyperlinks and what follows them: CT_Worksheet's order.)
      const AFTER = ['hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
      root.insertBefore(wrap, [...root.childNodes].find((n) => n.nodeType === 1 && AFTER.includes(n.localName)) || null);
    }
    const prefix = root.prefix ? `${root.prefix}:` : '';
    const xml = validationXml(rule, sel, date1904).replace(/<(\/?)(dataValidation|formula1|formula2)\b/g, `<$1${prefix}$2`);
    const frag = parseXml(`<w xmlns${root.prefix ? `:${root.prefix}` : ''}="${ns}">${xml}</w>`).documentElement.firstChild;
    wrap.appendChild(doc.importNode(frag, true));
  }
  if (wrap) {
    const n = [...wrap.childNodes].filter((x) => x.localName === 'dataValidation').length;
    if (n) wrap.setAttribute('count', String(n)); else wrap.parentNode.removeChild(wrap);
  }
}

/** A rule as the sheet holds it (sheet.js readValidations) in the box's terms, its formulas as Excel shows them in its
 *  box: a typed list unquoted, a reference or formula with "=", a date or time as one. null for none. */
export function boxRule(v, date1904 = false) {
  if (!v) return null;
  const type = Object.keys(TYPES).find((k) => TYPES[k] === v.type) || 'any';
  const op = Object.keys(OPERATORS).find((k) => OPERATORS[k] === v.operator) || 'between';
  const shown = (f) => {
    if (f == null) return '';
    if (type === 'list' && /^".*"$/.test(f)) return f.slice(1, -1).replace(/""/g, '"');
    if (type === 'custom' || !/^-?[\d.]+(E[+-]?\d+)?$/i.test(f)) return `=${f}`;
    const n = Number(f);
    if (type === 'date') { const d = new Date(Date.UTC(date1904 ? 1904 : 1899, date1904 ? 0 : 11, date1904 ? 1 : 30) + n * 86400000); return `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${d.getUTCFullYear()}`; }
    if (type === 'time') { const m = Math.round(n * 1440); return `${Math.floor(m / 60)}:${String(m % 60).padStart(2, '0')}`; }
    return f;
  };
  return {
    type, op, f1: shown(v.formulas[0]), f2: shown(v.formulas[1]), ignoreBlank: v.allowBlank, dropDown: v.dropDown,
    prompt: { show: v.showPrompt, title: v.promptTitle, text: v.prompt }, error: { show: v.showError, style: v.errorStyle, title: v.errorTitle, text: v.error },
  };
}
