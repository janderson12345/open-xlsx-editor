// Data validation rules written as Excel's Data Validation box writes them (measured: test/dv-excel.json holds the
// dataValidation element Excel saved for each of 15 rules; this writes the same, but for Excel's xr:uid). A rule for
// the selection takes its cells out of any rule they had (those rules' ranges split around them), then is added after
// the others; Clear All only takes them out. A rule whose formulas read another sheet goes in Excel's 2010 form, as
// Excel writes it (excel-dv-x14.ps1, test/dv-x14-excel.json); rules of that form are split, moved and added to too.
import { parseInput } from './input.js';
import { parseRange, rangeName, MAX_ROWS, MAX_COLS } from './refs.js';
import { parseXml, unescapeXml } from './xml.js';
import { shiftFormula, tokenize } from '../calc/tokens.js';
import { rulesOf, dropRule } from './rulemoves.js';
import { EXT_CF } from './cfext.js';

const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const XM = 'http://schemas.microsoft.com/office/excel/2006/main';
/** The ext of Excel 2010's data validation rules. */
const EXT_DV = '{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}';

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
  // (Looked up as own keys and the error style checked: a rule read from a file brings its own text for these, which
  // would otherwise be written into the attribute as it came, quotes and all.)
  const type = Object.hasOwn(TYPES, rule.type) ? TYPES[rule.type] : null;
  const op = Object.hasOwn(OPERATORS, rule.op) ? OPERATORS[rule.op] : 'between';
  const a = [];
  if (type) a.push(`type="${type}"`);
  if (rule.error?.style === 'warning' || rule.error?.style === 'information') a.push(`errorStyle="${rule.error.style}"`);
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

const sqrefOf = (rule) => rule.get().split(/\s+/).map(parseRange).filter(Boolean);

/** A rule whose first cell moved from `was` to `now` (its sqref changed): its formulas, written for its first cell,
 *  shifted so that each cell's still means the same, as Excel rewrote C2:C4's =C2>B2 as =C4>B4 when C2:C3 were cut. */
function rebase(rule, was, now) {
  if (!was || !now || (was.r0 === now.r0 && was.c0 === now.c0)) return;
  for (const f of rule.formulas()) f.textContent = shiftFormula(f.textContent, now.r0 - was.r0, now.c0 - was.c0);
}

/** The rules' count written, the list gone once empty. */
function recount(wrap) {
  if (!wrap || !wrap.parentNode) return;
  const n = [...wrap.childNodes].filter((x) => x.localName === 'dataValidation').length;
  if (n) wrap.setAttribute('count', String(n)); else wrap.parentNode.removeChild(wrap);
}

/** A rule's new ranges written, its formulas rebased; the rule gone when there are none. */
function reshape(rule, was, now) {
  if (!now.length) { dropRule(rule); return; }
  rule.set(now.map(rangeName).join(' '));
  rebase(rule, was[0], now[0]);
}

/** A cut of `rect` moved `dr` rows and `dc` columns on its sheet, as Excel moves validation with it (measured over
 *  COM): each rule's cells there go with them, after its other cells (C2:C4 with C2:C3 cut to F5: C4 F5:F6); the cells
 *  pasted over leave the rules they were in; a rule left with none goes; formulas rebased. Rules of both forms (the
 *  Excel 2010 one in the extLst too). Edits `doc`. */
export function moveValidation(doc, rect, dr, dc) {
  const area = { r0: rect.r0 + dr, c0: rect.c0 + dc, r1: rect.r1 + dr, c1: rect.c1 + dc };
  const meet = (a, b) => (a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1 ? { r0: Math.max(a.r0, b.r0), c0: Math.max(a.c0, b.c0), r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1) } : null);
  for (const rule of rulesOf(doc)) {
    const was = sqrefOf(rule);
    const stay = was.flatMap((p) => subtract(p, rect)).flatMap((p) => subtract(p, area));
    const moved = coalesce(was.map((p) => meet(p, rect)).filter(Boolean).map((p) => ({ r0: p.r0 + dr, c0: p.c0 + dc, r1: p.r1 + dr, c1: p.c1 + dc })));
    reshape(rule, was, [...stay, ...moved]);
  }
}

/** Whether a rule's stored formulas read another sheet than `sheet` (then Excel writes it in its 2010 form). */
function readsOtherSheet(formulas, sheet) {
  return formulas.some((f) => {
    if (!f) return false;
    try {
      return tokenize(f).some((t) => t.type === 'ref' && t.sheet != null && (sheet == null || String(t.sheet).toLowerCase() !== String(sheet).toLowerCase()));
    } catch { return false; }
  });
}

/** A rule put in the worksheet's extLst as Excel writes one of its 2010 form: an x14:dataValidation (the main one's
 *  attributes, its formulas as xm:f, its ranges as xm:sqref) in the ext {CCE6A557-...}, after the conditional formats'
 *  ext and before the others (Excel's order, excel-dv-x14.ps1), made when there is none. */
function addExtRule(doc, xml) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const p = root.prefix ? `${root.prefix}:` : '';
  let list = [...root.childNodes].find((n) => n.nodeType === 1 && n.localName === 'extLst');
  if (!list) { list = doc.createElementNS(ns, `${p}extLst`); root.appendChild(list); }
  const exts = [...list.childNodes].filter((n) => n.nodeType === 1 && n.localName === 'ext');
  let ext = exts.find((e) => e.getAttribute('uri') === EXT_DV);
  if (!ext) {
    ext = doc.importNode(parseXml(`<w xmlns="${ns}"><${p}ext uri="${EXT_DV}" xmlns:x14="${X14}"><x14:dataValidations count="0" xmlns:xm="${XM}"/></${p}ext></w>`).documentElement.firstChild, true);
    const cf = exts.find((e) => e.getAttribute('uri') === EXT_CF);
    list.insertBefore(ext, cf ? cf.nextSibling : list.firstChild);
  }
  const wrap = [...ext.childNodes].find((n) => n.nodeType === 1 && n.localName === 'dataValidations');
  wrap.appendChild(doc.importNode(parseXml(`<w xmlns:x14="${X14}" xmlns:xm="${XM}">${xml}</w>`).documentElement.firstChild, true));
  wrap.setAttribute('count', String([...wrap.childNodes].filter((n) => n.nodeType === 1).length));
}

/** validationXml's rule in Excel's 2010 form: the same attributes less sqref, x14:formula1 and 2 holding xm:f, then
 *  xm:sqref (excel-dv-x14.ps1: a list of Data!$A$1:$A$3, whole numbers 1 to Data!$B$1). */
function extXml(main) {
  const sqref = /\ssqref="([^"]*)"/.exec(main)[1];
  return main.replace(/\ssqref="[^"]*"/, '').replace(/<dataValidation\b/, '<x14:dataValidation')
    .replace(/<formula(\d)>([^<]*)<\/formula\d>/g,'<x14:formula$1><xm:f>$2</xm:f></x14:formula$1>')
    .replace(/<\/dataValidation>$/, `<xm:sqref>${sqref}</xm:sqref></x14:dataValidation>`)
    .replace(/\/>$/, `><xm:sqref>${sqref}</xm:sqref></x14:dataValidation>`);
}

/**
 * Give sheet si's cells `ranges` the validation `rule` (null: Clear All), as Excel's box does. Edits the sheet's XML
 * (`doc`); the caller reads the sheet again. `sheet`: the sheet's name, for a rule reading another sheet, which goes in
 * Excel's 2010 form as Excel writes it.
 */
export function setValidation(doc, ranges, rule, date1904 = false, sheet = null) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const sel = ranges.map((r) => ({ r0: r.r0, c0: r.c0, r1: Math.min(r.r1, MAX_ROWS - 1), c1: Math.min(r.c1, MAX_COLS - 1) }));
  // The selection's cells out of the rules they were in, of both forms; a rule left with none goes.
  for (const old of rulesOf(doc)) {
    const was = sqrefOf(old);
    let parts = was;
    for (const s of sel) parts = parts.flatMap((p) => subtract(p, s));
    if (parts.length !== was.length || parts.some((p, k) => rangeName(p) !== rangeName(was[k]))) reshape(old, was, parts);
  }
  let wrap = [...root.childNodes].find((n) => n.localName === 'dataValidations');
  if (rule) {
    const main = validationXml(rule, sel, date1904);
    // (A formula is escaped in validationXml, so it has no '<': [^<]* takes it in one pass.)
    const stored = [...main.matchAll(/<formula\d>([^<]*)<\/formula\d>/g)].map((m) => unescapeXml(m[1]));
    if (readsOtherSheet(stored, sheet)) { addExtRule(doc, extXml(main)); return; }
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
  recount(wrap);
}

/** Ranges joined where one carries on another, as Excel writes a rule's sqref after a paste: below it with the same
 *  columns (P5:P6 and P7:P8 as P5:P8), or beside it with the same rows. */
export function coalesce(ranges) {
  const out = [];
  for (const a of ranges) {
    const b = out.find((x) => (x.c0 === a.c0 && x.c1 === a.c1 && x.r1 + 1 === a.r0) || (x.r0 === a.r0 && x.r1 === a.r1 && x.c1 + 1 === a.c0));
    if (b) { b.r1 = Math.max(b.r1, a.r1); b.c1 = Math.max(b.c1, a.c1); } else out.push({ ...a });
  }
  return out;
}

/** The cells `ranges` added to the sheet's rule whose first cell is `first` ({ r, c }) and whose type and first
 *  formula are `type` and `formula1`, as Excel adds pasted cells to the rule they were copied from (its formulas, as
 *  written for its first cell, cover them too). Edits `doc`; false when there is no such rule. */
export function addToRule(doc, { first, type, formula1 }, ranges) {
  const rule = rulesOf(doc).find((x) => {
    if ((x.el.getAttribute('type') || 'none') !== type) return false;
    const f = x.formulas()[0];
    const a = sqrefOf(x)[0];
    return (f ? f.textContent : null) === (formula1 ?? null) && a && a.r0 === first.r && a.c0 === first.c;
  });
  if (!rule) return false;
  const now = sqrefOf(rule);
  rule.set([...now.slice(0, 1), ...coalesce([...now.slice(1), ...ranges])].map(rangeName).join(' '));
  return true;
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
