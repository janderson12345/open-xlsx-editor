// Data validation as Excel judges it (measured over COM: Validation.Add, then each cell's Range.Validation.Value;
// docs/VALIDATION.md). A blank passes when the rule ignores blanks; an error never does. Whole number, decimal, date
// and time rules take numbers (a formula's result too), not text, not even a number stored as text, and compare as
// Excel compares (TRUE is above every number); a whole number must have no fraction. Text length is of the shown
// text. A list typed into the rule matches case and leaves spaces round the value aside; a list from cells matches
// case aside. A custom rule is its formula, relative to the range's first cell.
import { shiftFormula } from '../calc/tokens.js';
import { compare } from './condformat.js';

/** Excel's words when a value is refused and the rule has none of its own. */
export const VALIDATION_MESSAGE = "This value doesn't match the data validation restrictions defined for this cell.";

const isError = (v) => v && typeof v === 'object' && v.code != null;
/** A formula's result as one value. */
function scalar(v) {
  if (v == null) return null;
  if (Array.isArray(v)) return scalar(v[0]);
  if (typeof v === 'object' && !isError(v)) {
    if (typeof v.get === 'function') return scalar(v.get(0, 0));
    if ('value' in v) return v.value;
  }
  return v;
}
/** A number or other value as text, as a General cell shows it. */
const asText = (v) => (typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : typeof v === 'number' ? String(Number(v.toPrecision(15))) : String(v));

/** The rule of sheet si that covers cell (r, c), or null. */
export function ruleAt(sheet, r, c) {
  return sheet.validations?.find((v) => v.type !== 'none' && v.ranges.some((a) => r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1)) || null;
}

/** A rule's formula i worked out for cell (r, c) of sheet si (`array`: a range as a whole). */
function formulaValue(wb, si, rule, i, r, c, array = false) {
  const base = rule.ranges[0];
  return wb.calc.formulaAt(si, r, c, shiftFormula(rule.formulas[i], r - base.r0, c - base.c0), array);
}

/**
 * A list rule's choices for cell (r, c): [{ text, value }] in order, from the rule's own text ("a,b,c") or the cells
 * (or name) it points to; `typed` tells which.
 */
export function listItems(wb, si, rule, r, c) {
  const f = rule.formulas[0] || '';
  if (/^".*"$/.test(f)) return { typed: true, items: f.slice(1, -1).replace(/""/g, '"').split(',').map((t) => ({ text: t, value: t })) };
  const got = formulaValue(wb, si, rule, 0, r, c, true);
  const out = [];
  const take = (v) => { if (v != null && v !== '') out.push({ text: asText(v), value: v }); };
  // (A range of cells, read in order across each row; rows of values; or one value.)
  if (got && got.r0 != null && got.c0 != null && got.si != null) {
    const sheet = wb.sheet(got.si);
    for (let i = got.r0; i <= got.r1; i++) for (let j = got.c0; j <= got.c1; j++) { const cell = sheet.cells.get(i)?.get(j); take(cell?.t === 'e' ? null : cell?.v); }
  } else if (Array.isArray(got)) got.flat().forEach((x) => take(scalar(x)));
  else if (got && Array.isArray(got.rows)) got.rows.flat().forEach((x) => take(scalar(x)));
  else take(scalar(got));
  return { typed: false, items: out };
}

/** Whether value `v` (a number, text, boolean, { code } or null) may go in cell (r, c) of sheet si under `rule`. */
export function valid(wb, si, rule, r, c, v) {
  if (v == null || v === '') return rule.allowBlank;
  if (isError(v)) return false;
  const bound = (i) => scalar(formulaValue(wb, si, rule, i, r, c));
  // (Numbers compared to 15 significant digits, as Excel keeps them: 17:00 is saved as 0.708333333333333.)
  const fix = (x) => (typeof x === 'number' ? Number(x.toPrecision(15)) : x);
  const within = (y) => {
    const x = fix(y);
    const a = fix(bound(0));
    const d = compare(x, a);
    if (d == null) return false;
    switch (rule.operator) {
      case 'between': case 'notBetween': {
        const b = bound(1);
        const inside = d >= 0 && compare(x, fix(b)) <= 0;
        return rule.operator === 'between' ? inside : !inside;
      }
      case 'equal': return d === 0;
      case 'notEqual': return d !== 0;
      case 'greaterThan': return d > 0;
      case 'lessThan': return d < 0;
      case 'greaterThanOrEqual': return d >= 0;
      case 'lessThanOrEqual': return d <= 0;
      default: return false;
    }
  };
  switch (rule.type) {
    case 'whole': case 'decimal': case 'date': case 'time':
      if (typeof v === 'string') return false;
      if (rule.type === 'whole' && typeof v === 'number' && !Number.isInteger(v)) return false;
      return within(v);
    case 'textLength': return within(asText(v).length);
    case 'list': {
      const { typed, items } = listItems(wb, si, rule, r, c);
      const text = asText(v);
      if (typed) return items.some((it) => it.text.trim() === text.trim());
      return items.some((it) => (typeof it.value === 'number' && typeof v === 'number' ? it.value === v : it.text.toLowerCase() === text.toLowerCase()));
    }
    case 'custom': {
      const got = scalar(formulaValue(wb, si, rule, 0, r, c));
      return got === true || (typeof got === 'number' && got !== 0);
    }
    default: return true;
  }
}

/** Data > Data Validation > Circle Invalid Data: Excel draws a circle round each cell of sheet si whose value breaks
 *  its rule, at most 255 (measured over COM: CircleInvalid on 300 broken cells made 255). Blank cells count when the
 *  rule doesn't ignore blanks; a merged area by its top-left cell. Rule by rule, each range row by row (Excel's own
 *  order not measured). [{ r, c }] */
export function invalidCells(wb, si, limit = 255) {
  const sheet = wb.sheet(si);
  const out = [];
  const seen = new Set();
  const inMerge = (r, c) => sheet.merges?.some((m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1 && (r !== m.r0 || c !== m.c0));
  for (const rule of sheet.validations || []) {
    if (rule.type === 'none') continue;
    for (const a of rule.ranges) {
      // (Blanks pass a rule that ignores them, so only the cells there are looked at; else every cell, up to the
      // limit.)
      const last = rule.allowBlank ? Math.min(a.r1, sheet.maxRow) : a.r1;
      for (let r = a.r0; r <= last && out.length < limit; r++) {
        const line = sheet.cells.get(r);
        if (rule.allowBlank && !line) continue;
        const cols = rule.allowBlank ? [...line.keys()].filter((c) => c >= a.c0 && c <= a.c1).sort((x, y) => x - y) : Array.from({ length: a.c1 - a.c0 + 1 }, (_, k) => a.c0 + k);
        for (const c of cols) {
          if (out.length >= limit) break;
          const key = `${r},${c}`;
          if (seen.has(key) || inMerge(r, c) || ruleAt(sheet, r, c) !== rule) continue;
          seen.add(key);
          if (breaksRule(wb, si, r, c)) out.push({ r, c });
        }
      }
    }
  }
  return out;
}

/** Whether cell (r, c) of sheet si breaks the rule over it (none: false). */
export function breaksRule(wb, si, r, c) {
  const sheet = wb.sheet(si);
  const rule = ruleAt(sheet, r, c);
  if (!rule) return false;
  const cell = sheet.cells.get(r)?.get(c);
  return !valid(wb, si, rule, r, c, !cell ? null : cell.t === 'e' ? { code: String(cell.v) } : cell.v);
}
