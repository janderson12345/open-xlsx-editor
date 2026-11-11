// Conditional formatting as Excel draws it (measured over COM: FormatConditions, then each cell's Range.DisplayFormat;
// docs/CONDFORMAT.md). The rules of a sheet go in priority order (1 first); a rule that holds for a cell gives it what
// its format sets and a rule before it hasn't, and Stop If True ends the cell's rules there. Cell value rules compare
// as Excel compares (a blank as 0 or "", text after every number, case aside, errors never); text, blank and error
// rules are the formulas Excel writes for them; top and bottom, average, duplicates and colour scales look at the
// range's values; data bars and icon sets by Excel's documented rules (cfbars.js).
import { shiftFormula } from '../calc/tokens.js';
import { resolveColor } from './theme.js';
import { MAX_ROWS, MAX_COLS } from './refs.js';
import { barSpans, iconIndexes } from './cfbars.js';

const isError = (v) => v && typeof v === 'object' && v.code != null;
/** A cell's value as a rule sees it: null for a blank. */
function valueOf(cell) {
  if (!cell || cell.v == null || cell.v === '') return (cell?.f != null && cell.f !== '' && cell.v === '') ? '' : null;
  if (cell.t === 'e') return { code: String(cell.v) };
  return cell.v;
}
/** A formula's result as one value (the first of an array or range). */
function scalar(v) {
  if (v == null) return null;
  if (Array.isArray(v)) return scalar(v[0]);
  if (typeof v === 'object' && !isError(v)) {
    if (typeof v.get === 'function') return scalar(v.get(0, 0));
    if (Array.isArray(v.rows)) return scalar(v.rows[0]);
    if ('value' in v) return v.value;
  }
  return v;
}
const rank = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
/** Excel's comparison of two values: below 0, 0, above 0 (a blank as 0 against a number, as "" against text); null
 *  when either is an error. */
export function compare(a, b) {
  if (isError(a) || isError(b)) return null;
  if (a == null) a = typeof b === 'string' ? '' : 0;
  if (b == null) b = typeof a === 'string' ? '' : 0;
  if (rank(a) !== rank(b)) return rank(a) - rank(b);
  if (typeof a === 'string') { const x = a.toLowerCase(); const y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
  return a === b ? 0 : a < b ? -1 : 1;
}
const truthy = (v) => (typeof v === 'boolean' ? v : typeof v === 'number' ? v !== 0 : false);

/** The cells of a rule's ranges, cut to what the sheet uses (a whole column needn't be walked to its end). */
function* cellsOf(sheet, ranges) {
  for (const a of ranges) {
    const r1 = Math.min(a.r1, Math.max(sheet.maxRow, a.r0));
    const c1 = Math.min(a.c1, Math.max(sheet.maxCol, a.c0));
    for (let r = a.r0; r <= r1 && r < MAX_ROWS; r++) for (let c = a.c0; c <= c1 && c < MAX_COLS; c++) yield [r, c];
  }
}

/** A colour scale's colour for each numeric cell of its ranges ("#RRGGBB"): each value placed between the scale's
 *  points (lowest, a percentile or number, highest), the colour mixed channel by channel. */
function scaleColors(wb, sheet, ranges, scale) {
  const theme = wb.theme;
  const cells = [...cellsOf(sheet, ranges)].map(([r, c]) => [r, c, sheet.cells.get(r)?.get(c)?.v]).filter(([, , v]) => typeof v === 'number');
  if (!cells.length) return new Map();
  const values = cells.map(([, , v]) => v).sort((a, b) => a - b);
  const lo = values[0];
  const hi = values[values.length - 1];
  const percentile = (p) => { const k = (values.length - 1) * p / 100; const i = Math.floor(k); return values[i] + (values[Math.min(i + 1, values.length - 1)] - values[i]) * (k - i); };
  const point = (v) => {
    const n = Number(v.val);
    switch (v.type) {
      case 'min': return lo;
      case 'max': return hi;
      case 'percent': return lo + (hi - lo) * n / 100;
      case 'percentile': return percentile(n);
      default: return n;
    }
  };
  const points = scale.cfvo.map(point);
  const rgb = scale.colors.map((c) => { const h = resolveColor(c, theme, wb.styles.palette) || '#FFFFFF'; return [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16)); });
  const out = new Map();
  for (const [r, c, v] of cells) {
    let k = 0;
    while (k < points.length - 2 && v > points[k + 1]) k++;
    const span = points[k + 1] - points[k];
    const t = span > 0 ? Math.min(1, Math.max(0, (v - points[k]) / span)) : (v >= points[k + 1] ? 1 : 0);
    // (Each channel moved from the lower point's by a whole step, the fraction cut off: Excel's three-colour scale
    // gives 250 for 248 + 2.8 and 132 for 132 - 0.8.)
    const mix = rgb[k].map((x, i) => x + Math.trunc((rgb[k + 1][i] - x) * t));
    out.set(`${r},${c}`, `#${mix.map((x) => x.toString(16).padStart(2, '0').toUpperCase()).join('')}`);
  }
  return out;
}

// (Blanks and errors are judged by the cell, not by the formula Excel writes for them: an error is not blank, though
// LEN(TRIM()) of it is an error. Text of spaces only is blank.)
const blankLike = (v) => v == null || (typeof v === 'string' && v.trim() === '');
const BY_CELL = { containsBlanks: blankLike, notContainsBlanks: (v) => !blankLike(v), containsErrors: isError, notContainsErrors: (v) => !isError(v) };

/** The test of a rule that looks at its whole range (top and bottom, average, duplicates): a function of a value. */
function rangeTest(sheet, ranges, rule) {
  const values = [...cellsOf(sheet, ranges)].map(([r, c]) => valueOf(sheet.cells.get(r)?.get(c)));
  if (rule.type === 'top10') {
    const nums = values.filter((v) => typeof v === 'number').sort((a, b) => (rule.bottom ? a - b : b - a));
    const n = rule.percent ? Math.max(1, Math.floor(nums.length * rule.rank / 100)) : rule.rank;
    if (!nums.length) return () => false;
    const edge = nums[Math.min(n, nums.length) - 1];
    return (v) => typeof v === 'number' && (rule.bottom ? v <= edge : v >= edge);
  }
  if (rule.type === 'aboveAverage') {
    const nums = values.filter((v) => typeof v === 'number');
    const avg = nums.reduce((a, b) => a + b, 0) / (nums.length || 1);
    return (v) => typeof v === 'number' && (rule.aboveAverage ? v > avg || (rule.equalAverage && v === avg) : v < avg || (rule.equalAverage && v === avg));
  }
  // (Duplicates and unique values: alike but for case; blanks left out; an error is a value of its own.)
  const key = (v) => (v == null ? null : isError(v) ? `e:${v.code}` : typeof v === 'string' ? `s:${v.toLowerCase()}` : `${typeof v}:${v}`);
  const counts = new Map();
  for (const v of values) { const k = key(v); if (k != null) counts.set(k, (counts.get(k) || 0) + 1); }
  const dup = rule.type === 'duplicateValues';
  return (v) => { const k = key(v); return k != null && (dup ? counts.get(k) > 1 : counts.get(k) === 1); };
}

/**
 * What sheet si's conditional formats give its cells: Map("r,c" => { fill, color ("#RRGGBB"), bold, italic, strike,
 * underline, bar: { from, to, color, gradient, border, axis, axisColor }, icon: { set, index }, hideValue }), only what a
 * rule sets.
 */
export function conditionalLooks(wb, si) {
  const sheet = wb.sheet(si);
  const out = new Map();
  if (!sheet.condFormats?.length) return out;
  const rules = sheet.condFormats.flatMap((cf) => cf.rules.map((rule) => ({ rule, ranges: cf.ranges }))).sort((a, b) => a.rule.priority - b.rule.priority);
  const stopped = new Set();
  const theme = wb.theme;
  const color = (c) => resolveColor(c, theme, wb.styles.palette);
  const give = (key, dxf) => {
    const look = out.get(key) || {};
    if (dxf.fill && look.fill === undefined) look.fill = color(dxf.fill);
    const f = dxf.font;
    if (f) {
      if (f.color && look.color === undefined) look.color = color(f.color);
      if (f.b !== undefined && look.bold === undefined) look.bold = f.b;
      if (f.i !== undefined && look.italic === undefined) look.italic = f.i;
      if (f.strike !== undefined && look.strike === undefined) look.strike = f.strike;
      if (f.u !== undefined && look.underline === undefined) look.underline = f.u;
    }
    out.set(key, look);
  };
  for (const { rule, ranges } of rules) {
    if (!ranges.length) continue;
    if (rule.type === 'colorScale' && rule.colorScale) {
      for (const [key, fill] of scaleColors(wb, sheet, ranges, rule.colorScale)) {
        if (stopped.has(key)) continue;
        const look = out.get(key) || {};
        if (look.fill === undefined) look.fill = fill;
        out.set(key, look);
      }
      continue;
    }
    if ((rule.type === 'dataBar' && rule.dataBar) || (rule.type === 'iconSet' && rule.iconSet)) {
      // (A cell's first data bar and first icon set, by priority; cfbars.js works them out.)
      const cells = [...cellsOf(sheet, ranges)].map(([r, c]) => [r, c, sheet.cells.get(r)?.get(c)?.v]).filter(([r, c]) => !stopped.has(`${r},${c}`));
      const bar = rule.dataBar;
      const found = bar ? barSpans(cells, bar) : iconIndexes(cells, rule.iconSet);
      for (const [key, got] of found) {
        const look = out.get(key) || {};
        if (bar && look.bar === undefined) {
          const neg = got.negative && !bar.negativeSame;
          look.bar = { ...got, color: color(neg ? bar.negativeFill || bar.fill : bar.fill) || '#638EC6', gradient: bar.gradient, border: bar.border ? color(neg ? bar.negativeBorder || bar.borderColor : bar.borderColor) || null : null, axisColor: color(bar.axisColor) || '#000000' };
          if (bar.showValue === false) look.hideValue = true;
        } else if (!bar && look.icon === undefined) {
          look.icon = { set: rule.iconSet.set, index: got };
          if (rule.iconSet.showValue === false) look.hideValue = true;
        }
        out.set(key, look);
      }
      continue;
    }
    const dxf = rule.dxfId == null ? null : wb.styles.dxfs?.[rule.dxfId];
    const base = ranges[0];
    const whole = ['top10', 'aboveAverage', 'duplicateValues', 'uniqueValues'].includes(rule.type) ? rangeTest(sheet, ranges, rule) : null;
    const at = (r, c, i) => scalar(wb.calc.formulaAt(si, r, c, shiftFormula(rule.formulas[i], r - base.r0, c - base.c0)));
    for (const [r, c] of cellsOf(sheet, ranges)) {
      const key = `${r},${c}`;
      if (stopped.has(key)) continue;
      const v = valueOf(sheet.cells.get(r)?.get(c));
      let holds = false;
      if (whole) holds = whole(v);
      else if (rule.type === 'cellIs') {
        const a = at(r, c, 0);
        const d = compare(v, a);
        if (d != null) {
          switch (rule.operator) {
            case 'between': case 'notBetween': {
              const b = at(r, c, 1);
              const lo = compare(a, b) <= 0 ? a : b;
              const hi = lo === a ? b : a;
              const inside = compare(v, lo) >= 0 && compare(v, hi) <= 0;
              holds = rule.operator === 'between' ? inside : !inside;
              break;
            }
            case 'equal': holds = d === 0; break;
            case 'notEqual': holds = d !== 0; break;
            case 'greaterThan': holds = d > 0; break;
            case 'lessThan': holds = d < 0; break;
            case 'greaterThanOrEqual': holds = d >= 0; break;
            case 'lessThanOrEqual': holds = d <= 0; break;
            default: holds = false;
          }
        }
      } else if (BY_CELL[rule.type]) holds = BY_CELL[rule.type](v);
      else if (rule.formulas.length) holds = truthy(at(r, c, 0));
      if (!holds) continue;
      if (dxf) give(key, dxf);
      if (rule.stopIfTrue) stopped.add(key);
    }
  }
  for (const [k, look] of out) if (!Object.keys(look).length) out.delete(k);
  return out;
}
