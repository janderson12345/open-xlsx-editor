// Lookup and reference, and the conditional sums and counts, as Excel defines them: registered into functions.js's
// table (evaluate.js imports this file after it).
import { Area, Arr, MISSING, Unsupported, XlError, E, compare, textToNumber } from './values.js';
import { tokenize } from './tokens.js';
import { def, given, one, num, text, opt, cellsIn, grid, vector, bool } from './functions.js';
import { wildcard } from './functions-text.js';

// ---- Lookup and reference

const sameType = (a, b) => typeof a === typeof b && !(a instanceof XlError) && !(b instanceof XlError);
/** A test for an exact match: same type, text without regard to case, wildcards in text if `wild`. */
function exactly(x, wild) {
  if (wild && typeof x === 'string' && /[*?~]/.test(x)) {
    const re = wildcard(x);
    return (v) => typeof v === 'string' && re.test(v);
  }
  return (v) => v !== null && sameType(v, x) && compare(v, x) === 0;
}
/**
 * Position of x in a list, as Excel finds it (fitted on Excel's answers for 400 unsorted lists, with blanks and text
 * among the numbers: docs/CALC.md); -1 if none. `how` 0: the first exact match. 1 (a list sorted up: MATCH 1, VLOOKUP
 * and HLOOKUP TRUE): Excel's binary search, which on a list that isn't sorted gives its own answer: from the middle,
 * a blank or a value of another type passed over to the right; an exact match taken at once, then the last of the
 * equal values that follow it. -1 (sorted down: MATCH -1): a scan from the start, blanks and other types passed over,
 * that takes the first exact match and stops at the first value below x.
 */
export function position(list, x, how, wild = true) {
  if (how === 0) {
    const eq = exactly(x, wild);
    for (let i = 0; i < list.length; i++) if (eq(list.at(i))) return i;
    return -1;
  }
  const fits = (v) => v !== null && sameType(v, x);
  if (how < 0) {
    let found = -1;
    for (let i = 0; i < list.length; i++) {
      const v = list.at(i);
      if (!fits(v)) continue;
      const k = compare(v, x);
      if (k === 0) return i;
      if (k < 0) break;
      found = i;
    }
    return found;
  }
  let lo = 0;
  let hi = list.length - 1;
  let found = -1;
  while (lo <= hi) {
    let m = (lo + hi) >> 1;
    while (m <= hi && !fits(list.at(m))) m++;
    if (m > hi) { hi = ((lo + hi) >> 1) - 1; continue; }
    const k = compare(list.at(m), x);
    if (k === 0) {
      // (The last of the equal values after it, blanks and other types passed over.)
      let last = m;
      for (let j = m + 1; j < list.length; j++) {
        const v = list.at(j);
        if (!fits(v)) continue;
        if (compare(v, x) === 0) last = j; else break;
      }
      return last;
    }
    if (k < 0) { found = m; lo = m + 1; } else hi = m - 1;
  }
  return found;
}
const lookupValue = (v, ctx) => { const x = one(v, ctx); return x === null ? 0 : x; };

function tableLookup(a, ctx, byRow) {
  const x = lookupValue(a[0], ctx);
  const g = grid(a[1], ctx);
  const k = Math.trunc(num(a[2], ctx));
  const approx = given(a[3]) ? bool(a[3], ctx) : a[3] !== MISSING;
  // (No table, its place left empty: #N/A before the column is looked at, as Excel gives =VLOOKUP(,,) and =VLOOKUP(2,,1).)
  if (a[1] === MISSING) throw E.na;
  if (k < 1) throw E.value;
  if (k > (byRow ? g.rows : g.cols)) throw E.ref;
  const list = byRow ? { length: g.cols, at: (i) => g.at(0, i) } : { length: Math.min(g.rows, Math.max(g.used, 1)), at: (i) => g.at(i, 0) };
  const i = position(list, x, approx ? 1 : 0);
  if (i < 0) throw E.na;
  return byRow ? g.at(k - 1, i) : g.at(i, k - 1);
}
def('VLOOKUP', (a, ctx) => tableLookup(a, ctx, false));
def('HLOOKUP', (a, ctx) => tableLookup(a, ctx, true));
def('MATCH', (a, ctx) => {
  const x = lookupValue(a[0], ctx);
  const list = vector(a[1], ctx);
  if (!list) throw E.na;
  const how = Math.sign(opt(a, 2, ctx, 1));
  const i = position(list, x, how);
  if (i < 0) throw E.na;
  return i + 1;
});
def('INDEX', (a, ctx) => {
  const src = a[0];
  let r = given(a[1]) ? Math.trunc(num(a[1], ctx)) : 0;
  let c = given(a[2]) ? Math.trunc(num(a[2], ctx)) : 0;
  const rows = src instanceof Area ? src.rows : src instanceof Arr ? src.height : 1;
  const cols = src instanceof Area ? src.cols : src instanceof Arr ? src.width : 1;
  // (One index into a single row picks a column: INDEX(A1:E1, 3) is C1.)
  if (a.length === 2 && rows === 1) { c = r; r = 0; if (cols === 1) c = 0; }
  if (r < 0 || c < 0) throw E.value;
  if (r > rows || c > cols) throw E.ref;
  if (src instanceof Area) {
    return new Area(src.si, r ? src.r0 + r - 1 : src.r0, c ? src.c0 + c - 1 : src.c0, r ? src.r0 + r - 1 : src.r1, c ? src.c0 + c - 1 : src.c1);
  }
  if (src instanceof Arr) {
    if (r && c) return src.rows[r - 1][c - 1];
    if (r) return new Arr([src.rows[r - 1]]);
    if (c) return new Arr(src.rows.map((row) => [row[c - 1]]));
    return src;
  }
  return one(src, ctx);
});
def('XLOOKUP', (a, ctx) => {
  const x = lookupValue(a[0], ctx);
  const list = vector(a[1], ctx);
  if (!list) throw E.value;
  const mode = opt(a, 4, ctx, 0);
  const search = opt(a, 5, ctx, 1);
  const order = [];
  for (let i = 0; i < list.length; i++) order.push(i);
  if (search < 0) order.reverse();
  let found = -1;
  if (mode === 0 || mode === 2) {
    const eq = exactly(x, mode === 2);
    found = order.find((i) => eq(list.at(i))) ?? -1;
  } else {
    let best = null;
    for (const i of order) {
      const v = list.at(i);
      if (v === null || !sameType(v, x)) continue;
      const k = compare(v, x);
      if (k === 0) { found = i; break; }
      if ((mode < 0 ? k < 0 : k > 0) && (best == null || (mode < 0 ? compare(v, list.at(best)) > 0 : compare(v, list.at(best)) < 0))) best = i;
    }
    if (found < 0 && best != null) found = best;
  }
  if (found < 0) { if (a[3] !== undefined && a[3] !== MISSING) return a[3]; throw E.na; }
  const ret = a[2];
  if (ret instanceof Area) {
    return list.vertical
      ? new Area(ret.si, ret.r0 + found, ret.c0, ret.r0 + found, ret.c1)
      : new Area(ret.si, ret.r0, ret.c0 + found, ret.r1, ret.c0 + found);
  }
  const g = grid(ret, ctx);
  return list.vertical ? g.at(found, 0) : g.at(0, found);
});
def('CHOOSE', (a, ctx) => {
  const i = Math.trunc(num(a[0], ctx));
  if (i < 1 || i >= a.length) throw E.value;
  return a[i] === MISSING ? 0 : a[i];
});
def('ROW', (a, ctx) => (a[0] instanceof Area ? a[0].r0 : ctx.here.r) + 1);
def('COLUMN', (a, ctx) => (a[0] instanceof Area ? a[0].c0 : ctx.here.c) + 1);
def('ROWS', (a, ctx) => (a[0] instanceof Area ? a[0].rows : a[0] instanceof Arr ? a[0].height : (one(a[0], ctx), 1)));
def('COLUMNS', (a, ctx) => (a[0] instanceof Area ? a[0].cols : a[0] instanceof Arr ? a[0].width : (one(a[0], ctx), 1)));
def('OFFSET', (a, ctx) => {
  const ref = a[0];
  if (!(ref instanceof Area)) throw E.value;
  const r0 = ref.r0 + Math.trunc(num(a[1], ctx));
  const c0 = ref.c0 + Math.trunc(num(a[2], ctx));
  const h = Math.trunc(opt(a, 3, ctx, ref.rows));
  const w = Math.trunc(opt(a, 4, ctx, ref.cols));
  if (h < 1 || w < 1 || r0 < 0 || c0 < 0 || r0 + h > 1048576 || c0 + w > 16384) throw E.ref;
  return new Area(ref.si, r0, c0, r0 + h - 1, c0 + w - 1);
}, { volatile: true });
def('INDIRECT', (a, ctx) => {
  if (given(a[1]) && !bool(a[1], ctx)) throw new Unsupported('R1C1 references');
  let toks;
  try { toks = tokenize(text(a[0], ctx)); } catch { throw E.ref; }
  const t = toks[0];
  if (toks.length !== 1 || t.type !== 'ref' || !t.area || t.book != null) throw E.ref;
  const si = t.sheet == null ? ctx.here.si : ctx.sheetIndex(t.sheet);
  if (si < 0) throw E.ref;
  return new Area(si, t.area.r0, t.area.c0, t.area.r1, t.area.c1);
}, { volatile: true });

// ---- Conditional sums and counts

/** A criterion ("Apple", ">3", "<>", "A*", 12) as a test of a cell's value, as COUNTIF reads one. */
export function criterion(c, ctx) {
  if (c instanceof XlError) return (v) => v === c;
  if (c === null || c === MISSING) c = 0;
  // (An equal-to criterion matches a number stored as text too: COUNTIF(B2:B9, 12) counts "12". The others don't.)
  const equals = (n) => (v) => (typeof v === 'number' ? compare(v, n) === 0 : typeof v === 'string' && textToNumber(v, ctx.date1904) === n);
  if (typeof c === 'number') return equals(c);
  if (typeof c === 'boolean') return (v) => v === c;
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(c);
  const op = m[1] || '=';
  const operand = m[2];
  const holds = (k) => (op === '=' ? k === 0 : op === '<>' ? k !== 0 : op === '<' ? k < 0 : op === '>' ? k > 0 : op === '<=' ? k <= 0 : k >= 0);
  const n = operand.trim() === '' ? null : textToNumber(operand, ctx.date1904);
  if (n != null) {
    if (op === '<>') return (v) => !(typeof v === 'number' && compare(v, n) === 0);
    if (op === '=') return equals(n);
    return (v) => typeof v === 'number' && holds(compare(v, n));
  }
  const upper = operand.toUpperCase();
  if (upper === 'TRUE' || upper === 'FALSE') {
    const b = upper === 'TRUE';
    if (op === '<>') return (v) => v !== b;
    return (v) => typeof v === 'boolean' && holds(compare(v, b));
  }
  if (operand.startsWith('#') && operand.toUpperCase() in ERR_NAMES) {
    const code = operand.toUpperCase();
    return op === '<>' ? (v) => !(v instanceof XlError && v.code === code) : (v) => v instanceof XlError && v.code === code;
  }
  if (op === '=' || op === '<>') {
    let match;
    if (operand === '') match = m[1] ? (v) => v === null : (v) => v === null || v === '';
    else { const re = wildcard(operand); match = (v) => typeof v === 'string' && re.test(v); }
    return op === '<>' ? (v) => !match(v) : match;
  }
  return (v) => typeof v === 'string' && holds(compare(v, operand));
}
const ERR_NAMES = { '#NULL!': 1, '#DIV/0!': 1, '#VALUE!': 1, '#REF!': 1, '#NAME?': 1, '#NUM!': 1, '#N/A': 1 };

/** Visit the offsets (i, j) where every [range, test] pair holds; ranges must have one shape. `onMatch(i, j)`.
 *  Returns how many matches lie past the cells in use (all blank), for counting. */
function eachMatch(pairs, ctx, onMatch) {
  const [first] = pairs;
  for (const [rg] of pairs) {
    if (!(rg instanceof Area)) throw E.value;
    if (rg.rows !== first[0].rows || rg.cols !== first[0].cols) throw E.value;
  }
  const rows = first[0].rows;
  const cols = first[0].cols;
  let ri = 0;
  let ci = 0;
  for (const [rg] of pairs) {
    const ext = ctx.extent(rg.si);
    ri = Math.max(ri, Math.min(rows, ext.maxRow - rg.r0 + 1));
    ci = Math.max(ci, Math.min(cols, ext.maxCol - rg.c0 + 1));
  }
  const blanksMatch = pairs.every(([, test]) => test(null));
  if (!blanksMatch && pairs.length === 1 && ri * ci > 4096) {
    // (Only cells that hold something can match: visit those.)
    const [rg, test] = first;
    for (const [r, c, v] of cellsIn(rg, ctx)) if (test(v)) onMatch(r - rg.r0, c - rg.c0);
    return 0;
  }
  for (let i = 0; i < ri; i++) {
    for (let j = 0; j < ci; j++) {
      if (pairs.every(([rg, test]) => test(ctx.cellValue(rg.si, rg.r0 + i, rg.c0 + j)))) onMatch(i, j);
    }
  }
  return blanksMatch ? rows * cols - Math.max(0, ri) * Math.max(0, ci) : 0;
}
const pairsFrom = (a, from, ctx) => {
  const pairs = [];
  for (let i = from; i + 1 < a.length; i += 2) pairs.push([a[i], criterion(one(a[i + 1], ctx), ctx)]);
  return pairs;
};
/** The range to add up or average for SUMIF and AVERAGEIF: the same shape as the criteria range, from its top-left. */
const sameShape = (range, like) => (range instanceof Area ? new Area(range.si, range.r0, range.c0, range.r0 + like.rows - 1, range.c0 + like.cols - 1) : like);

function gather(target, pairs, ctx) {
  if (!(target instanceof Area)) throw E.value;
  const out = [];
  eachMatch(pairs, ctx, (i, j) => {
    const v = ctx.cellValue(target.si, target.r0 + i, target.c0 + j);
    if (v instanceof XlError) throw v;
    if (typeof v === 'number') out.push(v);
  });
  return out;
}
const countIfs = (a, ctx) => {
  let n = 0;
  const past = eachMatch(pairsFrom(a, 0, ctx), ctx, () => n++);
  return n + past;
};
def('COUNTIF COUNTIFS', countIfs);
def('SUMIF', (a, ctx) => {
  if (!(a[0] instanceof Area)) throw E.value;
  return gather(given(a[2]) ? sameShape(a[2], a[0]) : a[0], pairsFrom(a, 0, ctx).slice(0, 1), ctx).reduce((s, x) => s + x, 0);
});
def('SUMIFS', (a, ctx) => gather(a[0], pairsFrom(a, 1, ctx), ctx).reduce((s, x) => s + x, 0));
const mean = (list) => { if (!list.length) throw E.div0; return list.reduce((s, x) => s + x, 0) / list.length; };
def('AVERAGEIF', (a, ctx) => {
  if (!(a[0] instanceof Area)) throw E.value;
  return mean(gather(given(a[2]) ? sameShape(a[2], a[0]) : a[0], pairsFrom(a, 0, ctx).slice(0, 1), ctx));
});
def('AVERAGEIFS', (a, ctx) => mean(gather(a[0], pairsFrom(a, 1, ctx), ctx)));
def('MAXIFS', (a, ctx) => { const n = gather(a[0], pairsFrom(a, 1, ctx), ctx); return n.length ? Math.max(...n) : 0; });
def('MINIFS', (a, ctx) => { const n = gather(a[0], pairsFrom(a, 1, ctx), ctx); return n.length ? Math.min(...n) : 0; });
