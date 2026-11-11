// Excel 365's dynamic array functions, which give arrays to spill (engine.js, spill.js): SEQUENCE, SORT, SORTBY,
// UNIQUE, FILTER and TRANSPOSE; @ (SINGLE in the file) and D1# (ANCHORARRAY); and LET and LAMBDA, which evaluate.js works
// out itself (registered here so that the engine knows them). Measured against Excel: test/dynamic-excel.json
// (tools/excel-compare/excel-dynamic.ps1). Registered into functions.js's table; evaluate.js imports this file after it.
import { Area, Arr, MISSING, XlError, E, scalar, matrix, toBool } from './values.js';
import { def, given, num } from './functions.js';
import { compareText } from '../ooxml/sortops.js';

const ARRAY = { array: true };

/** A value's place in Excel's sort order: numbers, text (as Sort orders it), FALSE and TRUE, errors; blanks last. */
function rank(v) {
  if (v == null || v === MISSING) return 4;
  if (v instanceof XlError) return 3;
  if (typeof v === 'boolean') return 2;
  if (typeof v === 'string') return 1;
  return 0;
}
/** Two values in Excel's ascending order (below 0, 0 or above); blanks last whichever way (`desc` keeps them last). */
function order(a, b, desc = false) {
  const ra = rank(a);
  const rb = rank(b);
  if (ra === 4 || rb === 4) return ra === rb ? 0 : ra === 4 ? 1 : -1;
  const d = ra !== rb ? ra - rb : ra === 1 ? compareText(a, b) : ra === 3 ? 0 : Number(a) - Number(b);
  return desc ? -d : d;
}
/** Two values the same for UNIQUE: text matching case aside, as Excel's UNIQUE does. */
const key = (v) => (v == null || v === MISSING ? '' : typeof v === 'string' ? `s${v.toLowerCase()}` : v instanceof XlError ? `e${v.code}` : `${typeof v}${v}`);

const rowsOf = (v, ctx) => matrix(v, ctx);
const columns = (m) => m[0].map((_, c) => m.map((row) => row[c]));
/** A whole number of at least 1 from an argument, or an error. */
function count(v, ctx, dflt) {
  if (!given(v)) return dflt;
  const n = Math.trunc(num(v, ctx));
  return n;
}

// SEQUENCE(rows, [columns], [start], [step]).
def('SEQUENCE', (args, ctx) => {
  const rows = count(args[0], ctx, 1);
  const cols = count(args[1], ctx, 1);
  const start = given(args[2]) ? num(args[2], ctx) : 1;
  const step = given(args[3]) ? num(args[3], ctx) : 1;
  if (rows < 1 || cols < 1) return rows === 0 || cols === 0 ? E.calc : E.value;
  return new Arr(Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (__, c) => start + (r * cols + c) * step)));
}, ARRAY);

// SORT(array, [sort_index], [sort_order], [by_col]).
def('SORT', (args, ctx) => {
  const m = rowsOf(args[0], ctx);
  const byCol = given(args[3]) ? toBool(scalar(args[3], ctx)) : false;
  const index = given(args[1]) ? Math.trunc(num(args[1], ctx)) : 1;
  const desc = given(args[2]) ? num(args[2], ctx) === -1 : false;
  if (given(args[2]) && ![1, -1].includes(num(args[2], ctx))) return E.value;
  const lines = byCol ? columns(m) : m.map((row) => [...row]);
  if (index < 1 || index > lines[0].length) return E.value;
  const sorted = lines.map((line, k) => ({ line, k })).sort((a, b) => order(a.line[index - 1], b.line[index - 1], desc) || a.k - b.k).map((x) => x.line);
  return new Arr(byCol ? columns(sorted) : sorted);
}, ARRAY);

// SORTBY(array, by_array1, [sort_order1], ...): by rows (by_arrays a column each) or by columns (a row each).
def('SORTBY', (args, ctx) => {
  const m = rowsOf(args[0], ctx);
  const keys = [];
  for (let k = 1; k < args.length; k += 2) {
    const by = rowsOf(args[k], ctx);
    const desc = given(args[k + 1]) ? num(args[k + 1], ctx) === -1 : false;
    if (given(args[k + 1]) && ![1, -1].includes(num(args[k + 1], ctx))) return E.value;
    if (by[0].length === 1 && by.length === m.length) keys.push({ values: by.map((row) => row[0]), desc, byCol: false });
    else if (by.length === 1 && by[0].length === m[0].length) keys.push({ values: by[0], desc, byCol: true });
    else return E.value;
  }
  if (!keys.length || keys.some((x) => x.byCol !== keys[0].byCol)) return E.value;
  const byCol = keys[0].byCol;
  const lines = byCol ? columns(m) : m.map((row) => [...row]);
  const sorted = lines.map((line, k) => ({ line, k })).sort((a, b) => {
    for (const x of keys) { const d = order(x.values[a.k], x.values[b.k], x.desc); if (d) return d; }
    return a.k - b.k;
  }).map((x) => x.line);
  return new Arr(byCol ? columns(sorted) : sorted);
}, ARRAY);

// UNIQUE(array, [by_col], [exactly_once]): the distinct rows (or columns) in the order they first come.
def('UNIQUE', (args, ctx) => {
  const m = rowsOf(args[0], ctx);
  const byCol = given(args[1]) ? toBool(scalar(args[1], ctx)) : false;
  const once = given(args[2]) ? toBool(scalar(args[2], ctx)) : false;
  const lines = byCol ? columns(m) : m;
  const counts = new Map();
  for (const line of lines) { const k = line.map(key).join('\u0001'); counts.set(k, (counts.get(k) || 0) + 1); }
  const seen = new Set();
  const out = [];
  for (const line of lines) {
    const k = line.map(key).join('\u0001');
    if (seen.has(k) || (once && counts.get(k) > 1)) continue;
    seen.add(k);
    out.push(line);
  }
  if (!out.length) return E.calc;
  return new Arr(byCol ? columns(out) : out);
}, ARRAY);

// FILTER(array, include, [if_empty]): the rows (or columns) whose include is TRUE; none, if_empty or #CALC!.
def('FILTER', (args, ctx) => {
  const m = rowsOf(args[0], ctx);
  const inc = rowsOf(args[1], ctx);
  let byCol;
  if (inc[0].length === 1 && inc.length === m.length) byCol = false;
  else if (inc.length === 1 && inc[0].length === m[0].length) byCol = true;
  else return E.value;
  const flags = byCol ? inc[0] : inc.map((row) => row[0]);
  const err = flags.find((f) => f instanceof XlError);
  if (err) return err;
  if (flags.some((f) => typeof f === 'string')) return E.value;
  const lines = byCol ? columns(m) : m;
  const out = lines.filter((_, k) => (typeof flags[k] === 'boolean' ? flags[k] : Number(flags[k] ?? 0) !== 0));
  if (!out.length) return given(args[2]) ? args[2] : E.calc;
  return new Arr(byCol ? columns(out) : out);
}, ARRAY);

// TRANSPOSE(array): rows to columns.
def('TRANSPOSE', (args, ctx) => new Arr(columns(rowsOf(args[0], ctx))), ARRAY);

// @x (SINGLE in the file): one value, by implicit intersection with the formula's row or column.
def('SINGLE', (args, ctx) => scalar(args[0], ctx));

// D1# (ANCHORARRAY in the file): the range the formula in D1 spills into; #REF! when it spills nowhere.
def('ANCHORARRAY', (args, ctx) => {
  const a = args[0];
  if (!(a instanceof Area) || a.size !== 1) return E.ref;
  const s = ctx.spillOf?.(a.si, a.r0, a.c0);
  return s ? new Area(a.si, s.r0, s.c0, s.r1, s.c1) : E.ref;
});

// (LET and LAMBDA are worked out by evaluate.js; these entries only say the engine knows them.)
def('LET LAMBDA', () => E.value);
