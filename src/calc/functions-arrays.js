// Excel 365's array functions that shape arrays (TAKE, DROP, VSTACK, HSTACK, CHOOSECOLS, CHOOSEROWS, TOCOL, TOROW,
// WRAPROWS, WRAPCOLS, EXPAND), split text into one (TEXTSPLIT), turn one into text (VALUETOTEXT, ARRAYTOTEXT), and
// call a LAMBDA over one (MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY, ISOMITTED). Measured against Excel with the
// formula probe (test/formulas-excel.json). Registered into functions.js's table; evaluate.js imports this file.
import { Arr, Ref3D, MISSING, XlError, E, scalar, matrix, toText, toBool, Lambda, Lambdas } from './values.js';
import { def, given, num, text } from './functions.js';
import { maxOf } from '../ooxml/listmath.js';

const ARRAY = { array: true };
const rowsOf = (v, ctx) => matrix(v, ctx);
/** The arguments with a 3D reference as its sheets' areas, one after another (=VSTACK(Jan:Apr!A1:B2) stacks them). */
const sheetwise = (args) => args.flatMap((x) => (x instanceof Ref3D ? x.areas : [x]));
const columns = (m) => m[0].map((_, c) => m.map((row) => row[c]));
const whole = (v, ctx) => Math.trunc(num(v, ctx));
const flag = (v, ctx, dflt = false) => (given(v) ? toBool(scalar(v, ctx)) : dflt);
const out = (rows) => (rows.length && rows[0].length ? new Arr(rows) : E.calc);

/** The first (n > 0) or last (n < 0) n of a list. */
const ends = (list, n) => (n >= 0 ? list.slice(0, n) : list.slice(Math.max(0, list.length + n)));

def('TAKE', (a, ctx) => {
  let m = rowsOf(a[0], ctx);
  if (given(a[1])) { const n = whole(a[1], ctx); if (n === 0) return E.calc; m = ends(m, n); }
  if (given(a[2])) { const n = whole(a[2], ctx); if (n === 0) return E.calc; m = m.map((row) => ends(row, n)); }
  return out(m);
}, ARRAY);
def('DROP', (a, ctx) => {
  let m = rowsOf(a[0], ctx);
  if (given(a[1])) { const n = whole(a[1], ctx); m = n >= 0 ? m.slice(n) : m.slice(0, Math.max(0, m.length + n)); }
  if (given(a[2])) { const n = whole(a[2], ctx); m = m.map((row) => (n >= 0 ? row.slice(n) : row.slice(0, Math.max(0, row.length + n)))); }
  return out(m);
}, ARRAY);

/** Arrays stacked down (or across), the narrower (or shorter) padded with #N/A. */
function stack(args, ctx, across) {
  const ms = sheetwise(args).filter((x) => x !== MISSING && x !== undefined).map((x) => rowsOf(x, ctx));
  if (!across) {
    const width = maxOf(ms.map((m) => m[0].length));
    return new Arr(ms.flatMap((m) => m.map((row) => [...row, ...Array(width - row.length).fill(E.na)])));
  }
  const height = maxOf(ms.map((m) => m.length));
  return new Arr(Array.from({ length: height }, (_, r) => ms.flatMap((m) => (r < m.length ? m[r] : Array(m[0].length).fill(E.na)))));
}
def('VSTACK', (a, ctx) => stack(a, ctx, false), ARRAY);
// (HSTACK of LAMBDAs, HSTACK(SUM,AVERAGE): the functions side by side, for GROUPBY and PIVOTBY: functions-groupby.js.)
def('HSTACK', (a, ctx) => (a.length && a.every((x) => x instanceof Lambda) ? new Lambdas(a) : stack(a, ctx, true)), ARRAY);

/** The rows (or columns) of `lines` by 1-based index, a negative one from the end; out of range: #VALUE!. */
function choose(lines, picks, ctx) {
  const out = [];
  for (const p of picks) {
    for (const v of rowsOf(p, ctx).flat()) {
      const n = Math.trunc(num(v, ctx));
      const k = n > 0 ? n - 1 : lines.length + n;
      if (n === 0 || k < 0 || k >= lines.length) return E.value;
      out.push(lines[k]);
    }
  }
  return out;
}
def('CHOOSEROWS', (a, ctx) => { const got = choose(rowsOf(a[0], ctx), a.slice(1), ctx); return got instanceof XlError ? got : new Arr(got); }, ARRAY);
def('CHOOSECOLS', (a, ctx) => { const got = choose(columns(rowsOf(a[0], ctx)), a.slice(1), ctx); return got instanceof XlError ? got : new Arr(columns(got)); }, ARRAY);

/** TOCOL and TOROW's values: by rows (or columns), blanks (1), errors (2) or both (3) left out. */
function flatten(a, ctx) {
  // (A 3D reference: its sheets' areas one below another, then read across, as Excel does.)
  const m = a[0] instanceof Ref3D ? a[0].areas.flatMap((x) => rowsOf(x, ctx)) : rowsOf(a[0], ctx);
  const ignore = given(a[1]) ? whole(a[1], ctx) : 0;
  const list = (flag(a[2], ctx) ? columns(m) : m).flat();
  return list.filter((v) => !((ignore & 1) && (v == null || v === MISSING)) && !((ignore & 2) && v instanceof XlError));
}
def('TOCOL', (a, ctx) => { const l = flatten(a, ctx); return l.length ? new Arr(l.map((v) => [v])) : E.calc; }, ARRAY);
def('TOROW', (a, ctx) => { const l = flatten(a, ctx); return l.length ? new Arr([l]) : E.calc; }, ARRAY);

/** A row or column wrapped into rows (or columns) of `count`, the last padded (#N/A, or `pad`). */
function wrap(a, ctx, byCols) {
  const m = rowsOf(a[0], ctx);
  if (m.length > 1 && m[0].length > 1) return E.value;
  const list = m.flat();
  const count = whole(a[1], ctx);
  if (count < 1) return E.num;
  const pad = given(a[2]) ? a[2] : E.na;
  const lines = [];
  for (let k = 0; k < list.length; k += count) lines.push([...list.slice(k, k + count), ...Array(Math.max(0, k + count - list.length)).fill(pad)]);
  return new Arr(byCols ? columns(lines) : lines);
}
def('WRAPROWS', (a, ctx) => wrap(a, ctx, false), ARRAY);
def('WRAPCOLS', (a, ctx) => wrap(a, ctx, true), ARRAY);

def('EXPAND', (a, ctx) => {
  const m = rowsOf(a[0], ctx);
  const rows = given(a[1]) ? whole(a[1], ctx) : m.length;
  const cols = given(a[2]) ? whole(a[2], ctx) : m[0].length;
  if (rows < m.length || cols < m[0].length) return E.value;
  const pad = given(a[3]) ? a[3] : E.na;
  return new Arr(Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (__, c) => (r < m.length && c < m[0].length ? m[r][c] : pad))));
}, ARRAY);

/** Text split at any of the delimiters (`mode` 1: capitals aside). */
function split(s, delims, mode) {
  if (!delims.length) return [s];
  const pieces = [];
  let at = 0;
  const hay = mode ? s.toLowerCase() : s;
  for (let i = 0; i <= s.length;) {
    const d = delims.find((x) => x && hay.startsWith(mode ? x.toLowerCase() : x, i));
    if (i === s.length) { pieces.push(s.slice(at)); break; }
    if (d) { pieces.push(s.slice(at, i)); i += d.length; at = i; } else i++;
  }
  return pieces;
}
def('TEXTSPLIT', (a, ctx) => {
  const s = text(a[0], ctx);
  const list = (v) => (given(v) ? rowsOf(v, ctx).flat().map((x) => toText(x)) : []);
  const cols = list(a[1]);
  const rows = list(a[2]);
  const ignore = flag(a[3], ctx);
  const mode = given(a[4]) ? whole(a[4], ctx) : 0;
  const pad = given(a[5]) ? a[5] : E.na;
  const keep = (xs) => (ignore ? xs.filter((x) => x !== '') : xs);
  const lines = keep(split(s, rows, mode)).map((line) => keep(split(line, cols, mode)));
  const width = maxOf(lines.map((l) => l.length));
  return out(lines.map((l) => [...l, ...Array(width - l.length).fill(pad)]));
}, ARRAY);

/** A value as VALUETOTEXT and ARRAYTOTEXT write it: strict (1) quotes text. */
const asText = (v, strict) => (typeof v === 'string' ? (strict ? `"${v.replace(/"/g, '""')}"` : v) : v == null || v === MISSING ? '' : toText(v));
def('VALUETOTEXT', (a, ctx) => asText(scalar(a[0], ctx), given(a[1]) && whole(a[1], ctx) === 1));
def('ARRAYTOTEXT', (a, ctx) => {
  const m = rowsOf(a[0], ctx);
  const strict = given(a[1]) && whole(a[1], ctx) === 1;
  return strict ? `{${m.map((row) => row.map((v) => asText(v, true)).join(',')).join(';')}}` : m.flat().map((v) => asText(v, false)).join(', ');
}, ARRAY);

// ---- A LAMBDA called over arrays

/** The LAMBDA a helper is given (its last argument), or #VALUE!. */
const lambdaOf = (v) => (v instanceof Lambda ? v : null);
/** One value from a LAMBDA's answer (an array of one, or one value; a bigger array is #CALC!). */
function single(v, ctx) {
  if (v instanceof Arr) return v.height === 1 && v.width === 1 ? v.rows[0][0] : E.calc;
  return scalar(v, ctx);
}

def('MAP', (a, ctx) => {
  const fn = lambdaOf(a[a.length - 1]);
  if (!fn) return E.value;
  const ms = a.slice(0, -1).map((x) => rowsOf(x, ctx));
  return new Arr(ms[0].map((row, r) => row.map((_, c) => single(ctx.callLambda(fn, ms.map((m) => m[r]?.[c] ?? E.na)), ctx))));
}, ARRAY);
def('REDUCE', (a, ctx) => {
  const fn = lambdaOf(a[2]);
  if (!fn) return E.value;
  let acc = given(a[0]) ? a[0] : 0;
  for (const v of rowsOf(a[1], ctx).flat()) acc = ctx.callLambda(fn, [acc, v]);
  return acc;
}, ARRAY);
def('SCAN', (a, ctx) => {
  const fn = lambdaOf(a[2]);
  if (!fn) return E.value;
  let acc = given(a[0]) ? a[0] : 0;
  return new Arr(rowsOf(a[1], ctx).map((row) => row.map((v) => { acc = single(ctx.callLambda(fn, [acc, v]), ctx); return acc; })));
}, ARRAY);
def('BYROW', (a, ctx) => {
  const fn = lambdaOf(a[1]);
  if (!fn) return E.value;
  return new Arr(rowsOf(a[0], ctx).map((row) => [single(ctx.callLambda(fn, [new Arr([row])]), ctx)]));
}, ARRAY);
def('BYCOL', (a, ctx) => {
  const fn = lambdaOf(a[1]);
  if (!fn) return E.value;
  return new Arr([columns(rowsOf(a[0], ctx)).map((col) => single(ctx.callLambda(fn, [new Arr(col.map((v) => [v]))]), ctx))]);
}, ARRAY);
def('MAKEARRAY', (a, ctx) => {
  const rows = whole(a[0], ctx);
  const cols = whole(a[1], ctx);
  const fn = lambdaOf(a[2]);
  if (!fn) return E.value;
  if (rows < 1 || cols < 1) return E.value;
  return new Arr(Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (__, c) => single(ctx.callLambda(fn, [r + 1, c + 1]), ctx))));
}, ARRAY);
def('ISOMITTED', (a) => a[0] === MISSING);
