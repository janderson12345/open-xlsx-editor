// AGGREGATE, FREQUENCY, TRIMMEAN, PERCENTRANK, MODE.MULT, the A versions of the deviations (STDEVA, STDEVPA, VARA,
// VARPA), STEYX and PROB; regression (LINEST, LOGEST, TREND, GROWTH); matrices (MMULT, MINVERSE, MDETERM, MUNIT); the
// sums of squares of differences (SUMX2MY2, SUMX2PY2, SUMXMY2); and POISSON and BINOMDIST, Excel 2007's names. As
// Excel defines them, measured against Excel with the formula probe (test/formulas-excel.json). Registered into
// functions.js's table; evaluate.js imports this file.
import { Area, Arr, Union, Ref3D, XlError, E, finite, matrix, scalar, toBool } from './values.js';
import { FUNCTIONS, def, given, num, numbers } from './functions.js';

const ARRAY = { array: true };
const isNum = (v) => typeof v === 'number';

// ---- AGGREGATE

const AGG = ['', 'AVERAGE', 'COUNT', 'COUNTA', 'MAX', 'MIN', 'PRODUCT', 'STDEV.S', 'STDEV.P', 'SUM', 'VAR.S', 'VAR.P', 'MEDIAN', 'MODE.SNGL', 'LARGE', 'SMALL',
  'PERCENTILE.INC', 'QUARTILE.INC', 'PERCENTILE.EXC', 'QUARTILE.EXC'];
/** AGGREGATE(function, options, ref1, ...) and AGGREGATE(function, options, array, k): options 0 to 7 leave out nested
 *  SUBTOTALs and AGGREGATEs (0 to 3), hidden rows (1, 3, 5, 7) and errors (2, 3, 6, 7). */
def('AGGREGATE', (a, ctx) => {
  const n = Math.trunc(num(a[0], ctx));
  const o = given(a[1]) ? Math.trunc(num(a[1], ctx)) : 0;
  const name = AGG[n];
  if (!name || o < 0 || o > 7) throw E.value;
  const hidden = o === 1 || o === 3 || o === 5 || o === 7;
  const errors = o === 2 || o === 3 || o === 6 || o === 7;
  const nested = o <= 3;
  const withK = n >= 14;
  const sources = withK ? [a[2]] : a.slice(2);
  const kept = [];
  for (const ref of sources) {
    if (ref instanceof Ref3D) throw E.value;
    const areas = ref instanceof Area ? [ref] : ref instanceof Union ? ref.areas : null;
    // (Functions 1 to 13 take references only: an array is #VALUE!.)
    if (!areas && !withK) throw E.value;
    if (!areas) { for (const v of matrix(ref, ctx).flat()) if (!(errors && v instanceof XlError)) kept.push(v); continue; }
    for (const x of areas) {
      for (let r = x.r0; r <= x.r1; r++) {
        if (hidden && ctx.rowHidden?.(x.si, r)) continue;
        for (let c = x.c0; c <= x.c1; c++) {
          const f = ctx.formulaOf?.(x.si, r, c);
          if (nested && f && /^\s*(_xlfn\.)?(SUBTOTAL|AGGREGATE)\s*\(/i.test(f)) continue;
          const cell = new Area(x.si, r, c, r, c);
          if (errors && scalar(cell, ctx) instanceof XlError) continue;
          kept.push(withK ? scalar(cell, ctx) : cell);
        }
      }
    }
  }
  if (withK) return FUNCTIONS[name].fn([new Arr([kept.length ? kept : [null]]), a[3]], ctx);
  return FUNCTIONS[name].fn(kept.map((v) => (v instanceof Area ? v : new Arr([[v]]))), ctx);
});

// ---- FREQUENCY, TRIMMEAN, PERCENTRANK, MODE.MULT

def('FREQUENCY', (a, ctx) => {
  const data = numbers([a[0]], ctx);
  const bins = numbers([a[1]], ctx);
  const order = bins.map((b, i) => [b, i]).sort((x, y) => x[0] - y[0]);
  const out = new Array(bins.length + 1).fill(0);
  for (const d of data) {
    const at = order.find(([b]) => d <= b);
    out[at ? at[1] : bins.length]++;
  }
  return new Arr(out.map((v) => [v]));
}, ARRAY);
def('TRIMMEAN', (a, ctx) => {
  const xs = numbers([a[0]], ctx).sort((x, y) => x - y);
  const p = num(a[1], ctx);
  if (p < 0 || p >= 1 || !xs.length) throw E.num;
  // (Trimmed: the count times the percent, rounded down to an even number, half from each end.)
  const cut = Math.floor((xs.length * p) / 2);
  const kept = xs.slice(cut, xs.length - cut);
  return kept.reduce((s, x) => s + x, 0) / kept.length;
});
/** PERCENTRANK: x's place among the numbers (between them by straight lines), cut down to `sig` digits. */
function percentRank(a, ctx, exc) {
  const xs = numbers([a[0]], ctx).sort((p, q) => p - q);
  const x = num(a[1], ctx);
  const sig = given(a[2]) ? Math.trunc(num(a[2], ctx)) : 3;
  const n = xs.length;
  if (!n || sig < 1 || x < xs[0] || x > xs[n - 1]) throw n ? E.na : E.num;
  const below = (v) => xs.filter((y) => y < v).length;
  let rank;
  const i = xs.indexOf(x);
  if (i >= 0) rank = exc ? (below(x) + 1) / (n + 1) : n === 1 ? 1 : below(x) / (n - 1);
  else {
    const k = below(x) - 1;
    const lo = xs[k];
    const hi = xs[k + 1];
    const r0 = exc ? (k + 1) / (n + 1) : k / (n - 1);
    const r1 = exc ? (k + 2) / (n + 1) : (k + 1) / (n - 1);
    rank = r0 + ((x - lo) / (hi - lo)) * (r1 - r0);
  }
  const f = 10 ** sig;
  return Math.floor(rank * f + 1e-9) / f;
}
def('PERCENTRANK PERCENTRANK.INC', (a, ctx) => percentRank(a, ctx, false));
def('PERCENTRANK.EXC', (a, ctx) => percentRank(a, ctx, true));
def('MODE.MULT', (a, ctx) => {
  const xs = numbers(a, ctx);
  const counts = new Map();
  for (const x of xs) counts.set(x, (counts.get(x) || 0) + 1);
  const most = Math.max(0, ...counts.values());
  if (most < 2) throw E.na;
  return new Arr([...counts].filter(([, k]) => k === most).map(([x]) => [x]));
}, ARRAY);

// ---- STDEVA, STDEVPA, VARA, VARPA (text 0, TRUE 1, in ranges too), STEYX, PROB

/** The values the A functions count: numbers, text as 0, TRUE 1 and FALSE 0 (in ranges and typed); blanks left out. */
function aValues(args, ctx) {
  const out = [];
  for (const v of args) {
    if (v instanceof Area || v instanceof Union || v instanceof Arr) {
      // (In an array, as typed or worked out, only the numbers: text and TRUE/FALSE left out.)
      const inArray = v instanceof Arr;
      for (const x of (v instanceof Union ? v.areas : [v]).flatMap((y) => matrix(y, ctx).flat())) {
        if (x instanceof XlError) throw x;
        if (isNum(x)) out.push(x); else if (inArray) continue; else if (typeof x === 'boolean') out.push(x ? 1 : 0); else if (typeof x === 'string') out.push(0);
      }
    } else out.push(num(v, ctx));
  }
  return out;
}
const variance = (xs, sample) => { const n = xs.length; if (n < (sample ? 2 : 1)) throw E.div0; const m = xs.reduce((s, x) => s + x, 0) / n; return xs.reduce((s, x) => s + (x - m) ** 2, 0) / (n - (sample ? 1 : 0)); };
def('VARA', (a, ctx) => variance(aValues(a, ctx), true));
def('VARPA', (a, ctx) => variance(aValues(a, ctx), false));
def('STDEVA', (a, ctx) => Math.sqrt(variance(aValues(a, ctx), true)));
def('STDEVPA', (a, ctx) => Math.sqrt(variance(aValues(a, ctx), false)));
/** Pairs of numbers from two equal-sized lists. */
function pairs(ys, xs, ctx) {
  const y = matrix(ys, ctx).flat();
  const x = matrix(xs, ctx).flat();
  if (x.length !== y.length) throw E.na;
  return y.map((v, k) => [x[k], v]).filter(([p, q]) => isNum(p) && isNum(q));
}
def('STEYX', (a, ctx) => {
  const p = pairs(a[0], a[1], ctx);
  const n = p.length;
  if (n < 3) throw E.div0;
  const mx = p.reduce((s, [x]) => s + x, 0) / n;
  const my = p.reduce((s, [, y]) => s + y, 0) / n;
  let sxx = 0; let syy = 0; let sxy = 0;
  for (const [x, y] of p) { sxx += (x - mx) ** 2; syy += (y - my) ** 2; sxy += (x - mx) * (y - my); }
  if (sxx === 0) throw E.div0;
  return Math.sqrt((syy - sxy * sxy / sxx) / (n - 2));
});
def('PROB', (a, ctx) => {
  const x = matrix(a[0], ctx).flat();
  const p = matrix(a[1], ctx).flat();
  if (x.length !== p.length) throw E.na;
  if (p.some((v) => isNum(v) && (v < 0 || v > 1))) throw E.num;
  const total = p.reduce((s, v) => s + (isNum(v) ? v : 0), 0);
  if (Math.abs(total - 1) > 1e-7) throw E.num;
  const lo = num(a[2], ctx);
  const hi = given(a[3]) ? num(a[3], ctx) : lo;
  return x.reduce((s, v, k) => s + (isNum(v) && isNum(p[k]) && v >= lo && v <= hi ? p[k] : 0), 0);
});

// ---- Matrices

const numMatrix = (v, ctx) => { const m = matrix(v, ctx); if (m.some((row) => row.some((x) => !isNum(x)))) throw E.value; return m; };
def('MMULT', (a, ctx) => {
  const x = numMatrix(a[0], ctx);
  const y = numMatrix(a[1], ctx);
  if (x[0].length !== y.length) throw E.value;
  return new Arr(x.map((row) => y[0].map((_, c) => finite(row.reduce((s, v, k) => s + v * y[k][c], 0)))));
}, ARRAY);
/** LU decomposition with partial pivoting: { lu, perm, sign }, or null when singular. */
function lu(m) {
  const n = m.length;
  const a = m.map((row) => [...row]);
  const perm = [...Array(n).keys()];
  let sign = 1;
  for (let k = 0; k < n; k++) {
    let p = k;
    for (let i = k + 1; i < n; i++) if (Math.abs(a[i][k]) > Math.abs(a[p][k])) p = i;
    if (a[p][k] === 0) return null;
    if (p !== k) { [a[p], a[k]] = [a[k], a[p]]; [perm[p], perm[k]] = [perm[k], perm[p]]; sign = -sign; }
    for (let i = k + 1; i < n; i++) {
      a[i][k] /= a[k][k];
      for (let j = k + 1; j < n; j++) a[i][j] -= a[i][k] * a[k][j];
    }
  }
  return { a, perm, sign };
}
const square = (v, ctx) => { const m = numMatrix(v, ctx); if (m.length !== m[0].length) throw E.value; return m; };
def('MDETERM', (a, ctx) => {
  const m = square(a[0], ctx);
  const d = lu(m);
  if (!d) return 0;
  return d.a.reduce((p, row, i) => p * row[i], d.sign);
});
def('MINVERSE', (a, ctx) => {
  const m = square(a[0], ctx);
  const n = m.length;
  const d = lu(m);
  if (!d) throw E.num;
  const out = Array.from({ length: n }, () => new Array(n).fill(0));
  for (let col = 0; col < n; col++) {
    const b = d.perm.map((p) => (p === col ? 1 : 0));
    for (let i = 0; i < n; i++) for (let k = 0; k < i; k++) b[i] -= d.a[i][k] * b[k];
    for (let i = n - 1; i >= 0; i--) { for (let k = i + 1; k < n; k++) b[i] -= d.a[i][k] * b[k]; b[i] /= d.a[i][i]; }
    for (let i = 0; i < n; i++) out[i][col] = b[i];
  }
  return new Arr(out);
}, ARRAY);
def('MUNIT', (a, ctx) => {
  const n = Math.trunc(num(a[0], ctx));
  if (n < 1) throw E.value;
  return new Arr(Array.from({ length: n }, (_, r) => Array.from({ length: n }, (__, c) => (r === c ? 1 : 0))));
}, ARRAY);
/** SUMX2MY2, SUMX2PY2 and SUMXMY2: over the pairs where both are numbers. */
const sumPairs = (f) => (a, ctx) => {
  const x = matrix(a[0], ctx).flat();
  const y = matrix(a[1], ctx).flat();
  if (x.length !== y.length) throw E.na;
  return x.reduce((s, v, k) => s + (isNum(v) && isNum(y[k]) ? f(v, y[k]) : 0), 0);
};
def('SUMX2MY2', sumPairs((x, y) => x * x - y * y));
def('SUMX2PY2', sumPairs((x, y) => x * x + y * y));
def('SUMXMY2', sumPairs((x, y) => (x - y) ** 2));

// ---- Regression

/** Least squares of y on the columns of X (with a constant unless `noConst`): coefficients [b_k ... b_1, a] as LINEST
 *  lists them, and the statistics. Solved by the normal equations (LU): small systems. */
function leastSquares(ys, xs, noConst) {
  const n = ys.length;
  const k = xs[0].length;
  if (n < k + (noConst ? 0 : 1)) throw E.num;
  // With a constant, the xs and ys taken from their means first, as Excel does: an exact line comes out exact (its
  // residuals 0). Then the normal equations, solved by LU: small systems, as these are.
  const mean = (v) => v.reduce((s, x) => s + x, 0) / n;
  const mx = noConst ? new Array(k).fill(0) : xs[0].map((_, i) => mean(xs.map((row) => row[i])));
  const my = noConst ? 0 : mean(ys);
  const X = xs.map((row) => row.map((v, i) => v - mx[i]));
  const Y = ys.map((y) => y - my);
  const p = k;
  const A = Array.from({ length: p }, (_, i) => Array.from({ length: p }, (__, j) => X.reduce((s, row) => s + row[i] * row[j], 0)));
  const b = Array.from({ length: p }, (_, i) => X.reduce((s, row, q) => s + row[i] * Y[q], 0));
  const d = lu(A);
  if (!d) throw E.num;
  const solve = (rhs) => {
    const v = d.perm.map((q) => rhs[q]);
    for (let i = 0; i < p; i++) for (let k = 0; k < i; k++) v[i] -= d.a[i][k] * v[k];
    for (let i = p - 1; i >= 0; i--) { for (let k = i + 1; k < p; k++) v[i] -= d.a[i][k] * v[k]; v[i] /= d.a[i][i]; }
    return v;
  };
  const coef = solve(b);
  const ssres = Y.reduce((s, y, q) => s + (y - X[q].reduce((t, v, i) => t + v * coef[i], 0)) ** 2, 0);
  const sstot = Y.reduce((s, y) => s + y * y, 0);
  const dfres = n - k - (noConst ? 0 : 1);
  const sey = dfres > 0 ? Math.sqrt(ssres / dfres) : 0;
  const inv = Array.from({ length: p }, (_, j) => solve(Array.from({ length: p }, (__, i) => (i === j ? 1 : 0))));
  const se = coef.map((_, i) => Math.sqrt(Math.max(0, inv[i][i])) * sey);
  if (!noConst) {
    // (The constant, and its standard error: from the means.)
    coef.push(my - coef.reduce((s, c, i) => s + c * mx[i], 0));
    const q = mx.reduce((s, xi, i) => s + xi * mx.reduce((t, xj, j) => t + inv[i][j] * xj, 0), 0);
    se.push(Math.sqrt(Math.max(0, 1 / n + q)) * sey);
  }
  const ssreg = sstot - ssres;
  const r2 = sstot === 0 ? 1 : ssreg / sstot;
  // (F: #NUM! when the line is exact, no residuals to set it against.)
  const f = dfres > 0 && ssres > 0 ? (ssreg / k) / (ssres / dfres) : E.num;
  return { coef, se, r2, sey, f, dfres, ssreg, ssres, k };
}
/** The known ys and xs of LINEST and its kin: xs as columns (one column of 1, 2, 3... when not given). */
function known(a, ctx, logY) {
  const ym = numMatrix(a[0], ctx);
  const ys = ym.flat().map((y) => { if (logY) { if (y <= 0) throw E.num; return Math.log(y); } return y; });
  let xs;
  if (given(a[1])) {
    const xm = numMatrix(a[1], ctx);
    const byRow = ym.length === 1 && ym[0].length > 1;
    if (byRow) xs = xm[0].map((_, c) => xm.map((row) => row[c]));
    else if (xm.length === ys.length) xs = xm;
    else if (xm[0].length === ys.length && xm.length >= 1) xs = xm[0].map((_, c) => xm.map((row) => row[c]));
    else throw E.ref;
  } else xs = ys.map((_, i) => [i + 1]);
  if (xs.length !== ys.length) throw E.ref;
  return { ys, xs, byRow: ym.length === 1 && ym[0].length > 1 };
}
function linest(a, ctx, logY) {
  const { ys, xs } = known(a, ctx, logY);
  const noConst = given(a[2]) && !toBool(scalar(a[2], ctx));
  const stats = given(a[3]) && toBool(scalar(a[3], ctx));
  const r = leastSquares(ys, xs, noConst);
  const k = xs[0].length;
  // (Coefficients last x first, then the constant; for LOGEST, the exponentials.)
  const order = [...Array(k).keys()].reverse();
  const coef = order.map((i) => r.coef[i]);
  const cons = noConst ? (logY ? 1 : 0) : r.coef[k];
  const row1 = [...coef.map((c) => (logY ? Math.exp(c) : c)), logY && !noConst ? Math.exp(cons) : cons];
  if (!stats) return new Arr([row1]);
  const na = E.na;
  const width = k + 1;
  const pad = (vals) => [...vals, ...new Array(width - vals.length).fill(na)];
  const row2 = [...order.map((i) => r.se[i]), noConst ? na : r.se[k]];
  return new Arr([row1, row2, pad([r.r2, r.sey]), pad([r.f, r.dfres]), pad([r.ssreg, r.ssres])]);
}
def('LINEST', (a, ctx) => linest(a, ctx, false), ARRAY);
def('LOGEST', (a, ctx) => linest(a, ctx, true), ARRAY);
function forecastAt(a, ctx, logY) {
  const { ys, xs, byRow } = known(a, ctx, logY);
  const noConst = given(a[3]) && !toBool(scalar(a[3], ctx));
  const r = leastSquares(ys, xs, noConst);
  const k = xs[0].length;
  let newX;
  if (given(a[2])) {
    const m = numMatrix(a[2], ctx);
    newX = k === 1 ? m.map((row) => row.map((x) => [x])) : byRow ? [m[0].map((_, c) => m.map((row) => row[c]))] : m.map((row) => [row]);
  } else newX = byRow ? [xs.map((row) => row)] : xs.map((row) => [row]);
  const at = (x) => { const v = x.reduce((s, xi, i) => s + xi * r.coef[i], 0) + (noConst ? 0 : r.coef[k]); return logY ? Math.exp(v) : v; };
  return new Arr(newX.map((row) => row.map(at)));
}
def('TREND', (a, ctx) => forecastAt(a, ctx, false), ARRAY);
def('GROWTH', (a, ctx) => forecastAt(a, ctx, true), ARRAY);

// ---- Excel 2007's names

FUNCTIONS.POISSON = FUNCTIONS['POISSON.DIST'];
FUNCTIONS.BINOMDIST = FUNCTIONS['BINOM.DIST'];
