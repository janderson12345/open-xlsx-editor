// The distributions Excel has beyond the normal ones: Student's t, chi-squared, F, gamma, beta, lognormal, Weibull,
// negative binomial, hypergeometric; their inverses; the tests (T.TEST, CHISQ.TEST, F.TEST, Z.TEST) and CONFIDENCE;
// GAMMA, GAMMALN, PHI, GAUSS, FISHERINV; and the names of Excel 2007 and before for them (TDIST, CHIDIST, NORMDIST...).
// Built on the gamma function (Lanczos), the regularized incomplete gamma (series and continued fraction) and the
// regularized incomplete beta (continued fraction), as numerical libraries work them out. Measured against Excel with
// the formula probe (test/formulas-excel.json). Registered into functions.js's table; evaluate.js imports this file.
import { E, finite, matrix } from './values.js';
import { def, given, num, numbers } from './functions.js';
import { normCdf, normPdf, normInv, cum, sampleSd, mean } from './functions-stats.js';

// ---- The core

const LANCZOS = [0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059, 12.507343278686905,
  -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7];
/** ln Γ(x) for x > 0. */
export function lnGamma(x) {
  if (x < 0.5) return Math.log(Math.PI / Math.abs(Math.sin(Math.PI * x))) - lnGamma(1 - x);
  const z = x - 1;
  let a = LANCZOS[0];
  const t = z + 7.5;
  for (let i = 1; i < 9; i++) a += LANCZOS[i] / (z + i);
  return 0.5 * Math.log(2 * Math.PI) + (z + 0.5) * Math.log(t) - t + Math.log(a);
}
/** Γ(x): exact for whole numbers up to 170, else by Lanczos. */
export function gamma(x) {
  if (Number.isInteger(x)) { if (x <= 0) throw E.num; let p = 1; for (let i = 2; i < x; i++) p *= i; return p; }
  if (x < 0.5) return Math.PI / (Math.sin(Math.PI * x) * gamma(1 - x));
  return Math.exp(lnGamma(x));
}
/** The regularized lower incomplete gamma P(a, x). */
export function gammaP(a, x) {
  if (x <= 0) return 0;
  if (x < a + 1) {
    let sum = 1 / a;
    let term = sum;
    for (let n = 1; n < 1000; n++) { term *= x / (a + n); sum += term; if (Math.abs(term) < Math.abs(sum) * 1e-17) break; }
    return Math.min(1, sum * Math.exp(-x + a * Math.log(x) - lnGamma(a)));
  }
  return 1 - gammaQcf(a, x);
}
/** The regularized upper incomplete gamma Q(a, x) = 1 - P(a, x), kept accurate when it is small. */
export const gammaQ = (a, x) => (x <= 0 ? 1 : x < a + 1 ? 1 - gammaP(a, x) : gammaQcf(a, x));
function gammaQcf(a, x) {
  const tiny = 1e-300;
  let b = x + 1 - a;
  let c = 1 / tiny;
  let d = 1 / b;
  let h = d;
  for (let i = 1; i < 1000; i++) {
    const an = -i * (i - a);
    b += 2;
    d = an * d + b; if (Math.abs(d) < tiny) d = tiny;
    c = b + an / c; if (Math.abs(c) < tiny) c = tiny;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < 1e-16) break;
  }
  return Math.exp(-x + a * Math.log(x) - lnGamma(a)) * h;
}
/** The continued fraction of the incomplete beta (Lentz). */
function betaCf(a, b, x) {
  const tiny = 1e-300;
  let c = 1;
  let d = 1 - (a + b) * x / (a + 1);
  if (Math.abs(d) < tiny) d = tiny;
  d = 1 / d;
  let h = d;
  for (let m = 1; m < 1000; m++) {
    const m2 = 2 * m;
    let aa = m * (b - m) * x / ((a + m2 - 1) * (a + m2));
    d = 1 + aa * d; if (Math.abs(d) < tiny) d = tiny;
    c = 1 + aa / c; if (Math.abs(c) < tiny) c = tiny;
    d = 1 / d; h *= d * c;
    aa = -(a + m) * (a + b + m) * x / ((a + m2) * (a + m2 + 1));
    d = 1 + aa * d; if (Math.abs(d) < tiny) d = tiny;
    c = 1 + aa / c; if (Math.abs(c) < tiny) c = tiny;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < 1e-16) break;
  }
  return h;
}
/** The regularized incomplete beta I_x(a, b). */
export function betaI(x, a, b) {
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  const front = Math.exp(lnGamma(a + b) - lnGamma(a) - lnGamma(b) + a * Math.log(x) + b * Math.log1p(-x));
  return x < (a + 1) / (a + b + 2) ? front * betaCf(a, b, x) / a : 1 - front * betaCf(b, a, 1 - x) / b;
}
/** A distribution's inverse: the x in [lo, hi] where `cdf` reaches p (bisection, then Newton's steps by `pdf`). */
function invert(p, cdf, pdf, lo, hi) {
  if (!(p >= 0 && p <= 1)) throw E.num;
  // (An upper end found by doubling, when the support has none.)
  while (hi === Infinity || cdf(hi) < p) { hi = hi === Infinity ? 1 : hi * 2; if (hi > 1e300) throw E.num; }
  let a = lo;
  let b = hi;
  for (let i = 0; i < 200 && b - a > 1e-15 * Math.max(1, Math.abs(b)); i++) { const m = (a + b) / 2; if (cdf(m) < p) a = m; else b = m; }
  let x = (a + b) / 2;
  for (let i = 0; i < 6; i++) { const f = pdf(x); if (!(f > 0)) break; const nx = x - (cdf(x) - p) / f; if (!(nx > lo) || !(nx < hi)) break; x = nx; }
  return x;
}
const positive = (x) => { if (!(x > 0)) throw E.num; return x; };
const prob = (p) => { if (!(p > 0 && p < 1)) throw E.num; return p; };

// ---- Student's t

const tCdf = (t, df) => { const p = 0.5 * betaI(df / (df + t * t), df / 2, 0.5); return t > 0 ? 1 - p : p; };
const tPdf = (t, df) => Math.exp(lnGamma((df + 1) / 2) - lnGamma(df / 2) - 0.5 * Math.log(df * Math.PI) - (df + 1) / 2 * Math.log1p(t * t / df));
const tInv = (p, df) => { if (p === 0.5) return 0; const right = p > 0.5; const q = right ? 1 - p : p; const x = invert(1 - q, (t) => tCdf(t, df), (t) => tPdf(t, df), 0, Infinity); return right ? x : -x; };
const tTail = (t, df) => 0.5 * betaI(df / (df + t * t), df / 2, 0.5);
const df1 = (v, ctx) => { const d = num(v, ctx); if (d < 1) throw E.num; return d; };
def('T.DIST', (a, ctx) => { const t = num(a[0], ctx); const df = Math.trunc(df1(a[1], ctx)); return cum(a[2], ctx) ? tCdf(t, df) : tPdf(t, df); });
def('T.DIST.2T', (a, ctx) => { const t = num(a[0], ctx); const df = Math.trunc(df1(a[1], ctx)); if (t < 0) throw E.num; return 2 * tTail(t, df); });
def('T.DIST.RT', (a, ctx) => { const t = num(a[0], ctx); const df = Math.trunc(df1(a[1], ctx)); return t >= 0 ? tTail(t, df) : 1 - tTail(t, df); });
def('T.INV', (a, ctx) => tInv(prob(num(a[0], ctx)), Math.trunc(df1(a[1], ctx))));
def('T.INV.2T', (a, ctx) => { const p = num(a[0], ctx); if (!(p > 0 && p <= 1)) throw E.num; return Math.abs(tInv(p / 2, Math.trunc(df1(a[1], ctx)))); });
def('TDIST', (a, ctx) => { const t = num(a[0], ctx); const df = Math.trunc(df1(a[1], ctx)); const tails = Math.trunc(num(a[2], ctx)); if (t < 0 || (tails !== 1 && tails !== 2)) throw E.num; return tails * tTail(t, df); });
def('TINV', (a, ctx) => { const p = num(a[0], ctx); if (!(p > 0 && p <= 1)) throw E.num; return Math.abs(tInv(p / 2, Math.trunc(df1(a[1], ctx)))); });

// ---- Chi-squared

const chiPdf = (x, k) => (x <= 0 ? 0 : Math.exp((k / 2 - 1) * Math.log(x) - x / 2 - (k / 2) * Math.LN2 - lnGamma(k / 2)));
const chiInv = (p, k) => invert(p, (x) => gammaP(k / 2, x / 2), (x) => chiPdf(x, k), 0, Infinity);
const dfChi = (v, ctx) => { const d = Math.trunc(num(v, ctx)); if (d < 1 || d > 1e10) throw E.num; return d; };
def('CHISQ.DIST', (a, ctx) => { const x = num(a[0], ctx); const k = dfChi(a[1], ctx); if (x < 0) throw E.num; return cum(a[2], ctx) ? gammaP(k / 2, x / 2) : chiPdf(x, k); });
def('CHISQ.DIST.RT CHIDIST', (a, ctx) => { const x = num(a[0], ctx); const k = dfChi(a[1], ctx); if (x < 0) throw E.num; return gammaQ(k / 2, x / 2); });
def('CHISQ.INV', (a, ctx) => { const p = num(a[0], ctx); if (!(p >= 0 && p < 1)) throw E.num; return p === 0 ? 0 : chiInv(p, dfChi(a[1], ctx)); });
def('CHISQ.INV.RT CHIINV', (a, ctx) => { const p = num(a[0], ctx); if (!(p > 0 && p <= 1)) throw E.num; return p === 1 ? 0 : chiInv(1 - p, dfChi(a[1], ctx)); });

// ---- F

const fCdf = (x, d1, d2) => (x <= 0 ? 0 : betaI(d1 * x / (d1 * x + d2), d1 / 2, d2 / 2));
const fPdf = (x, d1, d2) => (x <= 0 ? 0 : Math.exp(0.5 * (d1 * Math.log(d1 * x) + d2 * Math.log(d2) - (d1 + d2) * Math.log(d1 * x + d2)) - Math.log(x) - (lnGamma(d1 / 2) + lnGamma(d2 / 2) - lnGamma((d1 + d2) / 2))));
const dfs = (a, ctx, i) => { const d1 = Math.trunc(num(a[i], ctx)); const d2 = Math.trunc(num(a[i + 1], ctx)); if (d1 < 1 || d2 < 1) throw E.num; return [d1, d2]; };
def('F.DIST', (a, ctx) => { const x = num(a[0], ctx); const [d1, d2] = dfs(a, ctx, 1); if (x < 0) throw E.num; return cum(a[3], ctx) ? fCdf(x, d1, d2) : fPdf(x, d1, d2); });
def('F.DIST.RT FDIST', (a, ctx) => { const x = num(a[0], ctx); const [d1, d2] = dfs(a, ctx, 1); if (x < 0) throw E.num; return x === 0 ? 1 : betaI(d2 / (d2 + d1 * x), d2 / 2, d1 / 2); });
def('F.INV', (a, ctx) => { const p = num(a[0], ctx); const [d1, d2] = dfs(a, ctx, 1); if (!(p >= 0 && p < 1)) throw E.num; return p === 0 ? 0 : invert(p, (x) => fCdf(x, d1, d2), (x) => fPdf(x, d1, d2), 0, Infinity); });
def('F.INV.RT FINV', (a, ctx) => { const p = num(a[0], ctx); const [d1, d2] = dfs(a, ctx, 1); if (!(p > 0 && p <= 1)) throw E.num; return p === 1 ? 0 : invert(1 - p, (x) => fCdf(x, d1, d2), (x) => fPdf(x, d1, d2), 0, Infinity); });

// ---- Gamma and beta

def('GAMMA', (a, ctx) => { const x = num(a[0], ctx); if (x <= 0 && Number.isInteger(x)) throw E.num; return finite(gamma(x)); });
def('GAMMALN GAMMALN.PRECISE', (a, ctx) => lnGamma(positive(num(a[0], ctx))));
const gammaPdf = (x, al, be) => (x < 0 ? 0 : x === 0 ? (al === 1 ? 1 / be : al < 1 ? Infinity : 0) : Math.exp((al - 1) * Math.log(x) - x / be - lnGamma(al) - al * Math.log(be)));
def('GAMMA.DIST GAMMADIST', (a, ctx) => {
  const x = num(a[0], ctx); const al = positive(num(a[1], ctx)); const be = positive(num(a[2], ctx));
  if (x < 0) throw E.num;
  return cum(a[3], ctx) ? gammaP(al, x / be) : finite(gammaPdf(x, al, be));
});
def('GAMMA.INV GAMMAINV', (a, ctx) => {
  const p = num(a[0], ctx); const al = positive(num(a[1], ctx)); const be = positive(num(a[2], ctx));
  if (!(p >= 0 && p < 1)) throw E.num;
  return p === 0 ? 0 : invert(p, (x) => gammaP(al, x / be), (x) => gammaPdf(x, al, be), 0, Infinity);
});
const betaPdf = (x, al, be) => (x <= 0 || x >= 1 ? 0 : Math.exp((al - 1) * Math.log(x) + (be - 1) * Math.log1p(-x) + lnGamma(al + be) - lnGamma(al) - lnGamma(be)));
/** BETA.DIST's and BETA.INV's bounds A and B (0 and 1 when not given). */
const bounds = (a, ctx, i) => { const lo = given(a[i]) ? num(a[i], ctx) : 0; const hi = given(a[i + 1]) ? num(a[i + 1], ctx) : 1; if (lo >= hi) throw E.num; return [lo, hi]; };
def('BETA.DIST', (a, ctx) => {
  const x = num(a[0], ctx); const al = positive(num(a[1], ctx)); const be = positive(num(a[2], ctx)); const [lo, hi] = bounds(a, ctx, 4);
  if (x < lo || x > hi) throw E.num;
  const u = (x - lo) / (hi - lo);
  return cum(a[3], ctx) ? betaI(u, al, be) : betaPdf(u, al, be) / (hi - lo);
});
def('BETADIST', (a, ctx) => {
  const x = num(a[0], ctx); const al = positive(num(a[1], ctx)); const be = positive(num(a[2], ctx)); const [lo, hi] = bounds(a, ctx, 3);
  if (x < lo || x > hi) throw E.num;
  return betaI((x - lo) / (hi - lo), al, be);
});
def('BETA.INV BETAINV', (a, ctx) => {
  const p = num(a[0], ctx); const al = positive(num(a[1], ctx)); const be = positive(num(a[2], ctx)); const [lo, hi] = bounds(a, ctx, 3);
  if (!(p > 0 && p <= 1)) throw E.num;
  return lo + (hi - lo) * (p === 1 ? 1 : invert(p, (x) => betaI(x, al, be), (x) => betaPdf(x, al, be), 0, 1));
});

// ---- Lognormal, Weibull, negative binomial, hypergeometric, binomial inverse

def('LOGNORM.DIST', (a, ctx) => {
  const x = positive(num(a[0], ctx)); const m = num(a[1], ctx); const s = positive(num(a[2], ctx));
  const z = (Math.log(x) - m) / s;
  return cum(a[3], ctx) ? normCdf(z) : normPdf(z) / (x * s);
});
def('LOGNORMDIST', (a, ctx) => { const x = positive(num(a[0], ctx)); const s = positive(num(a[2], ctx)); return normCdf((Math.log(x) - num(a[1], ctx)) / s); });
def('LOGNORM.INV LOGINV', (a, ctx) => { const p = prob(num(a[0], ctx)); const s = positive(num(a[2], ctx)); return Math.exp(num(a[1], ctx) + s * normInv(p)); });
def('WEIBULL.DIST WEIBULL', (a, ctx) => {
  const x = num(a[0], ctx); const al = positive(num(a[1], ctx)); const be = positive(num(a[2], ctx));
  if (x < 0) throw E.num;
  return cum(a[3], ctx) ? -Math.expm1(-((x / be) ** al)) : (al / be ** al) * x ** (al - 1) * Math.exp(-((x / be) ** al));
});
const lnChoose = (n, k) => lnGamma(n + 1) - lnGamma(k + 1) - lnGamma(n - k + 1);
const negbinom = (f, s, p) => Math.exp(lnChoose(f + s - 1, s - 1) + s * Math.log(p) + f * Math.log1p(-p));
def('NEGBINOM.DIST', (a, ctx) => {
  const f = Math.trunc(num(a[0], ctx)); const s = Math.trunc(num(a[1], ctx)); const p = num(a[2], ctx);
  if (f < 0 || s < 1 || p <= 0 || p >= 1) throw E.num;
  if (!cum(a[3], ctx)) return negbinom(f, s, p);
  return betaI(p, s, f + 1);
});
def('NEGBINOMDIST', (a, ctx) => {
  const f = Math.trunc(num(a[0], ctx)); const s = Math.trunc(num(a[1], ctx)); const p = num(a[2], ctx);
  if (f < 0 || s < 1 || p <= 0 || p >= 1) throw E.num;
  return negbinom(f, s, p);
});
/** The number of ways, multiplied out (exact while it fits a double), else by logarithms. */
const choose = (n, k) => { k = Math.min(k, n - k); let c = 1; for (let i = 1; i <= k; i++) c = (c * (n - k + i)) / i; return Number.isFinite(c) ? Math.round(c) : Math.exp(lnChoose(n, k)); };
const hyp = (k, n, K, N) => (N <= 1000 ? (choose(K, k) * choose(N - K, n - k)) / choose(N, n) : Math.exp(lnChoose(K, k) + lnChoose(N - K, n - k) - lnChoose(N, n)));
const hypArgs = (a, ctx) => {
  const [k, n, K, N] = [0, 1, 2, 3].map((i) => Math.trunc(num(a[i], ctx)));
  if (k < 0 || k > n || k > K || n > N || K > N || n - k > N - K || N < 1) throw E.num;
  return [k, n, K, N];
};
def('HYPGEOM.DIST', (a, ctx) => {
  const [k, n, K, N] = hypArgs(a, ctx);
  if (!cum(a[4], ctx)) return hyp(k, n, K, N);
  let s = 0;
  for (let i = Math.max(0, n - (N - K)); i <= k; i++) s += hyp(i, n, K, N);
  return Math.min(1, s);
});
def('HYPGEOMDIST', (a, ctx) => { const [k, n, K, N] = hypArgs(a, ctx); return hyp(k, n, K, N); });
def('BINOM.INV CRITBINOM', (a, ctx) => {
  const n = Math.trunc(num(a[0], ctx)); const p = num(a[1], ctx); const al = num(a[2], ctx);
  if (n < 0 || p < 0 || p > 1 || al <= 0 || al >= 1) throw E.num;
  let s = 0;
  for (let k = 0; k <= n; k++) {
    s += Math.exp(lnChoose(n, k) + (k ? k * Math.log(p) : 0) + (n - k ? (n - k) * Math.log1p(-p) : 0));
    if (s >= al * (1 - 1e-15)) return k;
  }
  return n;
});

// ---- Confidence and the tests

def('CONFIDENCE.NORM CONFIDENCE', (a, ctx) => {
  const al = num(a[0], ctx); const s = positive(num(a[1], ctx)); const n = Math.trunc(num(a[2], ctx));
  if (!(al > 0 && al < 1) || n < 1) throw E.num;
  return normInv(1 - al / 2) * s / Math.sqrt(n);
});
def('CONFIDENCE.T', (a, ctx) => {
  const al = num(a[0], ctx); const s = positive(num(a[1], ctx)); const n = Math.trunc(num(a[2], ctx));
  if (!(al > 0 && al < 1) || n < 1) throw E.num;
  if (n === 1) throw E.div0;
  return Math.abs(tInv(al / 2, n - 1)) * s / Math.sqrt(n);
});
def('Z.TEST ZTEST', (a, ctx) => {
  const xs = numbers([a[0]], ctx); const x = num(a[1], ctx);
  if (xs.length < 1) throw E.na;
  const s = given(a[2]) ? num(a[2], ctx) : sampleSd(xs);
  if (!(s > 0)) throw E.div0;
  return 1 - normCdf((mean(xs) - x) / (s / Math.sqrt(xs.length)));
});
const sampleVar = (xs) => { const m = mean(xs); return xs.reduce((s, x) => s + (x - m) ** 2, 0) / (xs.length - 1); };
def('T.TEST TTEST', (a, ctx) => {
  const tails = Math.trunc(num(a[2], ctx)); const type = Math.trunc(num(a[3], ctx));
  if ((tails !== 1 && tails !== 2) || type < 1 || type > 3) throw E.num;
  let t;
  let df;
  if (type === 1) {
    const x = matrix(a[0], ctx).flat();
    const y = matrix(a[1], ctx).flat();
    if (x.length !== y.length) throw E.na;
    const d = x.map((v, k) => (typeof v === 'number' && typeof y[k] === 'number' ? v - y[k] : null)).filter((v) => v != null);
    if (d.length < 2) throw E.div0;
    t = mean(d) / Math.sqrt(sampleVar(d) / d.length);
    df = d.length - 1;
  } else {
    const x = numbers([a[0]], ctx);
    const y = numbers([a[1]], ctx);
    if (x.length < 2 || y.length < 2) throw E.div0;
    const v1 = sampleVar(x);
    const v2 = sampleVar(y);
    if (type === 2) {
      df = x.length + y.length - 2;
      const sp = ((x.length - 1) * v1 + (y.length - 1) * v2) / df;
      t = (mean(x) - mean(y)) / Math.sqrt(sp * (1 / x.length + 1 / y.length));
    } else {
      const s1 = v1 / x.length;
      const s2 = v2 / y.length;
      t = (mean(x) - mean(y)) / Math.sqrt(s1 + s2);
      df = (s1 + s2) ** 2 / (s1 * s1 / (x.length - 1) + s2 * s2 / (y.length - 1));
    }
  }
  if (!Number.isFinite(t)) throw E.div0;
  return tails * tTail(Math.abs(t), df);
});
def('CHISQ.TEST CHITEST', (a, ctx) => {
  const act = matrix(a[0], ctx);
  const exp = matrix(a[1], ctx);
  if (act.length !== exp.length || act[0].length !== exp[0].length) throw E.na;
  let chi = 0;
  for (let r = 0; r < act.length; r++) {
    for (let c = 0; c < act[0].length; c++) {
      const o = act[r][c]; const e = exp[r][c];
      if (typeof o !== 'number' || typeof e !== 'number') continue;
      if (e === 0) throw E.div0;
      chi += (o - e) ** 2 / e;
    }
  }
  const rows = act.length;
  const cols = act[0].length;
  const df = rows > 1 && cols > 1 ? (rows - 1) * (cols - 1) : rows * cols - 1;
  if (df < 1) throw E.na;
  return gammaQ(df / 2, chi / 2);
});
def('F.TEST FTEST', (a, ctx) => {
  const x = numbers([a[0]], ctx); const y = numbers([a[1]], ctx);
  if (x.length < 2 || y.length < 2) throw E.div0;
  const v1 = sampleVar(x); const v2 = sampleVar(y);
  if (v1 === 0 || v2 === 0) throw E.div0;
  const p = fCdf(v1 / v2, x.length - 1, y.length - 1);
  return 2 * Math.min(p, 1 - p);
});

// ---- Small ones, and the old names of the normal family

def('PHI', (a, ctx) => normPdf(num(a[0], ctx)));
def('GAUSS', (a, ctx) => normCdf(num(a[0], ctx)) - 0.5);
def('FISHERINV', (a, ctx) => Math.tanh(num(a[0], ctx)));
def('NORMDIST', (a, ctx) => { const x = num(a[0], ctx); const m = num(a[1], ctx); const s = positive(num(a[2], ctx)); const z = (x - m) / s; return cum(a[3], ctx) ? normCdf(z) : normPdf(z) / s; });
def('NORMINV', (a, ctx) => { const s = positive(num(a[2], ctx)); return num(a[1], ctx) + s * normInv(num(a[0], ctx)); });
def('NORMSDIST', (a, ctx) => normCdf(num(a[0], ctx)));
def('NORMSINV', (a, ctx) => normInv(num(a[0], ctx)));
def('EXPONDIST', (a, ctx) => { const x = num(a[0], ctx); const l = num(a[1], ctx); if (x < 0 || l <= 0) throw E.num; return cum(a[2], ctx) ? -Math.expm1(-l * x) : l * Math.exp(-l * x); });
