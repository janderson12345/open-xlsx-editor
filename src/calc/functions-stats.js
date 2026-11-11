// Statistics (CORREL, PEARSON, SLOPE, INTERCEPT, RSQ, FORECAST, COVARIANCE.P and .S, GEOMEAN, HARMEAN, AVEDEV, DEVSQ,
// SKEW, KURT, STANDARDIZE, PERMUT, FISHER; the normal, exponential, Poisson and binomial distributions), hyperbolic
// and reciprocal trigonometry, Roman and other numerals, and a cell's formula and sheet (ISFORMULA, FORMULATEXT,
// SHEET, SHEETS), as Excel defines them. Measured against Excel with the formula probe (test/formulas-excel.json).
// Registered into functions.js's table; evaluate.js imports this file.
import { Area, E, finite, scalar, toText, toBool, matrix } from './values.js';
import { def, given, num, text, numbers } from './functions.js';
import { displayFormula } from './tokens.js';

// ---- Pairs of numbers

/** The pairs of two ranges or arrays of one size where both are numbers (as CORREL and SLOPE take them). */
function pairs(a, b, ctx) {
  const list = (v) => matrix(v, ctx).flat();
  const xs = list(a);
  const ys = list(b);
  if (xs.length !== ys.length) throw E.na;
  const out = [];
  for (let k = 0; k < xs.length; k++) if (typeof xs[k] === 'number' && typeof ys[k] === 'number') out.push([xs[k], ys[k]]);
  return out;
}
function moments(p) {
  const n = p.length;
  const mx = p.reduce((s, [x]) => s + x, 0) / n;
  const my = p.reduce((s, [, y]) => s + y, 0) / n;
  let sxy = 0; let sxx = 0; let syy = 0;
  for (const [x, y] of p) { sxy += (x - mx) * (y - my); sxx += (x - mx) ** 2; syy += (y - my) ** 2; }
  return { n, mx, my, sxy, sxx, syy };
}
const correl = (a, ctx) => {
  const m = moments(pairs(a[0], a[1], ctx));
  if (m.n < 2 || m.sxx === 0 || m.syy === 0) throw E.div0;
  return m.sxy / Math.sqrt(m.sxx * m.syy);
};
def('CORREL PEARSON', correl);
def('RSQ', (a, ctx) => correl(a, ctx) ** 2);
// (SLOPE(known_y, known_x): x the second.)
const fit = (a, ctx) => { const m = moments(pairs(a[1], a[0], ctx)); if (m.n < 1 || m.sxx === 0) throw E.div0; return { slope: m.sxy / m.sxx, m }; };
def('SLOPE', (a, ctx) => fit(a, ctx).slope);
def('INTERCEPT', (a, ctx) => { const { slope, m } = fit(a, ctx); return m.my - slope * m.mx; });
def('FORECAST FORECAST.LINEAR', (a, ctx) => { const x = num(a[0], ctx); const { slope, m } = fit(a.slice(1), ctx); return m.my + slope * (x - m.mx); });
def('COVARIANCE.P', (a, ctx) => { const m = moments(pairs(a[0], a[1], ctx)); if (!m.n) throw E.div0; return m.sxy / m.n; });
def('COVARIANCE.S', (a, ctx) => { const m = moments(pairs(a[0], a[1], ctx)); if (m.n < 2) throw E.div0; return m.sxy / (m.n - 1); });

// ---- One list of numbers

const mean = (xs) => xs.reduce((s, x) => s + x, 0) / xs.length;
def('GEOMEAN', (a, ctx) => { const xs = numbers(a, ctx); if (!xs.length || xs.some((x) => x <= 0)) throw E.num; return Math.exp(mean(xs.map(Math.log))); });
def('HARMEAN', (a, ctx) => { const xs = numbers(a, ctx); if (!xs.length || xs.some((x) => x <= 0)) throw E.num; return xs.length / xs.reduce((s, x) => s + 1 / x, 0); });
def('AVEDEV', (a, ctx) => { const xs = numbers(a, ctx); if (!xs.length) throw E.num; const m = mean(xs); return mean(xs.map((x) => Math.abs(x - m))); });
def('DEVSQ', (a, ctx) => { const xs = numbers(a, ctx); if (!xs.length) throw E.num; const m = mean(xs); return xs.reduce((s, x) => s + (x - m) ** 2, 0); });
const sampleSd = (xs) => { const m = mean(xs); return Math.sqrt(xs.reduce((s, x) => s + (x - m) ** 2, 0) / (xs.length - 1)); };
def('SKEW', (a, ctx) => {
  const xs = numbers(a, ctx);
  const n = xs.length;
  const sd = sampleSd(xs);
  if (n < 3 || sd === 0) throw E.div0;
  const m = mean(xs);
  return n / ((n - 1) * (n - 2)) * xs.reduce((s, x) => s + ((x - m) / sd) ** 3, 0);
});
def('KURT', (a, ctx) => {
  const xs = numbers(a, ctx);
  const n = xs.length;
  const sd = sampleSd(xs);
  if (n < 4 || sd === 0) throw E.div0;
  const m = mean(xs);
  const s4 = xs.reduce((s, x) => s + ((x - m) / sd) ** 4, 0);
  return n * (n + 1) / ((n - 1) * (n - 2) * (n - 3)) * s4 - 3 * (n - 1) ** 2 / ((n - 2) * (n - 3));
});
def('STANDARDIZE', (a, ctx) => { const sd = num(a[2], ctx); if (sd <= 0) throw E.num; return (num(a[0], ctx) - num(a[1], ctx)) / sd; });
def('PERMUT', (a, ctx) => {
  const n = Math.trunc(num(a[0], ctx));
  const k = Math.trunc(num(a[1], ctx));
  if (n < 0 || k < 0 || n < k) throw E.num;
  let p = 1;
  for (let i = 0; i < k; i++) p *= n - i;
  return finite(p);
});
def('FISHER', (a, ctx) => { const x = num(a[0], ctx); if (x <= -1 || x >= 1) throw E.num; return 0.5 * Math.log((1 + x) / (1 - x)); });

// ---- Distributions

/** erfc by W. J. Cody's rational approximations (Math. Comp. 1969), accurate to about 1e-16. */
function erfc(x) {
  const ax = Math.abs(x);
  let r;
  if (ax < 0.5) {
    const t = x * x;
    const p = [3.1611237438705655e0, 1.1386415415105016e2, 3.77485237685302e2, 3.2093775891384694e3, 1.8577770618460315e-1];
    const q = [2.3601290952344122e1, 2.4402463793444417e2, 1.2826165260773723e3, 2.844236833439171e3];
    const num = ((((p[4] * t + p[0]) * t + p[1]) * t + p[2]) * t + p[3]);
    const den = ((((t + q[0]) * t + q[1]) * t + q[2]) * t + q[3]);
    return 1 - x * num / den;
  }
  if (ax < 4) {
    const p = [5.641884969886701e-1, 8.883149794388377e0, 6.611919063714163e1, 2.986351381974001e2, 8.81952221241769e2,
      1.7120476126340707e3, 2.0510783778260716e3, 1.2303393547979972e3, 2.1531153547440383e-8];
    const q = [1.5744926110709835e1, 1.176939508913125e2, 5.371811018620099e2, 1.6213895745666903e3, 3.2907992357334597e3,
      4.362619090143247e3, 3.4393676741437216e3, 1.2303393548037495e3];
    let n = p[8] * ax;
    let d = ax;
    for (let i = 0; i < 7; i++) { n = (n + p[i]) * ax; d = (d + q[i]) * ax; }
    r = (n + p[7]) / (d + q[7]);
  } else {
    const p = [3.0532663496123236e-1, 3.6034489994980445e-1, 1.2578172611122926e-1, 1.6083785148742275e-2, 6.587491615298378e-4,
      1.6315387137302097e-2];
    const q = [2.568520192289822e0, 1.8729528499234673e0, 5.279051029514285e-1, 6.051834131244132e-2, 2.3352049762686918e-3];
    const z = 1 / (ax * ax);
    let n = p[5] * z;
    let d = z;
    for (let i = 0; i < 4; i++) { n = (n + p[i]) * z; d = (d + q[i]) * z; }
    r = z * (n + p[4]) / (d + q[4]);
    r = (1 / Math.sqrt(Math.PI) - r) / ax;
  }
  // (exp(-x²) split so the square loses nothing: x = xs + (x - xs), xs to 1/16.)
  const xs = Math.trunc(ax * 16) / 16;
  r *= Math.exp(-xs * xs) * Math.exp(-(ax - xs) * (ax + xs));
  return x < 0 ? 2 - r : r;
}
const SQRT2 = Math.SQRT2;
const normCdf = (z) => 0.5 * erfc(-z / SQRT2);
const normPdf = (z) => Math.exp(-0.5 * z * z) / Math.sqrt(2 * Math.PI);
/** The inverse standard normal: Acklam's approximation, then Halley's steps on the cdf above. */
function normInv(p) {
  if (p <= 0 || p >= 1) throw E.num;
  const a = [-3.969683028665376e1, 2.209460984245205e2, -2.759285104469687e2, 1.38357751867269e2, -3.066479806614716e1, 2.506628277459239e0];
  const b = [-5.447609879822406e1, 1.615858368580409e2, -1.556989798598866e2, 6.680131188771972e1, -1.328068155288572e1];
  const c = [-7.784894002430293e-3, -3.223964580411365e-1, -2.400758277161838e0, -2.549732539343734e0, 4.374664141464968e0, 2.938163982698783e0];
  const d = [7.784695709041462e-3, 3.224671290700398e-1, 2.445134137142996e0, 3.754408661907416e0];
  let x;
  if (p < 0.02425) { const q = Math.sqrt(-2 * Math.log(p)); x = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1); }
  else if (p > 1 - 0.02425) { const q = Math.sqrt(-2 * Math.log(1 - p)); x = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1); }
  else { const q = p - 0.5; const r = q * q; x = (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1); }
  for (let i = 0; i < 3; i++) {
    const e = normCdf(x) - p;
    const u = e * Math.sqrt(2 * Math.PI) * Math.exp(x * x / 2);
    x -= u / (1 + x * u / 2);
  }
  return x;
}
const cum = (v, ctx) => toBool(scalar(v, ctx));
def('NORM.S.DIST', (a, ctx) => { const z = num(a[0], ctx); return cum(a[1], ctx) ? normCdf(z) : normPdf(z); });
def('NORM.DIST', (a, ctx) => {
  const x = num(a[0], ctx); const m = num(a[1], ctx); const s = num(a[2], ctx);
  if (s <= 0) throw E.num;
  const z = (x - m) / s;
  return cum(a[3], ctx) ? normCdf(z) : normPdf(z) / s;
});
def('NORM.S.INV', (a, ctx) => normInv(num(a[0], ctx)));
def('NORM.INV', (a, ctx) => { const s = num(a[2], ctx); if (s <= 0) throw E.num; return num(a[1], ctx) + s * normInv(num(a[0], ctx)); });
def('EXPON.DIST', (a, ctx) => {
  const x = num(a[0], ctx); const l = num(a[1], ctx);
  if (x < 0 || l <= 0) throw E.num;
  return cum(a[2], ctx) ? -Math.expm1(-l * x) : l * Math.exp(-l * x);
});
const logFact = (n) => { let s = 0; for (let i = 2; i <= n; i++) s += Math.log(i); return s; };
def('POISSON.DIST', (a, ctx) => {
  const k = Math.trunc(num(a[0], ctx)); const m = num(a[1], ctx);
  if (k < 0 || m < 0) throw E.num;
  const pmf = (i) => Math.exp(-m + i * Math.log(m) - logFact(i));
  if (!cum(a[2], ctx)) return pmf(k);
  let s = 0; for (let i = 0; i <= k; i++) s += pmf(i);
  return s;
});
def('BINOM.DIST', (a, ctx) => {
  const k = Math.trunc(num(a[0], ctx)); const n = Math.trunc(num(a[1], ctx)); const p = num(a[2], ctx);
  if (k < 0 || k > n || p < 0 || p > 1) throw E.num;
  // (The number of ways counted exactly, as a product: within Excel's last digit.)
  const ways = (i) => { let x = 1; for (let j = 1; j <= i; j++) x = x * (n - i + j) / j; return x; };
  const pmf = (i) => ways(i) * p ** i * (1 - p) ** (n - i);
  if (!cum(a[3], ctx)) return pmf(k);
  let s = 0; for (let i = 0; i <= k; i++) s += pmf(i);
  return s;
});

// ---- Trigonometry, numerals and bases

def('SINH', (a, ctx) => finite(Math.sinh(num(a[0], ctx))));
def('COSH', (a, ctx) => finite(Math.cosh(num(a[0], ctx))));
def('TANH', (a, ctx) => Math.tanh(num(a[0], ctx)));
const recip = (f) => (a, ctx) => { const v = f(num(a[0], ctx)); if (v === 0) throw E.div0; return 1 / v; };
def('SEC', recip(Math.cos));
def('CSC', recip(Math.sin));
def('COT', recip(Math.tan));
const ROMAN = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
def('ROMAN', (a, ctx) => {
  let n = Math.trunc(num(a[0], ctx));
  if (n < 0 || n > 3999) throw E.value;
  let s = '';
  for (const [v, r] of ROMAN) while (n >= v) { s += r; n -= v; }
  return s;
});
def('ARABIC', (a, ctx) => {
  let s = text(a[0], ctx).trim().toUpperCase();
  const sign = s.startsWith('-') ? -1 : 1;
  if (sign < 0) s = s.slice(1);
  const val = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
  if (!/^[IVXLCDM]*$/.test(s)) throw E.value;
  let n = 0;
  for (let i = 0; i < s.length; i++) n += val[s[i]] < (val[s[i + 1]] || 0) ? -val[s[i]] : val[s[i]];
  return sign * n;
});
def('BASE', (a, ctx) => {
  const n = Math.trunc(num(a[0], ctx)); const radix = Math.trunc(num(a[1], ctx)); const len = given(a[2]) ? Math.trunc(num(a[2], ctx)) : 0;
  if (n < 0 || n >= 2 ** 53 || radix < 2 || radix > 36 || len < 0 || len > 255) throw E.num;
  return n.toString(radix).toUpperCase().padStart(len, '0');
});
def('DECIMAL', (a, ctx) => {
  const s = text(a[0], ctx).trim(); const radix = Math.trunc(num(a[1], ctx));
  if (radix < 2 || radix > 36 || !s || [...s.toLowerCase()].some((ch) => !(parseInt(ch, 36) < radix))) throw E.num;
  return parseInt(s, radix);
});
/** DEC2BIN and DEC2HEX: in two's complement over 10 digits for a negative number, `places` padding a positive one. */
function fromDecimal(a, ctx, radix, bits) {
  const n = Math.trunc(num(a[0], ctx));
  const lim = 2 ** (bits - 1);
  if (n < -lim || n >= lim) throw E.num;
  if (n < 0) return (2 ** bits + n).toString(radix).toUpperCase();
  const s = n.toString(radix).toUpperCase();
  if (!given(a[1])) return s;
  const places = Math.trunc(num(a[1], ctx));
  if (places < s.length || places > 10) throw E.num;
  return s.padStart(places, '0');
}
/** BIN2DEC and HEX2DEC: up to 10 digits, the first bit the sign. */
function toDecimal(a, ctx, radix, bits) {
  const s = toText(scalar(a[0], ctx)).trim();
  if (s.length > 10 || ![...s].every((ch) => parseInt(ch, radix) < radix)) throw E.num;
  const n = s ? parseInt(s, radix) : 0;
  return n >= 2 ** (bits - 1) ? n - 2 ** bits : n;
}
def('DEC2BIN', (a, ctx) => fromDecimal(a, ctx, 2, 10));
def('DEC2HEX', (a, ctx) => fromDecimal(a, ctx, 16, 40));
def('BIN2DEC', (a, ctx) => toDecimal(a, ctx, 2, 10));
def('HEX2DEC', (a, ctx) => toDecimal(a, ctx, 16, 40));
def('MULTINOMIAL', (a, ctx) => {
  const xs = numbers(a, ctx).map(Math.trunc);
  if (xs.some((x) => x < 0)) throw E.num;
  return finite(Math.round(Math.exp(logFact(xs.reduce((s, x) => s + x, 0)) - xs.reduce((s, x) => s + logFact(x), 0))));
});

// ---- A cell's formula, and sheets

const cellOf = (v) => { if (!(v instanceof Area)) throw E.value; return v; };
def('ISFORMULA', (a, ctx) => { const r = cellOf(a[0]); return !!ctx.formulaOf(r.si, r.r0, r.c0); });
def('FORMULATEXT', (a, ctx) => { const r = cellOf(a[0]); const f = ctx.formulaOf(r.si, r.r0, r.c0); if (!f) throw E.na; return `=${displayFormula(f)}`; });
def('SHEET', (a, ctx) => {
  if (!given(a[0])) return ctx.here.si + 1;
  if (a[0] instanceof Area) return a[0].si + 1;
  const si = ctx.sheetIndex(text(a[0], ctx));
  if (si < 0) throw E.na;
  return si + 1;
});
def('SHEETS', (a, ctx) => (given(a[0]) ? 1 : ctx.sheetCount()));

// (For the other distributions, functions-dist.js.)
export { erfc, normCdf, normPdf, normInv, logFact, cum, sampleSd, mean };
