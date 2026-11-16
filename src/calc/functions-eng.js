// Engineering and the rest of the mathematics: bits (BITAND, BITOR, BITXOR, BITLSHIFT, BITRSHIFT), the base conversions
// to and from octal and between binary and hexadecimal, DELTA, GESTEP, ERF and ERFC; the inverse and reciprocal
// hyperbolic functions (ACOSH, ASINH, ATANH, ACOT, ACOTH, COTH, CSCH, SECH); COMBINA, PERMUTATIONA, FACTDOUBLE,
// SERIESSUM, SQRTPI; DAYS360; HYPERLINK and AREAS; and money (FVSCHEDULE, PDURATION, RRI, ISPMT, MIRR, SYD, VDB,
// DOLLARDE, DOLLARFR, TBILLEQ, TBILLPRICE, TBILLYIELD). As Excel defines them, measured against Excel with the formula
// probe (test/formulas-excel.json). Registered into functions.js's table; evaluate.js imports this file.
import { Area, Union, E, finite, scalar, toText, matrix } from './values.js';
import { def, given, num, text, numbers } from './functions.js';
import { erfc } from './functions-stats.js';
import { dateParts } from '../ooxml/format.js';

// ---- Bits

const whole48 = (v, ctx) => { const n = num(v, ctx); if (n < 0 || n >= 2 ** 48 || !Number.isInteger(n)) throw E.num; return n; };
/** A 48-bit operation by halves (JavaScript's bit operators are 32-bit). */
const bits = (f) => (a, ctx) => {
  const x = whole48(a[0], ctx);
  const y = whole48(a[1], ctx);
  const hi = (v) => Math.floor(v / 2 ** 24);
  const lo = (v) => v % 2 ** 24;
  return f(hi(x), hi(y)) * 2 ** 24 + f(lo(x), lo(y));
};
def('BITAND', bits((x, y) => (x & y) >>> 0));
def('BITOR', bits((x, y) => (x | y) >>> 0));
def('BITXOR', bits((x, y) => (x ^ y) >>> 0));
const shift = (dir) => (a, ctx) => {
  const x = whole48(a[0], ctx);
  const s = Math.trunc(num(a[1], ctx));
  if (Math.abs(s) > 53) throw E.num;
  const by = dir * s;
  const v = by >= 0 ? x * 2 ** by : Math.floor(x / 2 ** -by);
  if (v >= 2 ** 48) throw E.num;
  return v;
};
def('BITLSHIFT', shift(1));
def('BITRSHIFT', shift(-1));

// ---- Bases (10 digits; a negative number in two's complement over them)

const BITS = { 2: 10, 8: 30, 16: 40 };
/** A number written in base `from` (up to 10 digits, the first bit the sign) as a number. */
function readBase(v, ctx, from) {
  const s = toText(scalar(v, ctx)).trim();
  if (s.length > 10 || ![...s].every((ch) => parseInt(ch, from) < from)) throw E.num;
  const n = s ? parseInt(s, from) : 0;
  const b = BITS[from];
  return n >= 2 ** (b - 1) ? n - 2 ** b : n;
}
/** A number written in base `to`, as Excel's: two's complement over 10 digits when negative, else padded to `places`. */
function writeBase(n, to, places) {
  const b = BITS[to];
  if (n < -(2 ** (b - 1)) || n >= 2 ** (b - 1)) throw E.num;
  if (n < 0) return (2 ** b + n).toString(to).toUpperCase();
  const s = n.toString(to).toUpperCase();
  if (places == null) return s;
  if (places < s.length || places > 10) throw E.num;
  return s.padStart(places, '0');
}
const conv = (from, to) => (a, ctx) => {
  const n = from === 10 ? Math.trunc(num(a[0], ctx)) : readBase(a[0], ctx, from);
  if (to === 10) return n;
  return writeBase(n, to, given(a[1]) ? Math.trunc(num(a[1], ctx)) : null);
};
def('OCT2DEC', conv(8, 10));
def('DEC2OCT', conv(10, 8));
def('BIN2HEX', conv(2, 16));
def('HEX2BIN', conv(16, 2));
def('BIN2OCT', conv(2, 8));
def('OCT2BIN', conv(8, 2));
def('HEX2OCT', conv(16, 8));
def('OCT2HEX', conv(8, 16));

// ---- DELTA, GESTEP, ERF, ERFC

def('DELTA', (a, ctx) => (num(a[0], ctx) === (given(a[1]) ? num(a[1], ctx) : 0) ? 1 : 0));
def('GESTEP', (a, ctx) => (num(a[0], ctx) >= (given(a[1]) ? num(a[1], ctx) : 0) ? 1 : 0));
const erf = (x) => 1 - erfc(x);
def('ERF', (a, ctx) => { const lo = num(a[0], ctx); return given(a[1]) ? erf(num(a[1], ctx)) - erf(lo) : erf(lo); });
def('ERF.PRECISE', (a, ctx) => erf(num(a[0], ctx)));
def('ERFC ERFC.PRECISE', (a, ctx) => erfc(num(a[0], ctx)));

// ---- Hyperbolic and reciprocal

def('ACOSH', (a, ctx) => { const x = num(a[0], ctx); if (x < 1) throw E.num; return Math.acosh(x); });
def('ASINH', (a, ctx) => Math.asinh(num(a[0], ctx)));
def('ATANH', (a, ctx) => { const x = num(a[0], ctx); if (x <= -1 || x >= 1) throw E.num; return Math.atanh(x); });
def('ACOT', (a, ctx) => { const x = num(a[0], ctx); return x === 0 ? Math.PI / 2 : Math.atan(1 / x) + (x < 0 ? Math.PI : 0); });
def('ACOTH', (a, ctx) => { const x = num(a[0], ctx); if (Math.abs(x) <= 1) throw E.num; return 0.5 * Math.log((x + 1) / (x - 1)); });
def('COTH', (a, ctx) => { const x = num(a[0], ctx); if (x === 0) throw E.div0; return 1 / Math.tanh(x); });
def('CSCH', (a, ctx) => { const x = num(a[0], ctx); if (x === 0) throw E.div0; return 1 / Math.sinh(x); });
def('SECH', (a, ctx) => 1 / Math.cosh(num(a[0], ctx)));

// ---- Counting and series

const choose = (n, k) => { let x = 1; for (let j = 1; j <= k; j++) x = x * (n - k + j) / j; return Math.round(x); };
def('COMBINA', (a, ctx) => {
  const n = Math.trunc(num(a[0], ctx)); const k = Math.trunc(num(a[1], ctx));
  if (n < 0 || k < 0 || (n === 0 && k > 0)) throw E.num;
  return k === 0 ? 1 : finite(choose(n + k - 1, k));
});
def('PERMUTATIONA', (a, ctx) => { const n = Math.trunc(num(a[0], ctx)); const k = Math.trunc(num(a[1], ctx)); if (n < 0 || k < 0) throw E.num; return finite(n ** k); });
def('FACTDOUBLE', (a, ctx) => {
  const n = Math.trunc(num(a[0], ctx));
  if (n < -1) throw E.num;
  let p = 1;
  for (let i = n; i > 1; i -= 2) p *= i;
  return finite(p);
});
def('SERIESSUM', (a, ctx) => {
  const x = num(a[0], ctx); const n = num(a[1], ctx); const m = num(a[2], ctx);
  const cs = matrix(a[3], ctx).flat();
  if (cs.some((c) => typeof c !== 'number')) throw E.value;
  return finite(cs.reduce((s, c, i) => s + c * x ** (n + i * m), 0));
});
def('SQRTPI', (a, ctx) => { const x = num(a[0], ctx); if (x < 0) throw E.num; return Math.sqrt(x * Math.PI); });

// ---- DAYS360

// (Read in Excel's calendar, February 1900 with its 29th, serial 60, its last: DAYS360(59,61) is 3, DAYS360(60,61)
// 1, DAYS360(31,60,TRUE) 29, measured.)
const ymd = (serial, ctx) => { const p = dateParts(Math.floor(serial), ctx.date1904); return [p.y, p.m, p.d]; };
const lastOfFeb = (y, m, d) => m === 2 && d === (y === 1900 ? 29 : new Date(Date.UTC(y, 2, 0)).getUTCDate());
def('DAYS360', (a, ctx) => {
  let [y1, m1, d1] = ymd(num(a[0], ctx), ctx);
  let [y2, m2, d2] = ymd(num(a[1], ctx), ctx);
  const european = given(a[2]) && !!scalar(a[2], ctx);
  if (european) { if (d1 === 31) d1 = 30; if (d2 === 31) d2 = 30; }
  else {
    // (The US method as Excel has it: the last of February as the 30th when the start is (the end left as it is, even
    // when it is the last of February too); the 31st of the end month as the 30th when the start is the 30th or 31st.)
    if (lastOfFeb(y1, m1, d1)) d1 = 30;
    if (d2 === 31 && d1 >= 30) d2 = 30;
    if (d1 === 31) d1 = 30;
  }
  return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1);
});

// ---- HYPERLINK, AREAS

def('HYPERLINK', (a, ctx) => (given(a[1]) ? scalar(a[1], ctx) : text(a[0], ctx)));
def('AREAS', (a) => {
  if (a[0] instanceof Union) return a[0].areas.length;
  if (a[0] instanceof Area) return 1;
  throw E.value;
});

// ---- Money

def('FVSCHEDULE', (a, ctx) => numbers([a[1]], ctx).reduce((p, r) => p * (1 + r), num(a[0], ctx)));
def('PDURATION', (a, ctx) => { const r = num(a[0], ctx); const pv = num(a[1], ctx); const fv = num(a[2], ctx); if (r <= 0 || pv <= 0 || fv <= 0) throw E.num; return (Math.log(fv) - Math.log(pv)) / Math.log(1 + r); });
def('RRI', (a, ctx) => { const n = num(a[0], ctx); const pv = num(a[1], ctx); const fv = num(a[2], ctx); if (n <= 0 || pv === 0) throw E.num; return finite((fv / pv) ** (1 / n) - 1); });
def('ISPMT', (a, ctx) => { const r = num(a[0], ctx); const per = num(a[1], ctx); const n = num(a[2], ctx); const pv = num(a[3], ctx); if (n === 0) throw E.div0; return pv * r * (per / n - 1); });
def('MIRR', (a, ctx) => {
  const flows = numbers([a[0]], ctx);
  const fin = num(a[1], ctx); const re = num(a[2], ctx);
  const n = flows.length;
  // (As Excel defines it: the NPVs of the positive and of the negative flows, (-NPV(re, pos) (1+re)^n /
  // (NPV(fin, neg) (1+fin)))^(1/(n-1)) - 1.)
  const npv = (r, keep) => flows.reduce((s, v, i) => s + (keep(v) ? v / (1 + r) ** (i + 1) : 0), 0);
  const pos = npv(re, (v) => v > 0);
  const neg = npv(fin, (v) => v < 0);
  if (pos === 0 || neg === 0) throw E.div0;
  return ((-pos * (1 + re) ** n) / (neg * (1 + fin))) ** (1 / (n - 1)) - 1;
});
def('SYD', (a, ctx) => {
  const cost = num(a[0], ctx); const salvage = num(a[1], ctx); const life = num(a[2], ctx); const per = num(a[3], ctx);
  if (life <= 0 || per <= 0 || per > life) throw E.num;
  return ((cost - salvage) * (life - per + 1) * 2) / (life * (life + 1));
});
/** VDB: double-declining (or `factor`) balance from start to end, switching to straight line when that is more,
 *  unless `noSwitch`; parts of periods prorated. */
def('VDB', (a, ctx) => {
  const cost = num(a[0], ctx); const salvage = num(a[1], ctx); const life = num(a[2], ctx);
  const start = num(a[3], ctx); const end = num(a[4], ctx);
  const factor = given(a[5]) ? num(a[5], ctx) : 2;
  const noSwitch = given(a[6]) && !!scalar(a[6], ctx);
  if (cost < 0 || salvage < 0 || life <= 0 || start < 0 || end < start || end > life || factor <= 0) throw E.num;
  // Each whole period's depreciation.
  const per = [];
  let book = cost;
  let straight = false;
  for (let p = 0; p < Math.ceil(life); p++) {
    let d = Math.min(book * factor / life, Math.max(0, book - salvage));
    if (!noSwitch) {
      const left = life - p;
      const sl = left > 0 ? (book - salvage) / left : 0;
      if (straight || sl > d) { straight = true; d = Math.max(0, sl); }
    }
    per.push(d);
    book -= d;
  }
  let total = 0;
  for (let p = Math.floor(start); p < Math.ceil(end); p++) {
    const from = Math.max(start, p);
    const to = Math.min(end, p + 1);
    total += (per[p] || 0) * (to - from);
  }
  return total;
});
def('DOLLARDE', (a, ctx) => {
  const x = num(a[0], ctx); const f = Math.trunc(num(a[1], ctx));
  if (f < 0) throw E.num; if (f === 0) throw E.div0;
  const whole = Math.trunc(x);
  const digits = Math.ceil(Math.log10(f + 1e-12)) || 1;
  return whole + (x - whole) * 10 ** digits / f;
});
def('DOLLARFR', (a, ctx) => {
  const x = num(a[0], ctx); const f = Math.trunc(num(a[1], ctx));
  if (f < 0) throw E.num; if (f === 0) throw E.div0;
  const whole = Math.trunc(x);
  const digits = Math.ceil(Math.log10(f + 1e-12)) || 1;
  return whole + (x - whole) * f / 10 ** digits;
});
/** A Treasury bill's days from settlement to maturity (at most a year), as Excel checks them. */
const tbill = (a, ctx) => {
  const s = Math.trunc(num(a[0], ctx)); const m = Math.trunc(num(a[1], ctx));
  const days = m - s;
  if (days <= 0 || days > 366) throw E.num;
  return days;
};
def('TBILLPRICE', (a, ctx) => { const d = tbill(a, ctx); const r = num(a[2], ctx); if (r <= 0) throw E.num; const p = 100 * (1 - r * d / 360); if (p <= 0) throw E.num; return p; });
def('TBILLYIELD', (a, ctx) => { const d = tbill(a, ctx); const p = num(a[2], ctx); if (p <= 0) throw E.num; return (100 - p) / p * 360 / d; });
def('TBILLEQ', (a, ctx) => {
  const d = tbill(a, ctx); const r = num(a[2], ctx);
  if (r <= 0) throw E.num;
  if (d <= 182) return 365 * r / (360 - r * d);
  const price = 100 * (1 - r * d / 360);
  const t = d / 365;
  return (-2 * t + 2 * Math.sqrt(t * t - (2 * t - 1) * (1 - 100 / price))) / (2 * t - 1);
});
