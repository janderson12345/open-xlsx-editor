// Loans, depreciation and rates (IPMT, PPMT, CUMIPMT, CUMPRINC, SLN, DDB, DB, EFFECT, NOMINAL, XNPV, XIRR), and times and
// working days (TIMEVALUE, WORKDAY.INTL, NETWORKDAYS.INTL), as Excel defines them. Measured against Excel with the
// formula probe (test/formulas-excel.json). Registered into functions.js's table; evaluate.js imports this file.
import { E, finite } from './values.js';
import { def, given, num, text, numbers } from './functions.js';
import { pmt, grow } from './functions-more.js';
import { parseInput } from '../ooxml/input.js';
import { serialOf } from './functions-text.js';

// ---- Loans

/** A future value with (1 + rate)^n by exp and log1p: Excel's IPMT and CUMPRINC agree with it to their last digits. */
const fv = (rate, n, p, pv, type) => { const g = Math.exp(n * Math.log1p(rate)); return rate === 0 ? -(pv + p * n) : -(pv * g + p * (1 + rate * type) * (g - 1) / rate); };
/** The interest part of payment `per` (Excel's IPMT): the balance after the payments before it, times the rate. */
function ipmt(rate, per, nper, pv, f, type) {
  const p = pmt(rate, nper, pv, f, type);
  if (per === 1) return type === 1 ? 0 : -pv * rate;
  return type === 1 ? (fv(rate, per - 2, p, pv, 1) - p) * rate : fv(rate, per - 1, p, pv, 0) * rate;
}
const loan = (a, ctx) => {
  const rate = num(a[0], ctx);
  const per = Math.trunc(num(a[1], ctx));
  const nper = num(a[2], ctx);
  const pv = num(a[3], ctx);
  const f = given(a[4]) ? num(a[4], ctx) : 0;
  const type = given(a[5]) && num(a[5], ctx) !== 0 ? 1 : 0;
  if (per < 1 || per > nper) throw E.num;
  return { rate, per, nper, pv, f, type };
};
def('IPMT', (a, ctx) => { const l = loan(a, ctx); return finite(ipmt(l.rate, l.per, l.nper, l.pv, l.f, l.type)); });
def('PPMT', (a, ctx) => { const l = loan(a, ctx); return finite(pmt(l.rate, l.nper, l.pv, l.f, l.type) - ipmt(l.rate, l.per, l.nper, l.pv, l.f, l.type)); });
/** CUMIPMT and CUMPRINC: the interest (or principal) of payments start to end. */
function cumulative(a, ctx, principal) {
  const rate = num(a[0], ctx);
  const nper = num(a[1], ctx);
  const pv = num(a[2], ctx);
  const start = Math.trunc(num(a[3], ctx));
  const end = Math.trunc(num(a[4], ctx));
  const type = num(a[5], ctx);
  if (rate <= 0 || nper <= 0 || pv <= 0 || start < 1 || end < start || end > nper || (type !== 0 && type !== 1)) throw E.num;
  const p = pmt(rate, nper, pv, 0, type);
  let sum = 0;
  for (let k = start; k <= end; k++) { const i = ipmt(rate, k, nper, pv, 0, type); sum += principal ? p - i : i; }
  return finite(sum);
}
def('CUMIPMT', (a, ctx) => cumulative(a, ctx, false));
def('CUMPRINC', (a, ctx) => cumulative(a, ctx, true));

// ---- Depreciation

def('SLN', (a, ctx) => { const life = num(a[2], ctx); if (life === 0) throw E.div0; return (num(a[0], ctx) - num(a[1], ctx)) / life; });
def('DDB', (a, ctx) => {
  const cost = num(a[0], ctx);
  const salvage = num(a[1], ctx);
  const life = num(a[2], ctx);
  const period = num(a[3], ctx);
  const factor = given(a[4]) ? num(a[4], ctx) : 2;
  if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || period > life || factor <= 0) throw E.num;
  // (The book value at the period's start by repeated squaring, as Excel works it out to its last digit.)
  const book = cost * grow(-factor / life, Math.ceil(period) - 1);
  return Math.min(book * factor / life, Math.max(0, book - salvage));
});
def('DB', (a, ctx) => {
  const cost = num(a[0], ctx);
  const salvage = num(a[1], ctx);
  const life = Math.trunc(num(a[2], ctx));
  const period = Math.trunc(num(a[3], ctx));
  const month = given(a[4]) ? Math.trunc(num(a[4], ctx)) : 12;
  if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || month < 1 || month > 12 || period > life + (month < 12 ? 1 : 0)) throw E.num;
  if (cost === 0) return 0;
  // (The rate rounded to three places, as Excel rounds it.)
  const rate = Math.round((1 - (salvage / cost) ** (1 / life)) * 1000) / 1000;
  let total = cost * rate * month / 12;
  if (period === 1) return total;
  let dep = 0;
  for (let k = 2; k <= period; k++) {
    dep = k === life + 1 ? (cost - total) * rate * (12 - month) / 12 : (cost - total) * rate;
    total += dep;
  }
  return dep;
});

// ---- Rates

def('EFFECT', (a, ctx) => {
  const r = num(a[0], ctx);
  const n = Math.trunc(num(a[1], ctx));
  if (r <= 0 || n < 1) throw E.num;
  return grow(r / n, n) - 1;
});
def('NOMINAL', (a, ctx) => {
  const r = num(a[0], ctx);
  const n = Math.trunc(num(a[1], ctx));
  if (r <= 0 || n < 1) throw E.num;
  return n * ((1 + r) ** (1 / n) - 1);
});
/** XNPV's and XIRR's flows and their dates (whole days), checked as Excel checks them. */
function flows(a, ctx, from) {
  const v = numbers([a[from]], ctx);
  const d = numbers([a[from + 1]], ctx).map(Math.trunc);
  if (v.length !== d.length || !v.length) throw E.num;
  if (d.some((x) => x < d[0])) throw E.num;
  return { v, d };
}
const xnpv = (rate, { v, d }) => v.reduce((s, x, k) => s + x / (1 + rate) ** ((d[k] - d[0]) / 365), 0);
def('XNPV', (a, ctx) => finite(xnpv(num(a[0], ctx), flows(a, ctx, 1))));
def('XIRR', (a, ctx) => {
  const f = flows(a, ctx, 0);
  if (!f.v.some((x) => x > 0) || !f.v.some((x) => x < 0)) throw E.num;
  let r = given(a[2]) ? num(a[2], ctx) : 0.1;
  for (let i = 0; i < 100; i++) {
    const y = xnpv(r, f);
    const dy = f.v.reduce((s, x, k) => { const t = (f.d[k] - f.d[0]) / 365; return s - t * x / (1 + r) ** (t + 1); }, 0);
    if (!Number.isFinite(dy) || dy === 0) break;
    const next = r - y / dy;
    if (!Number.isFinite(next) || next <= -1) break;
    if (Math.abs(next - r) < 1e-12) return next;
    r = next;
  }
  throw E.num;
});

// ---- Times and working days

def('TIMEVALUE', (a, ctx) => {
  // (Read as an entry: a time, or a date with a time, whose fraction of a day it gives.)
  const p = parseInput(text(a[0], ctx), { date1904: ctx.date1904 });
  if (!['time', 'datetime', 'date'].includes(p.kind)) throw E.value;
  return p.value - Math.floor(p.value);
});
/** A weekend as Excel's .INTL functions take it: a number (1 to 7 two days, 11 to 17 one) or seven 0s and 1s from
 *  Monday. Returns a test of a serial date. */
function weekendOf(v, ctx) {
  const days = new Array(7).fill(false); // Monday first
  if (!given(v)) { days[5] = days[6] = true; } else if (typeof v === 'string' || (v && typeof v.valueOf() === 'string')) {
    const s = text(v, ctx);
    if (!/^[01]{7}$/.test(s) || s === '1111111') throw E.value;
    [...s].forEach((ch, k) => { days[k] = ch === '1'; });
  } else {
    const n = Math.trunc(num(v, ctx));
    if (n >= 1 && n <= 7) { days[(n + 4) % 7] = true; days[(n + 5) % 7] = true; } else if (n >= 11 && n <= 17) days[(n - 11 + 6) % 7] = true;
    else throw E.num;
  }
  // (Serial 1 is a Sunday: Monday-first index (serial + 5) % 7.)
  return (serial) => days[(Math.floor(serial) + 5) % 7];
}
const holidays = (v, ctx) => (given(v) ? new Set(numbers([v], ctx).map(Math.trunc)) : new Set());
def('NETWORKDAYS.INTL', (a, ctx) => {
  let s = Math.trunc(serialOf(a[0], ctx));
  let e = Math.trunc(serialOf(a[1], ctx));
  const off = weekendOf(a[2], ctx);
  const hol = holidays(a[3], ctx);
  const sign = s <= e ? 1 : -1;
  if (sign < 0) [s, e] = [e, s];
  let n = 0;
  for (let d = s; d <= e; d++) if (!off(d) && !hol.has(d)) n++;
  return sign * n;
});
def('WORKDAY.INTL', (a, ctx) => {
  let d = Math.trunc(serialOf(a[0], ctx));
  let k = Math.trunc(num(a[1], ctx));
  const off = weekendOf(a[2], ctx);
  const hol = holidays(a[3], ctx);
  const step = k < 0 ? -1 : 1;
  while (k !== 0) { d += step; if (!off(d) && !hol.has(d)) k -= step; }
  return d;
});
