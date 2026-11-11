// Securities, as Excel's Financial category has them: the coupon dates and days (COUPPCD, COUPNCD, COUPNUM, COUPDAYBS,
// COUPDAYS, COUPDAYSNC), a bond's PRICE, YIELD, DURATION and MDURATION, ACCRINT and ACCRINTM, and the discounted and
// at-maturity ones (DISC, INTRATE, RECEIVED, PRICEDISC, YIELDDISC, PRICEMAT, YIELDMAT). Coupon dates step back from
// maturity by 12/frequency months, kept at the month's end when maturity is; days are counted by the basis (0 US
// 30/360, 1 actual/actual, 2 actual/360, 3 actual/365, 4 European 30/360). Measured against Excel with the formula
// probe (test/formulas-excel.json).
import { E } from './values.js';
import { FUNCTIONS, def, given, num } from './functions.js';
import { dateSerial } from './values.js';
import { dateParts } from '../ooxml/format.js';

const leap = (y) => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
const monthDays = (y, m) => [31, leap(y) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1];

/** The arguments every security takes: whole-day dates, frequency 1, 2 or 4 and basis 0 to 4. */
const day = (v, ctx) => { const n = Math.trunc(num(v, ctx)); if (n < 0) throw E.num; return n; };
const basisOf = (a, i, ctx) => { const b = given(a[i]) ? Math.trunc(num(a[i], ctx)) : 0; if (b < 0 || b > 4) throw E.num; return b; };
const freqOf = (v, ctx) => { const f = Math.trunc(num(v, ctx)); if (f !== 1 && f !== 2 && f !== 4) throw E.num; return f; };

/** Days from s to e as `basis` counts them. */
function daysBetween(s, e, basis, ctx) {
  if (basis !== 0 && basis !== 4) return e - s;
  const p = dateParts(s, ctx.date1904);
  const q = dateParts(e, ctx.date1904);
  let d1 = p.d; let d2 = q.d;
  if (basis === 0) {
    // (US: the last of February as the 30th, the end's too when the start's is; the 31st as the 30th.)
    const lastFeb = (x) => x.m === 2 && x.d === monthDays(x.y, 2);
    if (lastFeb(p) && lastFeb(q)) d2 = 30;
    if (lastFeb(p)) d1 = 30;
    if (d2 === 31 && d1 >= 30) d2 = 30;
    if (d1 === 31) d1 = 30;
  } else { if (d1 === 31) d1 = 30; if (d2 === 31) d2 = 30; }
  return (q.y - p.y) * 360 + (q.m - p.m) * 30 + (d2 - d1);
}

/** The date `months` from serial `from` (back when negative), at the month's end when `from` is at one (unless
 *  `keepDay`: the same day, or the month's last when it is shorter). */
function addMonths(from, months, ctx, keepDay = false) {
  const p = dateParts(from, ctx.date1904);
  const eom = !keepDay && p.d === monthDays(p.y, p.m);
  const t = p.y * 12 + (p.m - 1) + months;
  const y = Math.floor(t / 12); const m = t - y * 12 + 1;
  return dateSerial(y, m, eom ? monthDays(y, m) : Math.min(p.d, monthDays(y, m)), ctx.date1904);
}

/** The coupon dates around settlement: { pcd, ncd, n } (n the coupons from settlement to maturity). */
function coupons(settle, mat, freq, ctx) {
  if (settle >= mat) throw E.num;
  const step = 12 / freq;
  let n = 1;
  while (addMonths(mat, -n * step, ctx) > settle) n += 1;
  return { pcd: addMonths(mat, -n * step, ctx), ncd: addMonths(mat, -(n - 1) * step, ctx), n };
}
/** { A, E, DSC, toNext, N }: days from the coupon before settlement, days of the period, days to the next coupon
 *  (COUPDAYSNC's), the same as PRICE, YIELD and DURATION take them, and the coupons left. Excel's, measured: with basis
 *  1 the period is the coupon date before settlement to the same day 12/frequency months on (Feb 29 to Aug 29: 182, not
 *  the 184 to Aug 31); with 2 and 3 a fixed 360 or 365 over the frequency. COUPDAYSNC counts 30/360 days for basis 4 and
 *  takes E - A for basis 0. */
function couponDays(settle, mat, freq, basis, ctx) {
  const c = coupons(settle, mat, freq, ctx);
  const A = daysBetween(c.pcd, settle, basis, ctx);
  const Eday = basis === 1 ? addMonths(c.pcd, 12 / freq, ctx, true) - c.pcd : basis === 3 ? 365 / freq : 360 / freq;
  const DSC = basis === 0 ? Eday - A : basis === 4 ? daysBetween(settle, c.ncd, 4, ctx) : c.ncd - settle;
  // (PRICE and its kin: basis 1 the actual coupon period and days to the next coupon; the others E - A.)
  const period = basis === 1 ? c.ncd - c.pcd : Eday;
  const toNext = basis === 1 ? c.ncd - settle : Eday - A;
  return { A, E: Eday, period, DSC, toNext, N: c.n, ...c };
}
const securityArgs = (a, ctx, fi, bi) => {
  const settle = day(a[0], ctx); const mat = day(a[1], ctx);
  const freq = freqOf(a[fi], ctx); const basis = basisOf(a, bi, ctx);
  if (settle >= mat) throw E.num;
  return { settle, mat, freq, basis };
};

def('COUPPCD', (a, ctx) => { const s = securityArgs(a, ctx, 2, 3); return coupons(s.settle, s.mat, s.freq, ctx).pcd; });
def('COUPNCD', (a, ctx) => { const s = securityArgs(a, ctx, 2, 3); return coupons(s.settle, s.mat, s.freq, ctx).ncd; });
def('COUPNUM', (a, ctx) => { const s = securityArgs(a, ctx, 2, 3); return coupons(s.settle, s.mat, s.freq, ctx).n; });
def('COUPDAYBS', (a, ctx) => { const s = securityArgs(a, ctx, 2, 3); return couponDays(s.settle, s.mat, s.freq, s.basis, ctx).A; });
def('COUPDAYS', (a, ctx) => { const s = securityArgs(a, ctx, 2, 3); return couponDays(s.settle, s.mat, s.freq, s.basis, ctx).E; });
def('COUPDAYSNC', (a, ctx) => { const s = securityArgs(a, ctx, 2, 3); return couponDays(s.settle, s.mat, s.freq, s.basis, ctx).DSC; });

/** A bond's price per 100 face value at yield `yld`. */
function bondPrice(cd, rate, yld, red, freq) {
  const c = (100 * rate) / freq;
  const accrued = (c * cd.A) / cd.period;
  if (cd.N === 1) return (red + c) / (1 + ((cd.toNext / cd.period) * yld) / freq) - accrued;
  const v = 1 + yld / freq;
  const t = cd.toNext / cd.period;
  let p = red / v ** (cd.N - 1 + t);
  for (let k = 1; k <= cd.N; k++) p += c / v ** (k - 1 + t);
  return p - accrued;
}
def('PRICE', (a, ctx) => {
  const s = securityArgs(a, ctx, 5, 6);
  const rate = num(a[2], ctx); const yld = num(a[3], ctx); const red = num(a[4], ctx);
  if (rate < 0 || yld < 0 || red <= 0) throw E.num;
  return bondPrice(couponDays(s.settle, s.mat, s.freq, s.basis, ctx), rate, yld, red, s.freq);
});
def('YIELD', (a, ctx) => {
  const s = securityArgs(a, ctx, 5, 6);
  const rate = num(a[2], ctx); const pr = num(a[3], ctx); const red = num(a[4], ctx);
  if (rate < 0 || pr <= 0 || red <= 0) throw E.num;
  const cd = couponDays(s.settle, s.mat, s.freq, s.basis, ctx);
  if (cd.N === 1) {
    // (One coupon left: Excel's closed form.)
    const c = rate / s.freq;
    const paid = pr / 100 + (cd.A / cd.period) * c;
    return ((red / 100 + c - paid) / paid) * ((s.freq * cd.period) / cd.toNext);
  }
  // (More: the yield whose price is pr, by Newton's method from the coupon rate, bisection if it strays.)
  let y = rate || 0.05;
  for (let i = 0; i < 100; i++) {
    const f = bondPrice(cd, rate, y, red, s.freq) - pr;
    const h = 1e-7 * Math.max(1, Math.abs(y));
    const d = (bondPrice(cd, rate, y + h, red, s.freq) - bondPrice(cd, rate, y - h, red, s.freq)) / (2 * h);
    const next = y - f / d;
    if (!Number.isFinite(next)) throw E.num;
    if (Math.abs(next - y) < 1e-14 * Math.max(1, Math.abs(y))) return next;
    y = next;
  }
  return y;
});
/** Macaulay duration in years. */
function duration(a, ctx) {
  const s = securityArgs(a, ctx, 4, 5);
  const coupon = num(a[2], ctx); const yld = num(a[3], ctx);
  if (coupon < 0 || yld < 0) throw E.num;
  const cd = couponDays(s.settle, s.mat, s.freq, s.basis, ctx);
  const v = 1 + yld / s.freq;
  const c = (100 * coupon) / s.freq;
  const t0 = cd.toNext / cd.period;
  let weighted = 0; let total = 0;
  for (let k = 1; k <= cd.N; k++) {
    const t = t0 + k - 1;
    const cash = c + (k === cd.N ? 100 : 0);
    weighted += (t * cash) / v ** t;
    total += cash / v ** t;
  }
  return { d: weighted / total / s.freq, v };
}
def('DURATION', (a, ctx) => duration(a, ctx).d);
def('MDURATION', (a, ctx) => { const r = duration(a, ctx); return r.d / r.v; });

/** ACCRINT, as Excel works it out (measured): par x rate / frequency x the coupon periods from issue to settlement,
 *  stepped from first_interest, each whole one (ended before settlement) counting 1 and a part its days (by the
 *  basis) over the period's length:
 *  360 or 365 over the frequency, or with basis 1 the days of the period that ends at first_interest. With calc_method
 *  FALSE, the days from issue to settlement over that length, whole periods or not. (With basis 1 over many periods
 *  Excel's last digits differ from these: docs/CALC.md.) */
def('ACCRINT', (a, ctx) => {
  const issue = day(a[0], ctx); const first = day(a[1], ctx); const settle = day(a[2], ctx);
  const rate = num(a[3], ctx); const par = given(a[4]) ? num(a[4], ctx) : 1000;
  const freq = freqOf(a[5], ctx); const basis = basisOf(a, 6, ctx);
  const byPeriod = !given(a[7]) || !!num(a[7], ctx);
  if (rate <= 0 || par <= 0 || issue >= settle) throw E.num;
  const step = 12 / freq;
  const length = basis === 1 ? first - addMonths(first, -step, ctx) : basis === 3 ? 365 / freq : 360 / freq;
  let total = 0;
  if (!byPeriod) total = daysBetween(issue, settle, basis, ctx) / length;
  else {
    let k = 0;
    while (addMonths(first, -k * step, ctx) > issue) k += 1;
    for (let i = -k; ; i++) {
      const p0 = addMonths(first, i * step, ctx); const p1 = addMonths(first, (i + 1) * step, ctx);
      if (p0 >= settle) break;
      const from = Math.max(p0, issue); const to = Math.min(p1, settle);
      // (Whole: from its start to before settlement. A period ending on settlement is a part, by its days: August 31
      // to February 28 is 178 days of 180 in 30/360.)
      const whole = from === p0 && p1 < settle;
      if (to > from) total += whole ? 1 : daysBetween(from, to, basis, ctx) / length;
    }
  }
  return (par * rate * total) / freq;
});

const yearFrac = (s, e, basis, ctx) => FUNCTIONS.YEARFRAC.fn([s, e, basis], ctx);
/** The days in a year as `basis` has it: for 1, the year up to settlement's (365, or 366 over a February 29). */
const yearDays = (settle, basis, ctx) => (basis === 1 ? settle - addMonths(settle, -12, ctx, true) : basis === 3 ? 365 : 360);
/** The arguments of the discounted securities: settlement before maturity, the basis, the days between (DSM) and the
 *  year's (B), and DSM / B. */
const discArgs = (a, ctx, bi) => {
  const settle = day(a[0], ctx); const mat = day(a[1], ctx); const basis = basisOf(a, bi, ctx);
  if (settle >= mat) throw E.num;
  const dsm = daysBetween(settle, mat, basis, ctx); const B = yearDays(settle, basis, ctx);
  // (DISC and its kin, basis 1: YEARFRAC's year, 366 over a February 29.)
  const yf = basis === 1 ? yearFrac(settle, mat, basis, ctx) : dsm / B;
  return { settle, mat, basis, dsm, B: basis === 1 ? dsm / yf : B, yf, yearB: B };
};
const positive = (v, ctx) => { const n = num(v, ctx); if (n <= 0) throw E.num; return n; };
def('ACCRINTM', (a, ctx) => {
  const issue = day(a[0], ctx); const settle = day(a[1], ctx); const rate = positive(a[2], ctx);
  const par = given(a[3]) ? positive(a[3], ctx) : 1000; const basis = basisOf(a, 4, ctx);
  if (issue >= settle) throw E.num;
  return par * rate * yearFrac(issue, settle, basis, ctx);
});
def('DISC', (a, ctx) => { const s = discArgs(a, ctx, 4); const pr = positive(a[2], ctx); const red = positive(a[3], ctx); return (1 - pr / red) * (s.B / s.dsm); });
def('INTRATE', (a, ctx) => { const s = discArgs(a, ctx, 4); const inv = positive(a[2], ctx); const red = positive(a[3], ctx); return (red - inv) / inv / s.yf; });
def('RECEIVED', (a, ctx) => {
  const s = discArgs(a, ctx, 4); const inv = positive(a[2], ctx); const disc = positive(a[3], ctx);
  const d = 1 - disc * s.yf;
  if (d <= 0) throw E.num;
  return inv / d;
});
def('PRICEDISC', (a, ctx) => { const s = discArgs(a, ctx, 4); const disc = positive(a[2], ctx); const red = positive(a[3], ctx); return red - disc * red * s.yf; });
def('YIELDDISC', (a, ctx) => { const s = discArgs(a, ctx, 4); const pr = positive(a[2], ctx); const red = positive(a[3], ctx); return (red - pr) / pr / s.yf; });
/** At maturity: { dim, a, dsm } as fractions of a year (issue to maturity, issue to settlement, settlement on). */
const atMaturity = (a, ctx) => {
  const s = discArgs(a, ctx, 5);
  const issue = day(a[2], ctx);
  if (issue > s.settle) throw E.num;
  // (Each over the year up to settlement's, for basis 1 as for the others.)
  const per = (x, y) => daysBetween(x, y, s.basis, ctx) / s.yearB;
  return { dim: per(issue, s.mat), a: per(issue, s.settle), dsm: per(s.settle, s.mat) };
};
def('PRICEMAT', (a, ctx) => {
  const t = atMaturity(a, ctx); const rate = num(a[3], ctx); const yld = num(a[4], ctx);
  if (rate < 0 || yld < 0) throw E.num;
  return (100 + t.dim * rate * 100) / (1 + t.dsm * yld) - t.a * rate * 100;
});
def('YIELDMAT', (a, ctx) => {
  const t = atMaturity(a, ctx); const rate = num(a[3], ctx); const pr = positive(a[4], ctx);
  if (rate < 0) throw E.num;
  const paid = pr / 100 + t.a * rate;
  return ((1 + t.dim * rate - paid) / paid) / t.dsm;
});
