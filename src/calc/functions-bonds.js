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

/** ACCRINT's coupon periods, in what Excel was seen to do (measured on 600 cases over the five bases, both
 *  calc_methods, then 387 of 400 more it wasn't worked out from; docs/CALC.md): par x rate / frequency x the periods
 *  from issue to settlement.
 *  - The coupon dates step from first_interest back past issue, each from the one before (from August 30: February
 *    28, then August 28), at the month's end when first_interest is at one; and on from the earliest the same way, so
 *    the period issue falls in can end before first_interest (August 28, not August 29).
 *  - A period's length L: 360 or 365 over the frequency, or with basis 1 the days of the period that ends at
 *    first_interest; issue's own period, with basis 1, its own days.
 *  - Issue in the period ending at first_interest: the days from issue over L. Otherwise up to first_interest the
 *    part of issue's period counts its days over its length, each later whole period before the last 1, and the rest
 *    from issue to settlement, less those periods' days, over L; so a long first period accrues at L's rate all
 *    through it (Excel's answer is linear in settlement there, with an offset).
 *  - Past first_interest: the period ending at it counts 1, each later whole one (ended before settlement) 1, and a
 *    part its days over L (August 31 to February 28: 178 of 180 in 30/360).
 *  - calc_method FALSE: issue's part as above, and the days from the start of the period ending at first_interest
 *    (stepped back from it) to settlement over L, whatever Microsoft's page says. */
function accruedPeriods(issue, first, settle, freq, basis, byPeriod, ctx) {
  const step = 12 / freq;
  const fp = dateParts(first, ctx.date1904);
  const keepDay = fp.d !== monthDays(fp.y, fp.m);
  // (A coupon date the walk steps to before 1900-01-01 or past 9999-12-31 is #NUM!, as in Excel: measured on 2000
  // cases near both ends, every one as Excel's. So ACCRINT(1,731,...) is #NUM!, its walk stepping back past 1900.)
  const inRange = (d) => { if (d < 1 || d > LAST_DATE) throw E.num; return d; };
  const back = (d) => inRange(addMonths(d, -step, ctx, keepDay));
  const on = (d) => inRange(addMonths(d, step, ctx, keepDay));
  const days = (s, e) => daysBetween(s, e, basis, ctx);
  const L = basis === 1 ? first - back(first) : basis === 3 ? 365 / freq : 360 / freq;
  if (first <= issue) {
    // (Issue on or after first_interest: the periods from it on, issue's a part.)
    // (Here too Excel steps back from first_interest once, whatever the basis: #NUM! when that's before 1900.)
    if (!byPeriod) { back(first); return days(issue, settle) / L; }
    let p0 = first;
    while (on(p0) <= issue) p0 = on(p0);
    let total = 0;
    for (; p0 < settle; p0 = on(p0)) {
      const from = Math.max(p0, issue);
      total += from === p0 && on(p0) < settle ? 1 : days(from, Math.min(on(p0), settle)) / L;
    }
    return total;
  }
  if (!byPeriod) {
    const B = [first];
    while (B[B.length - 1] > issue) B.push(back(B[B.length - 1]));
    const n = B.length;
    if (n === 2) return days(issue, settle) / L;
    // (Issue on a coupon date: its period counts nothing, ACCRINT(DATE(1900,1,29),DATE(1901,1,29),DATE(1900,5,30),
    // 0.1,1000,4,1,FALSE) a whole period less than issue's part and the rest would make it: measured.)
    if (B[n - 1] === issue) return days(B[1], settle) / L;
    return days(issue, B[n - 2]) / (basis === 1 ? B[n - 2] - B[n - 1] : L) + days(B[1], settle) / L;
  }
  let start = first;
  while (start > issue) start = back(start);
  const P = [start];
  while (P[P.length - 1] < first) P.push(on(P[P.length - 1]));
  const last = P.length - 1;
  // (Issue on a coupon date: its period counts 1, not its days over L: measured.)
  const whole = start === issue;
  if (last === 1) {
    if (settle <= first) return days(issue, settle) / L;
    let total = whole ? 1 : days(issue, P[1]) / L;
    for (let p0 = P[1]; p0 < settle; p0 = on(p0)) total += on(p0) < settle ? 1 : days(p0, settle) / L;
    return total;
  }
  const own = basis === 1 ? P[1] - P[0] : L;
  const dIssue = days(issue, P[1]);
  if (settle <= first && whole) {
    // (1 for each period ended before settlement, then the days from the last coupon date over L.)
    let k = 0;
    while (P[k + 1] < settle) k++;
    return k + days(P[k], settle) / L;
  }
  if (settle <= first) {
    let middle = 0;
    for (let i = 1; i < last - 1; i++) middle += days(P[i], P[i + 1]);
    return dIssue / own + (last - 2) + (days(issue, settle) - dIssue - middle) / L;
  }
  let total = (whole ? 1 : dIssue / own) + (last - 1);
  for (let p0 = P[last]; p0 < settle; p0 = on(p0)) total += on(p0) < settle ? 1 : days(p0, settle) / L;
  return total;
}
/** Excel's last date, 9999-12-31. */
const LAST_DATE = 2958465;
def('ACCRINT', (a, ctx) => {
  const issue = day(a[0], ctx); const first = day(a[1], ctx); const settle = day(a[2], ctx);
  const rate = num(a[3], ctx); const par = given(a[4]) ? num(a[4], ctx) : 1000;
  const freq = freqOf(a[5], ctx); const basis = basisOf(a, 6, ctx);
  const byPeriod = !given(a[7]) || !!num(a[7], ctx);
  if (rate <= 0 || par <= 0 || issue >= settle) throw E.num;
  // (A date past 9999-12-31, Excel's last (serial 2958465): #NUM!, as Excel gives for each of the three (measured over
  // COM); and no coupon-by-coupon walk across millions of days for a crafted one.)
  if (issue > LAST_DATE || first > LAST_DATE || settle > LAST_DATE) throw E.num;
  return (par * rate * accruedPeriods(issue, first, settle, freq, basis, byPeriod, ctx)) / freq;
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

// (For the odd-period bonds, functions-oddbonds.js.)
export { daysBetween, addMonths, coupons, couponDays, day, basisOf, freqOf };
