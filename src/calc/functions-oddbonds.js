// The securities with an odd first or last period: ODDFPRICE, ODDFYIELD, ODDLPRICE and ODDLYIELD, by Microsoft's
// formulas: a short or long odd first period (one or more quasi-coupon periods back from the first coupon), an odd last
// period summed over its quasi-coupon periods from the last interest date. Measured against Excel with cell values in
// a hidden Excel and the formula probe. Registered into functions.js's table; evaluate.js imports this file.
import { E, dateSerial } from './values.js';
import { dateParts } from '../ooxml/format.js';
import { def, num } from './functions.js';
import { daysBetween, addMonths, coupons, couponDays, day, basisOf, freqOf } from './functions-bonds.js';

/** The quasi-coupon periods Excel counts from settlement to the first coupon: stepping on from settlement (moved to
 *  its month's end when the first coupon is at one, which counts one) while short of the first coupon. So settlement
 *  on February 10, 2020 and a first coupon on June 30 make 1 half-year, not 0; November 11, 2008 to March 1, 2009
 *  makes 1 quarter. */
function periodsTo(settle, first, step, ctx) {
  const f = dateParts(first, ctx.date1904);
  const s = dateParts(settle, ctx.date1904);
  const days = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
  let eom = f.d === days(f.y, f.m);
  // (A first coupon on the 29th or 30th of a longer month, not February: by settlement's own month end.)
  if (!eom && f.m !== 2 && f.d > 28 && f.d < days(f.y, f.m)) eom = s.d === days(s.y, s.m);
  const start = eom ? dateSerial(s.y, s.m, days(s.y, s.m), ctx.date1904) : settle;
  let n = settle < start ? 1 : 0;
  for (let k = 1; addMonths(start, k * step, ctx, !eom) < first; k++) n++;
  return n;
}

/** The price of a bond with an odd first period, per 100 of face value. */
function oddFirstPrice(o, yld, ctx) {
  const { settle, mat, issue, first, rate, red, freq, basis } = o;
  const step = 12 / freq;
  const E0 = couponDays(settle, first, freq, basis, ctx).E;
  const N = coupons(first, mat, freq, ctx).n;
  const c = (100 * rate) / freq;
  const v = 1 + yld / freq;
  const dfc = daysBetween(issue, first, basis, ctx);
  if (dfc < E0) {
    // (A short first period.)
    const dsc = daysBetween(settle, first, basis, ctx);
    const A = daysBetween(issue, settle, basis, ctx);
    // (N the coupons from the first on, the first among them.)
    const n1 = N + 1;
    const t = dsc / E0;
    let p = red / v ** (n1 - 1 + t) + (c * dfc) / E0 / v ** t;
    for (let k = 2; k <= n1; k++) p += c / v ** (k - 1 + t);
    return p - (c * A) / E0;
  }
  // (A long first period: its quasi-coupon periods, back from the first coupon by whole months on the same day, not
  // kept at the month's end: June 30 to December 30, as Excel steps them.)
  // (Each step from the one before: December 31, September 30, June 30, March 30.)
  const dates = [first];
  const quasi = (k) => { while (dates.length <= k) dates.push(addMonths(dates[dates.length - 1], -step, ctx, true)); return dates[k]; };
  let nc = 1;
  while (quasi(nc) > issue) nc++;
  let dcnl = 0;
  let anl = 0;
  for (let i = nc; i >= 1; i--) {
    const late = quasi(nc - i);
    const early = quasi(nc - i + 1);
    const nl = basis === 1 ? daysBetween(early, late, basis, ctx) : E0;
    const dci = i > 1 ? nl : daysBetween(issue, late, basis, ctx);
    const from = issue > early ? issue : early;
    const to = settle < late ? settle : late;
    const a = from < to ? daysBetween(from, to, basis, ctx) : 0;
    dcnl += dci / nl;
    anl += a / nl;
  }
  // (The days to the next coupon as COUPDAYSNC's and COUPDAYBS's, by the coupon dates kept at the month's end.)
  const around = coupons(settle, first, freq, ctx);
  const dsc = basis === 2 || basis === 3 ? around.ncd - settle : E0 - daysBetween(around.pcd, settle, basis, ctx);
  const nq = periodsTo(settle, first, step, ctx);
  const t = dsc / E0;
  let p = red / v ** (t + nq + N) + (c * dcnl) / v ** (nq + t);
  for (let k = 1; k <= N; k++) p += c / v ** (k + nq + t);
  return p - c * anl;
}

/** ODDFPRICE's and ODDFYIELD's arguments: settlement, maturity, issue, first_coupon, rate, (yld or pr), redemption,
 *  frequency, [basis]. Issue before settlement before the first coupon before maturity, else #NUM!. */
function oddFirstArgs(a, ctx) {
  const o = {
    settle: day(a[0], ctx), mat: day(a[1], ctx), issue: day(a[2], ctx), first: day(a[3], ctx), rate: num(a[4], ctx),
    red: num(a[6], ctx), freq: freqOf(a[7], ctx), basis: basisOf(a, 8, ctx),
  };
  if (!(o.issue < o.settle && o.settle < o.first && o.first < o.mat) || o.rate < 0 || o.red <= 0) throw E.num;
  return o;
}

/** The yield whose price is `price(y)` = pr, by Newton's method from `y0`. */
function solve(price, pr, y0) {
  let y = y0;
  for (let i = 0; i < 100; i++) {
    const f = price(y) - pr;
    const h = 1e-7 * Math.max(1, Math.abs(y));
    const d = (price(y + h) - price(y - h)) / (2 * h);
    const next = y - f / d;
    if (!Number.isFinite(next)) throw E.num;
    if (Math.abs(next - y) < 1e-14 * Math.max(1, Math.abs(y))) return next;
    y = next;
  }
  return y;
}

def('ODDFPRICE', (a, ctx) => {
  const o = oddFirstArgs(a, ctx);
  const yld = num(a[5], ctx);
  if (yld < 0) throw E.num;
  return oddFirstPrice(o, yld, ctx);
});
def('ODDFYIELD', (a, ctx) => {
  const o = oddFirstArgs(a, ctx);
  const pr = num(a[5], ctx);
  if (pr <= 0) throw E.num;
  return solve((y) => oddFirstPrice(o, y, ctx), pr, o.rate || 0.05);
});

/** An odd last period's sums over its quasi-coupon periods from the last interest date: { dcnl, anl, dscnl }. */
function oddLastSums(settle, mat, last, freq, basis, ctx) {
  const step = 12 / freq;
  // (Quasi-coupon dates on from the last interest date, each from the one before on the same day, not kept at the
  // month's end: January 31, April 30, July 30.)
  let nc = 0;
  for (let d = last; d < mat; d = addMonths(d, step, ctx, true)) nc++;
  let early = last;
  let dcnl = 0;
  let anl = 0;
  let dscnl = 0;
  for (let i = 1; i <= nc; i++) {
    const late = addMonths(early, step, ctx, true);
    // (A quasi-period's normal length: 360 over the frequency for 30/360, else its actual days, bases 2 and 3 too.)
    const nl = basis === 0 || basis === 4 ? 360 / freq : late - early;
    const dci = i < nc ? nl : daysBetween(early, mat, basis, ctx);
    const A = late < settle ? dci : early < settle ? daysBetween(early, settle, basis, ctx) : 0;
    const from = settle > early ? settle : early;
    const to = mat < late ? mat : late;
    const dsc = from < to ? daysBetween(from, to, basis, ctx) : 0;
    dcnl += dci / nl;
    anl += A / nl;
    dscnl += dsc / nl;
    early = late;
  }
  return { dcnl, anl, dscnl };
}
/** ODDLPRICE's and ODDLYIELD's arguments: settlement, maturity, last_interest, rate, (yld or pr), redemption,
 *  frequency, [basis]. The last interest date before settlement before maturity, else #NUM!. */
function oddLastArgs(a, ctx) {
  const o = {
    settle: day(a[0], ctx), mat: day(a[1], ctx), last: day(a[2], ctx), rate: num(a[3], ctx), red: num(a[5], ctx),
    freq: freqOf(a[6], ctx), basis: basisOf(a, 7, ctx),
  };
  if (!(o.last < o.settle && o.settle < o.mat) || o.rate < 0 || o.red <= 0) throw E.num;
  return { ...o, ...oddLastSums(o.settle, o.mat, o.last, o.freq, o.basis, ctx) };
}
def('ODDLPRICE', (a, ctx) => {
  const o = oddLastArgs(a, ctx);
  const yld = num(a[4], ctx);
  if (yld < 0) throw E.num;
  const x = (100 * o.rate) / o.freq;
  return (o.dcnl * x + o.red) / (1 + (o.dscnl * yld) / o.freq) - o.anl * x;
});
def('ODDLYIELD', (a, ctx) => {
  const o = oddLastArgs(a, ctx);
  const pr = num(a[4], ctx);
  if (pr <= 0) throw E.num;
  const x = (100 * o.rate) / o.freq;
  const paid = o.anl * x + pr;
  return ((o.dcnl * x + o.red - paid) / paid) * (o.freq / o.dscnl);
});
