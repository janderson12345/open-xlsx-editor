// AMORLINC and AMORDEGRC, the French accounting system's depreciation: linear, and declining with a coefficient by the
// asset's life, each with the first period prorated from the purchase to the first period's end. Microsoft's page
// gives other coefficients and errors than Excel's; these are Excel's, measured with cell values in a hidden Excel
// (about 560 cases) and the formula probe. Registered into functions.js's table; evaluate.js imports this file.
import { E } from './values.js';
import { FUNCTIONS, def, given, num, one, roundDec } from './functions.js';
import { dateParts } from '../ooxml/format.js';

const leap = (y) => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;

/** The first period's part of a year. 30/360 (bases 0 and 4) as YEARFRAC has it; actual days (bases 1 and 3) with a
 *  February 29 counted as the 28th, over 365, or for basis 1 over the purchase year's days. */
function firstPart(bought, first, basis, ctx) {
  if (basis === 0 || basis === 4) return FUNCTIONS.YEARFRAC.fn([bought, first, basis], ctx);
  const feb29 = (d) => { const p = dateParts(d, ctx.date1904); return p.m === 2 && p.d === 29 ? d - 1 : d; };
  const days = feb29(first) - feb29(bought);
  return days / (basis === 1 && leap(dateParts(bought, ctx.date1904).y) ? 366 : 365);
}

/** The arguments: cost, date_purchased, first_period, salvage, period, rate, [basis]. */
function args(a, ctx) {
  const cost = num(a[0], ctx);
  const bought = Math.trunc(num(a[1], ctx));
  const first = Math.trunc(num(a[2], ctx));
  const salvage = num(a[3], ctx);
  const period = Math.trunc(num(a[4], ctx));
  const rate = num(a[5], ctx);
  let basis = 0;
  if (given(a[6])) { if (typeof one(a[6], ctx) === 'boolean') throw E.value; basis = Math.trunc(num(a[6], ctx)); }
  if (cost < 0 || salvage < 0 || salvage > cost || period < 0 || rate <= 0 || bought > first || ![0, 1, 3, 4].includes(basis)) throw E.num;
  // (No days in the first period, as when it ends on the purchase date: a whole year.)
  const part = firstPart(bought, first, basis, ctx);
  return { cost, salvage, period, rate, part: part || 1, whole: !part };
}

def('AMORLINC', (a, ctx) => {
  const { cost, salvage, period, rate, part } = args(a, ctx);
  const year = cost * rate;
  const first = part * rate * cost;
  const full = Math.trunc((cost - salvage - first) / year);
  if (period === 0) return first;
  if (period <= full) return year;
  if (period === full + 1) return cost - salvage - year * full - first;
  return 0;
});

def('AMORDEGRC', (a, ctx) => {
  const { cost: c0, salvage, period, rate: r0, part, whole } = args(a, ctx);
  const life = 1 / r0;
  // (Excel's coefficients by the life: up to 2 years #NUM!, to 4 years 1.5, to 6 years 2, past that 2.5.)
  if (life <= 2) throw E.num;
  const rate = r0 * (life <= 4 ? 1.5 : life <= 6 ? 2 : 2.5);
  // (The periods of the asset's life, the first one with them; one fewer where the first period has no days. The last
  // two take half of what is left each, and none is taken once the value is below salvage. Each period's depreciation
  // is rounded as given; the value it leaves isn't, but for the first period's.)
  const last = Math.ceil(life) - (whole ? 2 : 1);
  // (Bought on the first period's end with a life of 3 years or less: no halves, a declining depreciation up to the
  // life's periods. Seen on two such assets only.)
  const short = last < 2;
  let n = roundDec(part * rate * c0, 0);
  let cost = c0 - n;
  let half = 0;
  for (let k = 1; k <= period; k++) {
    if (k > (short ? Math.ceil(life) : last) || cost < salvage) return 0;
    if (k === last - 1 && !short) half = cost * 0.5;
    n = k >= last - 1 && !short ? half : rate * cost;
    cost -= n;
  }
  return roundDec(n, 0);
});
