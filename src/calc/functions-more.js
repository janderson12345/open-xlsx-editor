// More of Excel's worksheet functions, as Excel works them out (each checked against Excel's results in
// test/formulas-excel.json, docs/CALC.md): SUBTOTAL, ranks, percentiles and quartiles, the mode; working days,
// DATEDIF, week numbers and YEARFRAC; loan and investment functions; LOOKUP, XMATCH and ADDRESS; text before and after
// a delimiter, NUMBERVALUE, FIXED and DOLLAR; trigonometry, GCD and LCM, CEILING.MATH and FLOOR.MATH; TYPE and
// ERROR.TYPE. Registered into functions.js's table.
import { Area, Arr, Union, Ref3D, MISSING, XlError, E, toNum, compare, finite } from './values.js';
import { FUNCTIONS, def, lifted, given, one, num, text, opt, values, numbers, vector, roundDec } from './functions.js';
import { serialOf } from './functions-text.js';
import { position } from './functions-lookup.js';
import { dateParts } from '../ooxml/format.js';
import { colName } from '../ooxml/refs.js';
import { sin, cos, tan, TRIG_LIMIT } from './trig.js';

const sorted = (args, ctx) => numbers(args, ctx).sort((a, b) => a - b);

// ---- SUBTOTAL: the function by number (1 to 11, or 101 to 111 leaving hidden rows out), over cells that aren't
// themselves SUBTOTAL or AGGREGATE.

const SUBTOTALS = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV', 8: 'STDEVP', 9: 'SUM', 10: 'VAR', 11: 'VARP' };
def('SUBTOTAL', (a, ctx) => {
  const n = Math.trunc(num(a[0], ctx));
  const name = SUBTOTALS[n > 100 ? n - 100 : n];
  if (!name) throw E.value;
  const kept = [];
  for (const ref of a.slice(1)) {
    if (ref instanceof Ref3D) throw E.value;
    const areas = ref instanceof Area ? [ref] : ref instanceof Union ? ref.areas : null;
    if (!areas) throw E.value;
    for (const x of areas) {
      for (let r = x.r0; r <= x.r1; r++) {
        if (n > 100 && ctx.rowHidden?.(x.si, r)) continue;
        for (let c = x.c0; c <= x.c1; c++) {
          const f = ctx.formulaOf?.(x.si, r, c);
          if (f && /^\s*(_xlfn\.)?(SUBTOTAL|AGGREGATE)\s*\(/i.test(f)) continue;
          kept.push(new Area(x.si, r, c, r, c));
        }
      }
    }
  }
  return FUNCTIONS[name].fn(kept, ctx);
});

// ---- Ranks, percentiles, quartiles and the mode

const rank = (avg) => (a, ctx) => {
  const x = num(a[0], ctx);
  const list = numbers([a[1]], ctx);
  const up = given(a[2]) && num(a[2], ctx) !== 0;
  if (!list.some((v) => compare(v, x) === 0)) throw E.na;
  const before = list.filter((v) => (up ? v < x : v > x) && compare(v, x) !== 0).length;
  const ties = list.filter((v) => compare(v, x) === 0).length;
  return avg ? before + (ties + 1) / 2 : before + 1;
};
def('RANK RANK.EQ', rank(false));
def('RANK.AVG', rank(true));
/** The k-th percentile, inclusive (Excel's PERCENTILE) or exclusive. */
function percentile(list, k, exclusive) {
  if (!list.length) throw E.num;
  const n = list.length;
  if (exclusive) {
    const h = k * (n + 1) - 1;
    if (k <= 0 || k >= 1 || h < 0 || h > n - 1) throw E.num;
    const i = Math.floor(h);
    return i + 1 < n ? list[i] + (h - i) * (list[i + 1] - list[i]) : list[i];
  }
  if (k < 0 || k > 1) throw E.num;
  const h = k * (n - 1);
  const i = Math.floor(h);
  return i + 1 < n ? list[i] + (h - i) * (list[i + 1] - list[i]) : list[i];
}
def('PERCENTILE PERCENTILE.INC', (a, ctx) => percentile(sorted([a[0]], ctx), num(a[1], ctx), false));
def('PERCENTILE.EXC', (a, ctx) => percentile(sorted([a[0]], ctx), num(a[1], ctx), true));
const quartile = (exclusive) => (a, ctx) => {
  const q = Math.trunc(num(a[1], ctx));
  if (exclusive ? q < 1 || q > 3 : q < 0 || q > 4) throw E.num;
  return percentile(sorted([a[0]], ctx), q / 4, exclusive);
};
def('QUARTILE QUARTILE.INC', quartile(false));
def('QUARTILE.EXC', quartile(true));
def('MODE MODE.SNGL', (a, ctx) => {
  // (The value that comes most often, the first of them as they are listed; none repeated: #N/A.)
  const list = numbers(a, ctx);
  const counts = new Map();
  let best = null;
  let most = 1;
  for (const v of list) {
    const k = (counts.get(v) || 0) + 1;
    counts.set(v, k);
    if (k > most || (k === most && best !== null && list.indexOf(v) < list.indexOf(best))) { most = k; best = v; }
  }
  if (best === null) throw E.na;
  return best;
});
/** The values of the arguments as AVERAGEA, MAXA and MINA see them: text as 0, TRUE as 1. */
function valuesA(args, ctx) {
  const out = [];
  for (const { v, direct } of values(args, ctx)) {
    if (v instanceof XlError) throw v;
    if (v === null || v === MISSING) continue;
    if (typeof v === 'number') out.push(v);
    else if (typeof v === 'boolean') out.push(v ? 1 : 0);
    else out.push(direct ? toNum(v, ctx.date1904) : 0);
  }
  return out;
}
def('AVERAGEA', (a, ctx) => { const n = valuesA(a, ctx); if (!n.length) throw E.div0; return n.reduce((s, x) => s + x, 0) / n.length; });
def('MAXA', (a, ctx) => { const n = valuesA(a, ctx); return n.length ? Math.max(...n) : 0; });
def('MINA', (a, ctx) => { const n = valuesA(a, ctx); return n.length ? Math.min(...n) : 0; });

// ---- Working days and dates

/** Holidays as whole serial dates. */
const holidays = (v, ctx) => (given(v) ? new Set(numbers([v], ctx).map(Math.trunc)) : new Set());
const weekend = (serial) => { const d = (Math.floor(serial) + 6) % 7; return d === 0 || d === 6; }; // (serial 1 is a Sunday: 1 + 6 is 0)
def('NETWORKDAYS', (a, ctx) => {
  let s = Math.trunc(serialOf(a[0], ctx));
  let e = Math.trunc(serialOf(a[1], ctx));
  const sign = s <= e ? 1 : -1;
  if (sign < 0) [s, e] = [e, s];
  const off = holidays(a[2], ctx);
  let n = 0;
  for (let d = s; d <= e; d++) if (!weekend(d) && !off.has(d)) n++;
  return sign * n;
});
def('WORKDAY', (a, ctx) => {
  let d = Math.trunc(serialOf(a[0], ctx));
  let k = Math.trunc(num(a[1], ctx));
  const off = holidays(a[2], ctx);
  const step = k < 0 ? -1 : 1;
  while (k !== 0) { d += step; if (!weekend(d) && !off.has(d)) k -= step; }
  return d;
});
def('DATEDIF', (a, ctx) => {
  const s = Math.trunc(serialOf(a[0], ctx));
  const e = Math.trunc(serialOf(a[1], ctx));
  if (s > e) throw E.num;
  const p = dateParts(s, ctx.date1904);
  const q = dateParts(e, ctx.date1904);
  const months = (q.y - p.y) * 12 + (q.m - p.m) - (q.d < p.d ? 1 : 0);
  switch (text(a[2], ctx).toUpperCase()) {
    case 'D': return e - s;
    case 'M': return months;
    case 'Y': return Math.floor(months / 12);
    case 'YM': return months % 12;
    case 'MD': { const back = new Date(Date.UTC(q.y, q.m - 1 - (q.d < p.d ? 1 : 0), p.d)); return Math.round((Date.UTC(q.y, q.m - 1, q.d) - back.getTime()) / 86400000); }
    case 'YD': {
      let y = q.y;
      if (q.m < p.m || (q.m === p.m && q.d < p.d)) y -= 1;
      return Math.round((Date.UTC(q.y, q.m - 1, q.d) - Date.UTC(y, p.m - 1, p.d)) / 86400000);
    }
    default: throw E.num;
  }
});
/** The weekday of a serial date, Sunday 0. */
const dayOf = (serial) => (Math.floor(serial) + 6) % 7;
def('WEEKNUM', (a, ctx) => {
  const s = Math.trunc(serialOf(a[0], ctx));
  const type = given(a[1]) ? Math.trunc(num(a[1], ctx)) : 1;
  if (type === 21) return isoWeek(s, ctx);
  const first = { 1: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6, 17: 0 }[type];
  if (first === undefined) throw E.num;
  const p = dateParts(s, ctx.date1904);
  const jan1 = s - (Math.round((Date.UTC(p.y, p.m - 1, p.d) - Date.UTC(p.y, 0, 1)) / 86400000));
  const lead = (dayOf(jan1) - first + 7) % 7;
  return Math.floor((s - jan1 + lead) / 7) + 1;
});
function isoWeek(s, ctx) {
  // (The week of the year's first Thursday is week 1; weeks start on Monday.)
  const thursday = s - ((dayOf(s) + 6) % 7) + 3;
  const p = dateParts(thursday, ctx.date1904);
  const jan1 = thursday - Math.round((Date.UTC(p.y, p.m - 1, p.d) - Date.UTC(p.y, 0, 1)) / 86400000);
  return Math.floor((thursday - jan1) / 7) + 1;
}
def('ISOWEEKNUM', (a, ctx) => isoWeek(Math.trunc(serialOf(a[0], ctx)), ctx));
def('YEARFRAC', (a, ctx) => {
  let s = Math.trunc(serialOf(a[0], ctx));
  let e = Math.trunc(serialOf(a[1], ctx));
  const basis = given(a[2]) ? Math.trunc(num(a[2], ctx)) : 0;
  if (basis < 0 || basis > 4) throw E.num;
  if (s > e) [s, e] = [e, s];
  const p = dateParts(s, ctx.date1904);
  const q = dateParts(e, ctx.date1904);
  const leap = (y) => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
  const lastOfFeb = (d) => d.m === 2 && d.d === (leap(d.y) ? 29 : 28);
  switch (basis) {
    case 0: {
      // (US 30/360, with its February rules.)
      let d1 = p.d;
      let d2 = q.d;
      if (lastOfFeb(p) && lastOfFeb(q)) d2 = 30;
      if (lastOfFeb(p)) d1 = 30;
      // (The 31st as the end becomes the 30th only after a start on the 30th or 31st itself, not after the end of February
      // made the 30th: February 28, 2021 to March 31 is 31 days, February 29, 2020 to December 31 301.)
      if (d2 === 31 && p.d >= 30) d2 = 30;
      if (d1 === 31) d1 = 30;
      return ((q.y - p.y) * 360 + (q.m - p.m) * 30 + (d2 - d1)) / 360;
    }
    case 1: {
      // (Actual/actual: within a year (or a year and a bit up to the same day), the days over that year's length; over
      // longer, over the average length of the years spanned.)
      const days = e - s;
      const oneYear = q.y === p.y || (q.y === p.y + 1 && (q.m < p.m || (q.m === p.m && q.d <= p.d)));
      if (oneYear) {
        const spansLeap = q.y === p.y ? leap(p.y) : (leap(p.y) && (p.m < 3)) || (leap(q.y) && (q.m > 2 || (q.m === 2 && q.d === 29)));
        return days / (spansLeap ? 366 : 365);
      }
      let total = 0;
      for (let y = p.y; y <= q.y; y++) total += leap(y) ? 366 : 365;
      return days / (total / (q.y - p.y + 1));
    }
    case 2: return (e - s) / 360;
    case 3: return (e - s) / 365;
    default: {
      // (European 30/360.)
      const d1 = Math.min(p.d, 30);
      const d2 = Math.min(q.d, 30);
      return ((q.y - p.y) * 360 + (q.m - p.m) * 30 + (d2 - d1)) / 360;
    }
  }
});

// ---- Loans and investments (Excel's signs: money paid out negative). Excel raises 1 + rate to a whole number of
// periods by repeated squaring in PV and FV, but through logarithms in PMT: each as Excel was seen to, to its last
// digit (PMT(0.05/12,360,200000) is -1073.6432460242781, FV(0.06/12,120,-100) 16387.934680645605).

/** (1 + r)^n by repeated squaring (a whole n), else as JavaScript raises it. */
export function grow(r, n) {
  if (!Number.isInteger(n)) return (1 + r) ** n;
  let b = 1 + r;
  let e = Math.abs(n);
  let g = 1;
  while (e) { if (e & 1) g *= b; b *= b; e = Math.floor(e / 2); }
  return n < 0 ? 1 / g : g;
}
const pv = (rate, n, pmt, fv, type) => (rate === 0 ? -(fv + pmt * n) : -(fv + pmt * (1 + rate * type) * (grow(rate, n) - 1) / rate) / grow(rate, n));
export const fv = (rate, n, pmt, p, type) => (rate === 0 ? -(p + pmt * n) : -(p * grow(rate, n) + pmt * (1 + rate * type) * (grow(rate, n) - 1) / rate));
export const pmt = (rate, n, p, f, type) => {
  if (n === 0) throw E.num;
  if (rate === 0) return -(p + f) / n;
  const g = Math.exp(n * Math.log1p(rate));
  return (-(p * g + f) * rate / (g - 1)) / (1 + rate * type);
};
const loanArgs = (a, ctx, from) => [opt(a, from, ctx, 0), given(a[from + 1]) && num(a[from + 1], ctx) !== 0 ? 1 : 0];
def('PMT', (a, ctx) => { const [f, t] = loanArgs(a, ctx, 3); return finite(pmt(num(a[0], ctx), num(a[1], ctx), num(a[2], ctx), f, t)); });
def('PV', (a, ctx) => { const [f, t] = loanArgs(a, ctx, 3); return finite(pv(num(a[0], ctx), num(a[1], ctx), num(a[2], ctx), f, t)); });
def('FV', (a, ctx) => { const [p, t] = loanArgs(a, ctx, 3); return finite(fv(num(a[0], ctx), num(a[1], ctx), num(a[2], ctx), p, t)); });
def('NPER', (a, ctx) => {
  const rate = num(a[0], ctx);
  const p = num(a[1], ctx);
  const v = num(a[2], ctx);
  const [f, t] = loanArgs(a, ctx, 3);
  if (rate === 0) { if (p === 0) throw E.div0; return -(v + f) / p; }
  const k = p * (1 + rate * t) / rate;
  const x = (k - f) / (k + v);
  if (x <= 0) throw E.num;
  return finite(Math.log(x) / Math.log(1 + rate));
});
// (RATE and IRR are searched for by Newton's method. Excel promises them to 0.0000001 and its last digits are its own
// search's: they are tested to 1e-9.)
def('RATE', (a, ctx) => {
  const n = num(a[0], ctx);
  const p = num(a[1], ctx);
  const v = num(a[2], ctx);
  const [f, t] = loanArgs(a, ctx, 3);
  let r = given(a[5]) ? num(a[5], ctx) : 0.1;
  for (let i = 0; i < 50; i++) {
    if (r === 0) r = 1e-10;
    const g = grow(r, n);
    const y = v * g + p * (1 + r * t) * (g - 1) / r + f;
    const dg = n * grow(r, n - 1);
    const dy = v * dg + p * t * (g - 1) / r + p * (1 + r * t) * (dg * r - (g - 1)) / (r * r);
    if (!Number.isFinite(dy) || dy === 0) break;
    const next = r - y / dy;
    if (!Number.isFinite(next) || next <= -1) break;
    if (Math.abs(next - r) < 1e-10) return next;
    r = next;
  }
  throw E.num;
});
def('NPV', (a, ctx) => {
  const rate = num(a[0], ctx);
  if (rate === -1) throw E.div0;
  return numbers(a.slice(1), ctx).reduce((s, x, i) => s + x / (1 + rate) ** (i + 1), 0);
});
def('IRR', (a, ctx) => {
  const flows = numbers([a[0]], ctx);
  if (!flows.some((x) => x > 0) || !flows.some((x) => x < 0)) throw E.num;
  // (Searched for on the discount factor 1 / (1 + rate), on which the flows are a polynomial: IRR({-100,10}) is -0.9.)
  let d = 1 / (1 + (given(a[1]) ? num(a[1], ctx) : 0.1));
  for (let i = 0; i < 50; i++) {
    let y = 0;
    let dy = 0;
    for (let k = flows.length - 1; k >= 0; k--) { dy = dy * d + y; y = y * d + flows[k]; }
    if (dy === 0) break;
    const next = d - y / dy;
    if (!Number.isFinite(next) || next <= 0) break;
    if (Math.abs(next - d) < 1e-12) return 1 / next - 1;
    d = next;
  }
  throw E.num;
});

// ---- LOOKUP, XMATCH, ADDRESS

def('LOOKUP', (a, ctx) => {
  const x = one(a[0], ctx);
  const look = vector(a[1], ctx);
  if (!look) throw E.na;
  const i = position(look, x === null ? 0 : x, 1, false);
  if (i < 0) throw E.na;
  const out = given(a[2]) ? vector(a[2], ctx) : look;
  if (!out) throw E.na;
  const v = out.at(i);
  return v === null ? 0 : v;
});
def('XMATCH', (a, ctx) => {
  const x = one(a[0], ctx);
  const list = vector(a[1], ctx);
  if (!list) throw E.value;
  const mode = opt(a, 2, ctx, 0);
  const search = opt(a, 3, ctx, 1);
  if (![0, -1, 1, 2].includes(mode) || ![1, -1, 2, -2].includes(search)) throw E.value;
  const order = search < 0 ? [...Array(list.length).keys()].reverse() : [...Array(list.length).keys()];
  const at = (i) => list.at(i);
  if (mode === 0 || mode === 2) {
    const eq = (v) => (mode === 2 && typeof x === 'string' && typeof v === 'string' ? new RegExp(`^${x.replace(/[.+^${}()|[\]\\/]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?')}$`, 'i').test(v) : v !== null && compare(v, x) === 0 && typeof v === typeof x);
    for (const i of order) if (eq(at(i))) return i + 1;
    throw E.na;
  }
  // (Next smaller (-1) or next larger (1): the closest on that side, the first of equals in search order.)
  let best = -1;
  for (const i of order) {
    const v = at(i);
    if (v === null || typeof v !== typeof x) continue;
    const k = compare(v, x);
    if (k === 0) return i + 1;
    if ((mode === -1 && k < 0) || (mode === 1 && k > 0)) {
      if (best < 0 || (mode === -1 ? compare(v, at(best)) > 0 : compare(v, at(best)) < 0)) best = i;
    }
  }
  if (best < 0) throw E.na;
  return best + 1;
});
def('ADDRESS', (a, ctx) => {
  const r = Math.trunc(num(a[0], ctx));
  const c = Math.trunc(num(a[1], ctx));
  const abs = opt(a, 2, ctx, 1);
  const a1 = given(a[3]) ? one(a[3], ctx) : true;
  if (r < 1 || c < 1 || r > 1048576 || c > 16384 || abs < 1 || abs > 4) throw E.value;
  const useA1 = !(a1 === false || a1 === 0);
  let ref;
  if (useA1) ref = `${abs === 1 || abs === 3 ? '$' : ''}${colName(c - 1)}${abs === 1 || abs === 2 ? '$' : ''}${r}`;
  else ref = `R${abs === 1 || abs === 2 ? r : `[${r}]`}C${abs === 1 || abs === 3 ? c : `[${c}]`}`;
  if (!given(a[4])) return ref;
  const sheet = text(a[4], ctx);
  return `${/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`}!${ref}`;
});

// ---- Text

const around = (after) => (a, ctx) => {
  const s = text(a[0], ctx);
  const d = text(a[1], ctx);
  let n = given(a[2]) ? Math.trunc(num(a[2], ctx)) : 1;
  const nocase = given(a[3]) && num(a[3], ctx) === 1;
  if (n === 0 || Math.abs(n) > s.length + 1) throw E.value;
  if (d === '') return after ? (n > 0 ? s : '') : (n > 0 ? '' : s);
  const hay = nocase ? s.toLowerCase() : s;
  const needle = nocase ? d.toLowerCase() : d;
  let at = -1;
  if (n > 0) { let from = 0; while (n--) { at = hay.indexOf(needle, from); if (at < 0) break; from = at + needle.length; } }
  else { let from = hay.length; while (n++) { at = hay.lastIndexOf(needle, from - 1); if (at < 0) break; from = at; } }
  if (at < 0) { if (given(a[5])) return one(a[5], ctx); throw E.na; }
  return after ? s.slice(at + d.length) : s.slice(0, at);
};
def('TEXTBEFORE', around(false));
def('TEXTAFTER', around(true));
lifted('NUMBERVALUE', (a, ctx) => {
  const dec = given(a[1]) ? text(a[1], ctx)[0] : '.';
  const group = given(a[2]) ? text(a[2], ctx)[0] : ',';
  let s = text(a[0], ctx).replace(/\s+/g, '');
  if (s === '') return 0;
  let pct = 0;
  while (s.endsWith('%')) { s = s.slice(0, -1); pct++; }
  const at = s.indexOf(dec);
  if (at >= 0 && s.indexOf(dec, at + 1) >= 0) throw E.value;
  const whole = (at >= 0 ? s.slice(0, at) : s).split(group).join('');
  const frac = at >= 0 ? s.slice(at + 1) : '';
  if (frac.includes(group)) throw E.value;
  const n = Number(`${whole}${frac ? `.${frac}` : ''}`);
  if (!/^[+-]?\d*(\.\d*)?([eE][+-]?\d+)?$/.test(`${whole}${frac ? `.${frac}` : ''}`) || !Number.isFinite(n)) throw E.value;
  return n / 100 ** pct;
});
/** A number with `dec` decimals (negative: rounded to tens), with or without thousands separators. */
function fixed(x, dec, commas) {
  const r = roundDec(x, dec, 'half');
  const s = Math.abs(r).toFixed(Math.max(0, dec));
  const [i, f] = s.split('.');
  const grouped = commas ? i.replace(/\B(?=(\d{3})+(?!\d))/g, ',') : i;
  return `${r < 0 ? '-' : ''}${grouped}${f ? `.${f}` : ''}`;
}
lifted('FIXED', (a, ctx) => fixed(num(a[0], ctx), given(a[1]) ? Math.trunc(num(a[1], ctx)) : 2, !(given(a[2]) && toNum(one(a[2], ctx)) !== 0)));
lifted('DOLLAR', (a, ctx) => {
  const x = num(a[0], ctx);
  const s = fixed(Math.abs(x), given(a[1]) ? Math.trunc(num(a[1], ctx)) : 2, true);
  return x < 0 ? `($${s})` : `$${s}`;
});
lifted('UNICODE', (a, ctx) => { const s = text(a[0], ctx); if (!s) throw E.value; return s.codePointAt(0); });
lifted('UNICHAR', (a, ctx) => { const n = Math.trunc(num(a[0], ctx)); if (n < 1 || n > 0x10ffff) throw E.value; return String.fromCodePoint(n); });

// ---- Trigonometry, GCD and LCM, rounding to a multiple

const trig = (names, fn) => lifted(names, (a, ctx) => finite(fn(num(a[0], ctx))));
// (Excel's sine and cosine, trig.js: =SIN(PI()) is 1.22514845490862E-16, =SIN(134217728) #NUM!.)
const circular = (names, fn) => lifted(names, (a, ctx) => { const x = num(a[0], ctx); if (Math.abs(x) >= TRIG_LIMIT) throw E.num; return finite(fn(x)); });
circular('SIN', sin);
circular('COS', cos);
circular('TAN', tan);
lifted('ASIN', (a, ctx) => { const x = num(a[0], ctx); if (x < -1 || x > 1) throw E.num; return Math.asin(x); });
lifted('ACOS', (a, ctx) => { const x = num(a[0], ctx); if (x < -1 || x > 1) throw E.num; return Math.acos(x); });
trig('ATAN', Math.atan);
lifted('ATAN2', (a, ctx) => { const x = num(a[0], ctx); const y = num(a[1], ctx); if (x === 0 && y === 0) throw E.div0; return Math.atan2(y, x); });
trig('RADIANS', (x) => x * Math.PI / 180);
trig('DEGREES', (x) => x * 180 / Math.PI);
const ints = (a, ctx) => numbers(a, ctx).map((x) => { if (x < 0) throw E.num; return Math.trunc(x); });
const gcd2 = (x, y) => { while (y) [x, y] = [y, x % y]; return x; };
def('GCD', (a, ctx) => ints(a, ctx).reduce(gcd2, 0));
def('LCM', (a, ctx) => ints(a, ctx).reduce((m, x) => (m === 0 || x === 0 ? 0 : (m / gcd2(m, x)) * x), 1));
const toMultiple = (up) => (a, ctx) => {
  const x = num(a[0], ctx);
  const sig = Math.abs(given(a[1]) ? num(a[1], ctx) : 1);
  const away = given(a[2]) && num(a[2], ctx) !== 0;
  if (sig === 0) return 0;
  const q = x / sig;
  // (Negative numbers go toward zero, unless the mode says away from it.)
  const r = x < 0 && away ? (up ? Math.floor(q) : Math.ceil(q)) : (up ? Math.ceil(Number(q.toPrecision(15))) : Math.floor(Number(q.toPrecision(15))));
  return Number((r * sig).toPrecision(15));
};
def('CEILING.MATH', toMultiple(true));
def('FLOOR.MATH', toMultiple(false));
def('COMBIN', (a, ctx) => {
  const n = Math.trunc(num(a[0], ctx));
  const k = Math.trunc(num(a[1], ctx));
  if (n < 0 || k < 0 || k > n) throw E.num;
  let r = 1;
  for (let i = 1; i <= k; i++) r = (r * (n - k + i)) / i;
  return Math.round(r);
});

// ---- TYPE and ERROR.TYPE

def('TYPE', (a, ctx) => {
  const v = a[0];
  if (v instanceof Arr || (v instanceof Area && v.size > 1)) return 64;
  const x = v instanceof Area ? ctx.cellValue(v.si, v.r0, v.c0) : v;
  if (x instanceof XlError) return 16;
  if (typeof x === 'string') return 2;
  if (typeof x === 'boolean') return 4;
  return 1;
});
// (Excel 365's newer errors too: #SPILL! 9, #CALC! 14, measured =ERROR.TYPE(RANDARRAY(0)).)
const ERROR_TYPES = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CONNECT!', '#BLOCKED!', '#UNKNOWN!', '#FIELD!', '#CALC!'];
def('ERROR.TYPE', (a, ctx) => {
  const v = a[0] instanceof Area ? ctx.cellValue(a[0].si, a[0].r0, a[0].c0) : a[0];
  if (!(v instanceof XlError)) throw E.na;
  const i = ERROR_TYPES.indexOf(v.code);
  if (i < 0) throw E.na;
  return i + 1;
});

