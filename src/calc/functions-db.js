// The database functions (DAVERAGE, DCOUNT, DCOUNTA, DGET, DMAX, DMIN, DPRODUCT, DSTDEV, DSTDEVP, DSUM, DVAR, DVARP):
// a list with its headings in the first row; the field a heading (case aside) or a column's number; the criteria a
// range of headings with rows of conditions under them, as Advanced Filter reads them (criteria.js): the rows are Or,
// a row's cells And, a blank cell anything. And a few more of Excel 2010's: CEILING.PRECISE, FLOOR.PRECISE,
// ISO.CEILING, SKEW.P, BINOM.DIST.RANGE, COVAR. Measured against Excel with the formula probe (test/formulas-excel.json).
import { Area, E, XlError, matrix, toText } from './values.js';
import { FUNCTIONS, def, given, num, numbers } from './functions.js';
import { cellTest } from './criteria.js';

const isNum = (v) => typeof v === 'number';
const blank = (v) => v == null || v === '';

/** The field's values in the rows of the list that the criteria pass. */
function picked(a, ctx, needField = true) {
  if (!(a[0] instanceof Area) || !(a[2] instanceof Area)) throw E.value;
  const list = matrix(a[0], ctx);
  const heads = list[0].map((h) => toText(h ?? '').trim().toLowerCase());
  let col = -1;
  if (given(a[1])) {
    const f = matrix(a[1], ctx)[0][0];
    if (f instanceof XlError) throw f;
    if (isNum(f)) { col = Math.trunc(f) - 1; if (col < 0 || col >= heads.length) throw E.value; }
    else { col = heads.indexOf(toText(f ?? '').trim().toLowerCase()); if (col < 0) throw E.value; }
  } else if (needField) throw E.value;
  const crit = matrix(a[2], ctx);
  const cols = crit[0].map((h) => heads.indexOf(toText(h ?? '').trim().toLowerCase()));
  const ways = crit.slice(1).map((row) => row.flatMap((v, k) => (blank(v) ? [] : [[cols[k], cellTest(toText(v))]])));
  const passes = (row) => ways.some((tests) => tests.every(([c, t]) => c >= 0 && t(row[c], toText(row[c] ?? ''))));
  const rows = list.slice(1).filter(passes);
  return col < 0 ? rows : rows.map((row) => row[col]);
}
const nums = (vals) => vals.filter(isNum);
def('DSUM', (a, ctx) => nums(picked(a, ctx)).reduce((s, v) => s + v, 0));
def('DPRODUCT', (a, ctx) => { const v = nums(picked(a, ctx)); return v.length ? v.reduce((s, x) => s * x, 1) : 0; });
def('DCOUNT', (a, ctx) => {
  const got = picked(a, ctx, false);
  // (No field: the records passing; a field: its numbers.)
  return given(a[1]) ? nums(got).length : got.length;
});
def('DCOUNTA', (a, ctx) => { const got = picked(a, ctx, false); return given(a[1]) ? got.filter((v) => !blank(v)).length : got.length; });
def('DAVERAGE', (a, ctx) => { const v = nums(picked(a, ctx)); if (!v.length) throw E.div0; return v.reduce((s, x) => s + x, 0) / v.length; });
def('DMAX', (a, ctx) => { const v = nums(picked(a, ctx)); return v.length ? Math.max(...v) : 0; });
def('DMIN', (a, ctx) => { const v = nums(picked(a, ctx)); return v.length ? Math.min(...v) : 0; });
def('DGET', (a, ctx) => {
  const v = picked(a, ctx);
  if (!v.length) throw E.value;
  if (v.length > 1) throw E.num;
  return v[0] ?? 0;
});
/** The deviations and variances, on the field's numbers as the plain functions take them. */
const spread = (name) => (a, ctx) => {
  const v = nums(picked(a, ctx));
  return FUNCTIONS[name].fn(v, ctx);
};
def('DSTDEV', spread('STDEV.S'));
def('DSTDEVP', spread('STDEV.P'));
def('DVAR', spread('VAR.S'));
def('DVARP', spread('VAR.P'));

// ---- CEILING.PRECISE, FLOOR.PRECISE, ISO.CEILING: to a multiple of the significance's size, up or down

const toMultiple = (up) => (a, ctx) => {
  const x = num(a[0], ctx);
  const s = Math.abs(given(a[1]) ? num(a[1], ctx) : 1);
  if (s === 0 || x === 0) return 0;
  const q = x / s;
  // (Within 15 digits of a whole multiple: that multiple.)
  const r = Math.round(q);
  const k = Math.abs(q - r) < 1e-12 * Math.max(1, Math.abs(q)) ? r : up ? Math.ceil(q) : Math.floor(q);
  return k * s;
};
def('CEILING.PRECISE ISO.CEILING', toMultiple(true));
def('FLOOR.PRECISE', toMultiple(false));

// ---- SKEW.P, BINOM.DIST.RANGE, COVAR

def('SKEW.P', (a, ctx) => {
  const xs = numbers(a, ctx);
  const n = xs.length;
  if (n === 0) throw E.div0;
  const m = xs.reduce((s, x) => s + x, 0) / n;
  const sd = Math.sqrt(xs.reduce((s, x) => s + (x - m) ** 2, 0) / n);
  if (sd === 0) throw E.div0;
  return xs.reduce((s, x) => s + ((x - m) / sd) ** 3, 0) / n;
});
def('BINOM.DIST.RANGE', (a, ctx) => {
  const n = Math.trunc(num(a[0], ctx)); const p = num(a[1], ctx); const s1 = Math.trunc(num(a[2], ctx));
  const s2 = given(a[3]) ? Math.trunc(num(a[3], ctx)) : s1;
  if (n < 0 || p < 0 || p > 1 || s1 < 0 || s1 > n || s2 < s1 || s2 > n) throw E.num;
  const dist = FUNCTIONS['BINOM.DIST'].fn;
  let total = 0;
  for (let k = s1; k <= s2; k++) total += dist([k, n, p, false], ctx);
  return total;
});
def('COVAR', (a, ctx) => FUNCTIONS['COVARIANCE.P'].fn(a, ctx));
