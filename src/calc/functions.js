// The worksheet functions the engine knows, as Excel defines them. Each gets its arguments worked out (references as
// Area, arrays as Arr, an argument left out as MISSING) and returns a value or throws an error value. A function the
// engine doesn't know leaves its cell with the value Excel saved (values.js, Unsupported).
import { Area, Arr, Union, MISSING, XlError, E, scalar, matrix, toNum, toText, toBool, compare, finite, textToNumber, closeToZero } from './values.js';

export const FUNCTIONS = Object.create(null);
/** `names`, space separated. `lift`: a function of single values, worked element by element when an array formula
 *  gives it arrays; `array`: its arguments are worked out as arrays (SUMPRODUCT); `volatile`: worked out again at every
 *  change (NOW, OFFSET). */
export const def = (names, fn, opts = {}) => { for (const n of names.split(' ')) FUNCTIONS[n] = { fn, ...opts }; };
export const lifted = (names, fn) => def(names, fn, { lift: true });

export const given = (v) => v !== undefined && v !== MISSING;
/** One value of an argument (implicit intersection for a range); an error is thrown. */
export const one = (v, ctx) => {
  const x = scalar(v, ctx);
  if (x instanceof XlError) throw x;
  return x;
};
export const num = (v, ctx) => toNum(one(v, ctx), ctx.date1904);
export const text = (v, ctx) => toText(one(v, ctx));
export const bool = (v, ctx) => toBool(one(v, ctx));
export const opt = (args, i, ctx, dflt, as = num) => (given(args[i]) ? as(args[i], ctx) : args[i] === MISSING ? as(0, ctx) : dflt);

// ---- Reading ranges

/** The cells of an area that hold something, as [r, c, value], in rows then columns: by position for a small area,
 *  else from the sheet's stored rows (=SUM(A:A) doesn't visit a million empty cells). */
export function* cellsIn(a, ctx) {
  const ext = ctx.extent(a.si);
  const r1 = Math.min(a.r1, ext.maxRow);
  const c1 = Math.min(a.c1, ext.maxCol);
  if (r1 < a.r0 || c1 < a.c0) return;
  if ((r1 - a.r0 + 1) * (c1 - a.c0 + 1) <= 4096) {
    for (let r = a.r0; r <= r1; r++) {
      for (let c = a.c0; c <= c1; c++) {
        const v = ctx.cellValue(a.si, r, c);
        if (v !== null) yield [r, c, v];
      }
    }
    return;
  }
  const rows = ctx.rowsOf(a.si);
  for (const r of [...rows.keys()].filter((k) => k >= a.r0 && k <= r1).sort((x, y) => x - y)) {
    for (const c of [...rows.get(r).keys()].filter((k) => k >= a.c0 && k <= c1).sort((x, y) => x - y)) {
      const v = ctx.cellValue(a.si, r, c);
      if (v !== null) yield [r, c, v];
    }
  }
}

/** Every value the arguments hold, { v, direct }: `direct` for one given as an argument itself, not read from a
 *  reference or an array (SUM counts "2" and TRUE typed as arguments, not in cells). */
export function* values(args, ctx) {
  for (const a of args) {
    if (a instanceof Area) { for (const [, , v] of cellsIn(a, ctx)) yield { v, direct: false }; }
    else if (a instanceof Union) { for (const x of a.areas) for (const [, , v] of cellsIn(x, ctx)) yield { v, direct: false }; }
    else if (a instanceof Arr) { for (const row of a.rows) for (const v of row) yield { v, direct: false }; }
    else yield { v: a, direct: true };
  }
}

/** The numbers of the arguments, as SUM, AVERAGE, MIN and MAX see them. */
export function numbers(args, ctx) {
  const out = [];
  for (const { v, direct } of values(args, ctx)) {
    if (v instanceof XlError) throw v;
    if (typeof v === 'number') out.push(v);
    else if (direct && v !== null) out.push(v === MISSING ? 0 : toNum(v, ctx.date1904));
  }
  return out;
}

/** A range or array as rows and columns with element access. */
export function grid(v, ctx) {
  if (v instanceof Area) {
    const ext = ctx.extent(v.si);
    // (Rows beyond the last one used are empty: searches stop there.)
    const used = Math.max(0, Math.min(v.rows, ext.maxRow - v.r0 + 1));
    return { rows: v.rows, cols: v.cols, used, at: (i, j) => ctx.cellValue(v.si, v.r0 + i, v.c0 + j), area: v };
  }
  if (v instanceof Arr) return { rows: v.height, cols: v.width, used: v.height, at: (i, j) => v.rows[i][j] };
  if (v instanceof Union) throw E.value;
  const x = one(v, ctx);
  return { rows: 1, cols: 1, used: 1, at: () => x };
}

/** A row or column of a range or array as a list: { length, at(i) }; null if it is two-dimensional. */
export function vector(v, ctx) {
  const g = grid(v, ctx);
  if (g.rows === 1) return { length: g.cols, at: (i) => g.at(0, i), vertical: false, g };
  if (g.cols === 1) return { length: Math.min(g.rows, Math.max(g.used, 1)), at: (i) => g.at(i, 0), vertical: true, g };
  return null;
}

// ---- Rounding, as Excel rounds: on the value's 15 significant decimal digits, so ROUND(2.675, 2) is 2.68.

export function roundDec(x, n, how) {
  if (!Number.isFinite(x) || x === 0) return x;
  n = Math.trunc(n);
  const [mant, e = '0'] = Math.abs(x).toPrecision(15).split('e');
  const shifted = Number(`${mant}e${Number(e) + n}`);
  if (shifted >= 1e15) return x;
  const r = how === 'up' ? Math.ceil(shifted) : how === 'down' ? Math.floor(shifted) : Math.floor(shifted + 0.5);
  return Math.sign(x) * Number(`${r}e${-n}`);
}
const digits15 = (x) => Number(x.toPrecision(15));

// ---- Math and statistics

def('SUM', (args, ctx) => numbers(args, ctx).reduce((s, x) => closeToZero(s, x, s + x), 0));
def('PRODUCT', (args, ctx) => { const n = numbers(args, ctx); return n.length ? finite(n.reduce((p, x) => p * x, 1)) : 0; });
def('SUMSQ', (args, ctx) => numbers(args, ctx).reduce((s, x) => s + x * x, 0));
def('AVERAGE', (args, ctx) => { const n = numbers(args, ctx); if (!n.length) throw E.div0; return n.reduce((s, x) => s + x, 0) / n.length; });
def('MIN', (args, ctx) => { const n = numbers(args, ctx); return n.length ? Math.min(...n) : 0; });
def('MAX', (args, ctx) => { const n = numbers(args, ctx); return n.length ? Math.max(...n) : 0; });
def('MEDIAN', (args, ctx) => {
  const n = numbers(args, ctx).sort((a, b) => a - b);
  if (!n.length) throw E.num;
  const h = n.length >> 1;
  return n.length % 2 ? n[h] : (n[h - 1] + n[h]) / 2;
});
const kth = (smallest) => (args, ctx) => {
  const n = numbers([args[0]], ctx).sort((a, b) => (smallest ? a - b : b - a));
  const k = Math.ceil(num(args[1], ctx));
  if (k < 1 || k > n.length) throw E.num;
  return n[k - 1];
};
def('LARGE', kth(false));
def('SMALL', kth(true));
const variance = (sample) => (args, ctx) => {
  const n = numbers(args, ctx);
  if (n.length < (sample ? 2 : 1)) throw E.div0;
  const mean = n.reduce((s, x) => s + x, 0) / n.length;
  return n.reduce((s, x) => s + (x - mean) ** 2, 0) / (n.length - (sample ? 1 : 0));
};
def('VAR VAR.S', variance(true));
def('VAR.P VARP', variance(false));
def('STDEV STDEV.S', (args, ctx) => Math.sqrt(variance(true)(args, ctx)));
def('STDEV.P STDEVP', (args, ctx) => Math.sqrt(variance(false)(args, ctx)));
def('COUNT', (args, ctx) => {
  let n = 0;
  for (const { v, direct } of values(args, ctx)) {
    if (typeof v === 'number') n++;
    // (An argument left out counts: =COUNT(1,) is 2.)
    else if (direct && (v === MISSING || typeof v === 'boolean' || (typeof v === 'string' && textToNumber(v, ctx.date1904) != null))) n++;
  }
  return n;
});
def('COUNTA', (args, ctx) => {
  let n = 0;
  for (const { v } of values(args, ctx)) if (v !== null) n++;
  return n;
});
def('COUNTBLANK', (args, ctx) => {
  const a = args[0];
  if (!(a instanceof Area)) throw E.value;
  let filled = 0;
  for (const [, , v] of cellsIn(a, ctx)) if (v !== '') filled++;
  return a.size - filled;
});
def('SUMPRODUCT', (args, ctx) => {
  const ms = args.map((a) => matrix(a, ctx));
  const rows = ms[0].length;
  const cols = ms[0][0].length;
  if (ms.some((m) => m.length !== rows || m[0].length !== cols)) throw E.value;
  let sum = 0;
  for (let i = 0; i < rows; i++) {
    for (let j = 0; j < cols; j++) {
      let p = 1;
      for (const m of ms) {
        const v = m[i][j];
        if (v instanceof XlError) throw v;
        p *= typeof v === 'number' ? v : 0;
      }
      sum += p;
    }
  }
  return sum;
}, { array: true });

lifted('ABS', (a, ctx) => Math.abs(num(a[0], ctx)));
lifted('SIGN', (a, ctx) => Math.sign(num(a[0], ctx)));
lifted('INT', (a, ctx) => Math.floor(digits15(num(a[0], ctx))));
lifted('SQRT', (a, ctx) => { const x = num(a[0], ctx); if (x < 0) throw E.num; return Math.sqrt(x); });
lifted('EXP', (a, ctx) => finite(Math.exp(num(a[0], ctx))));
lifted('LN', (a, ctx) => { const x = num(a[0], ctx); if (x <= 0) throw E.num; return Math.log(x); });
lifted('LOG10', (a, ctx) => { const x = num(a[0], ctx); if (x <= 0) throw E.num; return Math.log10(x); });
lifted('LOG', (a, ctx) => {
  const x = num(a[0], ctx);
  const b = opt(a, 1, ctx, 10);
  if (x <= 0 || b <= 0) throw E.num;
  if (b === 1) throw E.div0;
  return b === 10 ? Math.log10(x) : Math.log(x) / Math.log(b);
});
lifted('POWER', (a, ctx) => {
  const x = num(a[0], ctx);
  const y = num(a[1], ctx);
  if (x === 0 && y === 0) throw E.num;
  if (x === 0 && y < 0) throw E.div0;
  if (x < 0 && !Number.isInteger(y)) throw E.num;
  return finite(x ** y);
});
def('PI', () => Math.PI);
def('RAND', () => Math.random(), { volatile: true });
def('RANDBETWEEN', (a, ctx) => {
  const lo = Math.ceil(num(a[0], ctx));
  const hi = Math.floor(num(a[1], ctx));
  if (lo > hi) throw E.num;
  return lo + Math.floor(Math.random() * (hi - lo + 1));
}, { volatile: true });
lifted('MOD', (a, ctx) => {
  const x = num(a[0], ctx);
  const y = num(a[1], ctx);
  if (y === 0) throw E.div0;
  return x - y * Math.floor(x / y);
});
lifted('QUOTIENT', (a, ctx) => { const y = num(a[1], ctx); if (y === 0) throw E.div0; return Math.trunc(num(a[0], ctx) / y); });
lifted('ROUND', (a, ctx) => roundDec(num(a[0], ctx), num(a[1], ctx), 'half'));
lifted('ROUNDUP', (a, ctx) => roundDec(num(a[0], ctx), num(a[1], ctx), 'up'));
lifted('ROUNDDOWN', (a, ctx) => roundDec(num(a[0], ctx), num(a[1], ctx), 'down'));
lifted('TRUNC', (a, ctx) => roundDec(num(a[0], ctx), opt(a, 1, ctx, 0), 'down'));
lifted('CEILING', (a, ctx) => {
  const x = num(a[0], ctx);
  const s = num(a[1], ctx);
  if (s === 0 || x === 0) return 0;
  if (x > 0 && s < 0) throw E.num;
  return Math.ceil(digits15(x / s)) * s;
});
lifted('FLOOR', (a, ctx) => {
  const x = num(a[0], ctx);
  const s = num(a[1], ctx);
  if (x > 0 && s < 0) throw E.num;
  if (s === 0) { if (x === 0) return 0; throw E.div0; }
  return Math.floor(digits15(x / s)) * s;
});
lifted('MROUND', (a, ctx) => {
  const x = num(a[0], ctx);
  const m = num(a[1], ctx);
  if (m === 0 || x === 0) return 0;
  if (Math.sign(x) !== Math.sign(m)) throw E.num;
  return roundDec(x / m, 0, 'half') * m;
});
lifted('EVEN', (a, ctx) => { const x = num(a[0], ctx); let n = Math.ceil(digits15(Math.abs(x))); if (n % 2) n++; return Math.sign(x) * n; });
lifted('ODD', (a, ctx) => { const x = num(a[0], ctx); let n = Math.ceil(digits15(Math.abs(x))); if (n % 2 === 0) n++; return x < 0 ? -n : n; });
lifted('FACT', (a, ctx) => {
  const n = Math.trunc(num(a[0], ctx));
  if (n < 0) throw E.num;
  let p = 1;
  for (let k = 2; k <= n; k++) p *= k;
  return finite(p);
});

// ---- Logic

def('TRUE', () => true);
def('FALSE', () => false);
def('IF', (a, ctx, mode) => {
  const yes = a.length < 2 ? true : a[1] === MISSING ? 0 : a[1];
  const no = a.length < 3 ? false : a[2] === MISSING ? 0 : a[2];
  const test = (c) => { if (c instanceof XlError) return c; try { return toBool(c) ? yes : no; } catch (e) { return e; } };
  if (mode === 'array' && (a[0] instanceof Arr || (a[0] instanceof Area && a[0].size > 1))) {
    const m = matrix(a[0], ctx);
    const pickFrom = (v, i, j) => { if (!(v instanceof Arr || v instanceof Area)) return v; const mm = matrix(v, ctx); return mm[mm.length === 1 ? 0 : i]?.[mm[0].length === 1 ? 0 : j] ?? E.na; };
    return new Arr(m.map((row, i) => row.map((c, j) => { const t = test(c); return t === yes ? pickFrom(yes, i, j) : t === no ? pickFrom(no, i, j) : t; })));
  }
  return test(one(a[0], ctx));
});
const orElse = (which) => (a, ctx) => {
  const v = scalar(a[0], ctx);
  const alt = a[1] === MISSING || a[1] === undefined ? 0 : a[1];
  return v instanceof XlError && (which === 'any' || v === E.na) ? alt : v;
};
lifted('IFERROR', orElse('any'));
lifted('IFNA', orElse('na'));
/** The TRUE and FALSE of AND, OR and XOR's arguments: numbers count, text in cells doesn't; none at all is #VALUE!. */
function logicals(args, ctx) {
  const out = [];
  for (const { v, direct } of values(args, ctx)) {
    if (v instanceof XlError) throw v;
    if (typeof v === 'boolean') out.push(v);
    else if (typeof v === 'number') out.push(v !== 0);
    else if (direct && typeof v === 'string') out.push(toBool(v));
  }
  if (!out.length) throw E.value;
  return out;
}
def('AND', (a, ctx) => logicals(a, ctx).every(Boolean));
def('OR', (a, ctx) => logicals(a, ctx).some(Boolean));
def('XOR', (a, ctx) => logicals(a, ctx).filter(Boolean).length % 2 === 1);
lifted('NOT', (a, ctx) => !bool(a[0], ctx));
def('IFS', (a, ctx) => {
  for (let i = 0; i + 1 < a.length; i += 2) if (bool(a[i], ctx)) return a[i + 1] === MISSING ? 0 : a[i + 1];
  throw E.na;
});
def('SWITCH', (a, ctx) => {
  const x = one(a[0], ctx);
  let i = 1;
  for (; i + 1 < a.length; i += 2) {
    const v = one(a[i], ctx);
    if (typeof v === typeof x && compare(v, x) === 0) return a[i + 1];
  }
  if (i < a.length) return a[i];
  throw E.na;
});

// ---- Information

const kind = (a, ctx) => scalar(a[0], ctx);
lifted('ISBLANK', (a, ctx) => kind(a, ctx) === null);
lifted('ISNUMBER', (a, ctx) => typeof kind(a, ctx) === 'number');
lifted('ISTEXT', (a, ctx) => typeof kind(a, ctx) === 'string');
lifted('ISNONTEXT', (a, ctx) => typeof kind(a, ctx) !== 'string');
lifted('ISLOGICAL', (a, ctx) => typeof kind(a, ctx) === 'boolean');
lifted('ISERROR', (a, ctx) => kind(a, ctx) instanceof XlError);
lifted('ISERR', (a, ctx) => { const v = kind(a, ctx); return v instanceof XlError && v !== E.na; });
lifted('ISNA', (a, ctx) => kind(a, ctx) === E.na);
def('ISREF', (a) => a[0] instanceof Area || a[0] instanceof Union);
lifted('ISEVEN', (a, ctx) => Math.trunc(num(a[0], ctx)) % 2 === 0);
lifted('ISODD', (a, ctx) => Math.abs(Math.trunc(num(a[0], ctx)) % 2) === 1);
def('NA', () => E.na);
lifted('N', (a, ctx) => { const v = one(a[0], ctx); return typeof v === 'number' ? v : typeof v === 'boolean' ? +v : 0; });
lifted('T', (a, ctx) => { const v = one(a[0], ctx); return typeof v === 'string' ? v : ''; });
