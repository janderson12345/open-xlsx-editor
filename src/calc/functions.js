// The worksheet functions the engine knows, as Excel defines them. Each gets its arguments worked out (references as
// Area, arrays as Arr, an argument left out as MISSING) and returns a value or throws an error value. A function the
// engine doesn't know leaves its cell with the value Excel saved (values.js, Unsupported).
import { Area, Arr, Union, MISSING, Unsupported, XlError, E, scalar, matrix, toNum, toText, toBool, compare, finite, textToNumber,
  dateSerial, closeToZero } from './values.js';
import { tokenize } from './tokens.js';
import { formatValue, dateParts } from '../ooxml/format.js';

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
const bool = (v, ctx) => toBool(one(v, ctx));
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

// ---- Text

/** A wildcard pattern (* any run, ? one character, ~ before either to mean itself) as a whole-text regular expression,
 *  case ignored. */
export function wildcard(pattern) {
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && '*?~'.includes(pattern[i + 1] || '-')) re += escapeRe(pattern[++i]);
    else if (ch === '*') re += '[\\s\\S]*';
    else if (ch === '?') re += '[\\s\\S]';
    else re += escapeRe(ch);
  }
  return new RegExp(`^${re}$`, 'i');
}
const escapeRe = (s) => s.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&');
const textLimit = (s) => { if (s.length > 32767) throw E.value; return s; };

def('CONCATENATE', (a, ctx) => textLimit(a.map((x) => text(x, ctx)).join('')));
def('CONCAT', (a, ctx) => {
  let s = '';
  for (const { v } of values(a, ctx)) { if (v instanceof XlError) throw v; s += toText(v); }
  return textLimit(s);
});
def('TEXTJOIN', (a, ctx) => {
  const delim = text(a[0], ctx);
  const skip = bool(a[1], ctx);
  const parts = [];
  for (const x of a.slice(2)) {
    const list = x instanceof Area || x instanceof Arr ? matrix(x instanceof Area ? clampArea(x, ctx) : x, ctx).flat() : [one(x, ctx)];
    for (const v of list) {
      if (v instanceof XlError) throw v;
      const s = toText(v);
      if (!(skip && s === '')) parts.push(s);
    }
  }
  return textLimit(parts.join(delim));
});
function clampArea(a, ctx) {
  const ext = ctx.extent(a.si);
  return new Area(a.si, a.r0, a.c0, Math.max(a.r0, Math.min(a.r1, ext.maxRow)), Math.max(a.c0, Math.min(a.c1, ext.maxCol)));
}
lifted('LEN', (a, ctx) => text(a[0], ctx).length);
lifted('LEFT', (a, ctx) => { const n = opt(a, 1, ctx, 1); if (n < 0) throw E.value; return text(a[0], ctx).slice(0, Math.trunc(n)); });
lifted('RIGHT', (a, ctx) => {
  const n = Math.trunc(opt(a, 1, ctx, 1));
  if (n < 0) throw E.value;
  const s = text(a[0], ctx);
  return n === 0 ? '' : s.slice(-n);
});
lifted('MID', (a, ctx) => {
  const s = text(a[0], ctx);
  const start = Math.trunc(num(a[1], ctx));
  const n = Math.trunc(num(a[2], ctx));
  if (start < 1 || n < 0) throw E.value;
  return s.substr(start - 1, n);
});
lifted('UPPER', (a, ctx) => text(a[0], ctx).toUpperCase());
lifted('LOWER', (a, ctx) => text(a[0], ctx).toLowerCase());
lifted('PROPER', (a, ctx) => {
  let after = false;
  let out = '';
  for (const ch of text(a[0], ctx)) {
    const letter = ch.toLowerCase() !== ch.toUpperCase();
    out += letter ? (after ? ch.toLowerCase() : ch.toUpperCase()) : ch;
    after = letter;
  }
  return out;
});
lifted('TRIM', (a, ctx) => text(a[0], ctx).replace(/ +/g, ' ').replace(/^ | $/g, ''));
lifted('CLEAN', (a, ctx) => text(a[0], ctx).replace(/[\x00-\x1f]/g, ''));
lifted('REPT', (a, ctx) => {
  const n = Math.trunc(num(a[1], ctx));
  if (n < 0) throw E.value;
  const s = text(a[0], ctx);
  if (s.length * n > 32767) throw E.value;
  return s.repeat(n);
});
lifted('SUBSTITUTE', (a, ctx) => {
  const s = text(a[0], ctx);
  const from = text(a[1], ctx);
  const to = text(a[2], ctx);
  if (!given(a[3])) return from ? s.split(from).join(to) : s;
  const nth = Math.trunc(num(a[3], ctx));
  if (nth < 1) throw E.value;
  if (!from) return s;
  let at = -1;
  for (let k = 0; k < nth; k++) { at = s.indexOf(from, at + 1); if (at < 0) return s; }
  return s.slice(0, at) + to + s.slice(at + from.length);
});
lifted('REPLACE', (a, ctx) => {
  const s = text(a[0], ctx);
  const start = Math.trunc(num(a[1], ctx));
  const n = Math.trunc(num(a[2], ctx));
  if (start < 1 || n < 0) throw E.value;
  return s.slice(0, start - 1) + text(a[3], ctx) + s.slice(start - 1 + n);
});
const finder = (insensitive) => (a, ctx) => {
  const what = text(a[0], ctx);
  const s = text(a[1], ctx);
  const start = Math.trunc(opt(a, 2, ctx, 1));
  if (start < 1 || start > s.length + 1) throw E.value;
  if (!insensitive) {
    const i = s.indexOf(what, start - 1);
    if (i < 0) throw E.value;
    return i + 1;
  }
  const re = new RegExp(wildcard(what).source.slice(1, -1), 'i');
  const m = re.exec(s.slice(start - 1));
  if (!m) throw E.value;
  return m.index + start;
};
lifted('FIND', finder(false));
lifted('SEARCH', finder(true));
lifted('EXACT', (a, ctx) => text(a[0], ctx) === text(a[1], ctx));
lifted('VALUE', (a, ctx) => {
  const v = one(a[0], ctx);
  if (typeof v === 'number') return v;
  if (v === null) return 0;
  if (typeof v !== 'string') throw E.value;
  const n = textToNumber(v, ctx.date1904);
  if (n == null) throw E.value;
  return n;
});
lifted('DATEVALUE', (a, ctx) => {
  const n = textToNumber(text(a[0], ctx), ctx.date1904);
  if (n == null) throw E.value;
  return Math.floor(n);
});
lifted('TEXT', (a, ctx) => {
  let v = one(a[0], ctx);
  if (v === null) v = 0;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'string') { const n = textToNumber(v, ctx.date1904); if (n != null) v = n; }
  return formatValue(v, text(a[1], ctx), { date1904: ctx.date1904 }).text;
});
lifted('CHAR', (a, ctx) => { const n = Math.trunc(num(a[0], ctx)); if (n < 1 || n > 255) throw E.value; return String.fromCharCode(n); });
lifted('CODE', (a, ctx) => { const s = text(a[0], ctx); if (!s) throw E.value; return s.charCodeAt(0); });

// ---- Lookup and reference

const sameType = (a, b) => typeof a === typeof b && !(a instanceof XlError) && !(b instanceof XlError);
/** A test for an exact match: same type, text without regard to case, wildcards in text if `wild`. */
function exactly(x, wild) {
  if (wild && typeof x === 'string' && /[*?~]/.test(x)) {
    const re = wildcard(x);
    return (v) => typeof v === 'string' && re.test(v);
  }
  return (v) => v !== null && sameType(v, x) && compare(v, x) === 0;
}
/**
 * Position of x in a list, as Excel finds it (fitted on Excel's answers for 400 unsorted lists, with blanks and text
 * among the numbers: docs/CALC.md); -1 if none. `how` 0: the first exact match. 1 (a list sorted up: MATCH 1, VLOOKUP
 * and HLOOKUP TRUE): Excel's binary search, which on a list that isn't sorted gives its own answer: from the middle,
 * a blank or a value of another type passed over to the right; an exact match taken at once, then the last of the
 * equal values that follow it. -1 (sorted down: MATCH -1): a scan from the start, blanks and other types passed over,
 * that takes the first exact match and stops at the first value below x.
 */
export function position(list, x, how, wild = true) {
  if (how === 0) {
    const eq = exactly(x, wild);
    for (let i = 0; i < list.length; i++) if (eq(list.at(i))) return i;
    return -1;
  }
  const fits = (v) => v !== null && sameType(v, x);
  if (how < 0) {
    let found = -1;
    for (let i = 0; i < list.length; i++) {
      const v = list.at(i);
      if (!fits(v)) continue;
      const k = compare(v, x);
      if (k === 0) return i;
      if (k < 0) break;
      found = i;
    }
    return found;
  }
  let lo = 0;
  let hi = list.length - 1;
  let found = -1;
  while (lo <= hi) {
    let m = (lo + hi) >> 1;
    while (m <= hi && !fits(list.at(m))) m++;
    if (m > hi) { hi = ((lo + hi) >> 1) - 1; continue; }
    const k = compare(list.at(m), x);
    if (k === 0) {
      // (The last of the equal values after it, blanks and other types passed over.)
      let last = m;
      for (let j = m + 1; j < list.length; j++) {
        const v = list.at(j);
        if (!fits(v)) continue;
        if (compare(v, x) === 0) last = j; else break;
      }
      return last;
    }
    if (k < 0) { found = m; lo = m + 1; } else hi = m - 1;
  }
  return found;
}
const lookupValue = (v, ctx) => { const x = one(v, ctx); return x === null ? 0 : x; };

function tableLookup(a, ctx, byRow) {
  const x = lookupValue(a[0], ctx);
  const g = grid(a[1], ctx);
  const k = Math.trunc(num(a[2], ctx));
  const approx = given(a[3]) ? bool(a[3], ctx) : a[3] !== MISSING;
  if (k < 1) throw E.value;
  if (k > (byRow ? g.rows : g.cols)) throw E.ref;
  const list = byRow ? { length: g.cols, at: (i) => g.at(0, i) } : { length: Math.min(g.rows, Math.max(g.used, 1)), at: (i) => g.at(i, 0) };
  const i = position(list, x, approx ? 1 : 0);
  if (i < 0) throw E.na;
  return byRow ? g.at(k - 1, i) : g.at(i, k - 1);
}
def('VLOOKUP', (a, ctx) => tableLookup(a, ctx, false));
def('HLOOKUP', (a, ctx) => tableLookup(a, ctx, true));
def('MATCH', (a, ctx) => {
  const x = lookupValue(a[0], ctx);
  const list = vector(a[1], ctx);
  if (!list) throw E.na;
  const how = Math.sign(opt(a, 2, ctx, 1));
  const i = position(list, x, how);
  if (i < 0) throw E.na;
  return i + 1;
});
def('INDEX', (a, ctx) => {
  const src = a[0];
  let r = given(a[1]) ? Math.trunc(num(a[1], ctx)) : 0;
  let c = given(a[2]) ? Math.trunc(num(a[2], ctx)) : 0;
  const rows = src instanceof Area ? src.rows : src instanceof Arr ? src.height : 1;
  const cols = src instanceof Area ? src.cols : src instanceof Arr ? src.width : 1;
  // (One index into a single row picks a column: INDEX(A1:E1, 3) is C1.)
  if (a.length === 2 && rows === 1) { c = r; r = 0; if (cols === 1) c = 0; }
  if (r < 0 || c < 0) throw E.value;
  if (r > rows || c > cols) throw E.ref;
  if (src instanceof Area) {
    return new Area(src.si, r ? src.r0 + r - 1 : src.r0, c ? src.c0 + c - 1 : src.c0, r ? src.r0 + r - 1 : src.r1, c ? src.c0 + c - 1 : src.c1);
  }
  if (src instanceof Arr) {
    if (r && c) return src.rows[r - 1][c - 1];
    if (r) return new Arr([src.rows[r - 1]]);
    if (c) return new Arr(src.rows.map((row) => [row[c - 1]]));
    return src;
  }
  return one(src, ctx);
});
def('XLOOKUP', (a, ctx) => {
  const x = lookupValue(a[0], ctx);
  const list = vector(a[1], ctx);
  if (!list) throw E.value;
  const mode = opt(a, 4, ctx, 0);
  const search = opt(a, 5, ctx, 1);
  const order = [];
  for (let i = 0; i < list.length; i++) order.push(i);
  if (search < 0) order.reverse();
  let found = -1;
  if (mode === 0 || mode === 2) {
    const eq = exactly(x, mode === 2);
    found = order.find((i) => eq(list.at(i))) ?? -1;
  } else {
    let best = null;
    for (const i of order) {
      const v = list.at(i);
      if (v === null || !sameType(v, x)) continue;
      const k = compare(v, x);
      if (k === 0) { found = i; break; }
      if ((mode < 0 ? k < 0 : k > 0) && (best == null || (mode < 0 ? compare(v, list.at(best)) > 0 : compare(v, list.at(best)) < 0))) best = i;
    }
    if (found < 0 && best != null) found = best;
  }
  if (found < 0) { if (a[3] !== undefined && a[3] !== MISSING) return a[3]; throw E.na; }
  const ret = a[2];
  if (ret instanceof Area) {
    return list.vertical
      ? new Area(ret.si, ret.r0 + found, ret.c0, ret.r0 + found, ret.c1)
      : new Area(ret.si, ret.r0, ret.c0 + found, ret.r1, ret.c0 + found);
  }
  const g = grid(ret, ctx);
  return list.vertical ? g.at(found, 0) : g.at(0, found);
});
def('CHOOSE', (a, ctx) => {
  const i = Math.trunc(num(a[0], ctx));
  if (i < 1 || i >= a.length) throw E.value;
  return a[i] === MISSING ? 0 : a[i];
});
def('ROW', (a, ctx) => (a[0] instanceof Area ? a[0].r0 : ctx.here.r) + 1);
def('COLUMN', (a, ctx) => (a[0] instanceof Area ? a[0].c0 : ctx.here.c) + 1);
def('ROWS', (a, ctx) => (a[0] instanceof Area ? a[0].rows : a[0] instanceof Arr ? a[0].height : (one(a[0], ctx), 1)));
def('COLUMNS', (a, ctx) => (a[0] instanceof Area ? a[0].cols : a[0] instanceof Arr ? a[0].width : (one(a[0], ctx), 1)));
def('OFFSET', (a, ctx) => {
  const ref = a[0];
  if (!(ref instanceof Area)) throw E.value;
  const r0 = ref.r0 + Math.trunc(num(a[1], ctx));
  const c0 = ref.c0 + Math.trunc(num(a[2], ctx));
  const h = Math.trunc(opt(a, 3, ctx, ref.rows));
  const w = Math.trunc(opt(a, 4, ctx, ref.cols));
  if (h < 1 || w < 1 || r0 < 0 || c0 < 0 || r0 + h > 1048576 || c0 + w > 16384) throw E.ref;
  return new Area(ref.si, r0, c0, r0 + h - 1, c0 + w - 1);
}, { volatile: true });
def('INDIRECT', (a, ctx) => {
  if (given(a[1]) && !bool(a[1], ctx)) throw new Unsupported('R1C1 references');
  let toks;
  try { toks = tokenize(text(a[0], ctx)); } catch { throw E.ref; }
  const t = toks[0];
  if (toks.length !== 1 || t.type !== 'ref' || !t.area || t.book != null) throw E.ref;
  const si = t.sheet == null ? ctx.here.si : ctx.sheetIndex(t.sheet);
  if (si < 0) throw E.ref;
  return new Area(si, t.area.r0, t.area.c0, t.area.r1, t.area.c1);
}, { volatile: true });

// ---- Conditional sums and counts

/** A criterion ("Apple", ">3", "<>", "A*", 12) as a test of a cell's value, as COUNTIF reads one. */
export function criterion(c, ctx) {
  if (c instanceof XlError) return (v) => v === c;
  if (c === null || c === MISSING) c = 0;
  // (An equal-to criterion matches a number stored as text too: COUNTIF(B2:B9, 12) counts "12". The others don't.)
  const equals = (n) => (v) => (typeof v === 'number' ? compare(v, n) === 0 : typeof v === 'string' && textToNumber(v, ctx.date1904) === n);
  if (typeof c === 'number') return equals(c);
  if (typeof c === 'boolean') return (v) => v === c;
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(c);
  const op = m[1] || '=';
  const operand = m[2];
  const holds = (k) => (op === '=' ? k === 0 : op === '<>' ? k !== 0 : op === '<' ? k < 0 : op === '>' ? k > 0 : op === '<=' ? k <= 0 : k >= 0);
  const n = operand.trim() === '' ? null : textToNumber(operand, ctx.date1904);
  if (n != null) {
    if (op === '<>') return (v) => !(typeof v === 'number' && compare(v, n) === 0);
    if (op === '=') return equals(n);
    return (v) => typeof v === 'number' && holds(compare(v, n));
  }
  const upper = operand.toUpperCase();
  if (upper === 'TRUE' || upper === 'FALSE') {
    const b = upper === 'TRUE';
    if (op === '<>') return (v) => v !== b;
    return (v) => typeof v === 'boolean' && holds(compare(v, b));
  }
  if (operand.startsWith('#') && operand.toUpperCase() in ERR_NAMES) {
    const code = operand.toUpperCase();
    return op === '<>' ? (v) => !(v instanceof XlError && v.code === code) : (v) => v instanceof XlError && v.code === code;
  }
  if (op === '=' || op === '<>') {
    let match;
    if (operand === '') match = m[1] ? (v) => v === null : (v) => v === null || v === '';
    else { const re = wildcard(operand); match = (v) => typeof v === 'string' && re.test(v); }
    return op === '<>' ? (v) => !match(v) : match;
  }
  return (v) => typeof v === 'string' && holds(compare(v, operand));
}
const ERR_NAMES = { '#NULL!': 1, '#DIV/0!': 1, '#VALUE!': 1, '#REF!': 1, '#NAME?': 1, '#NUM!': 1, '#N/A': 1 };

/** Visit the offsets (i, j) where every [range, test] pair holds; ranges must have one shape. `onMatch(i, j)`.
 *  Returns how many matches lie past the cells in use (all blank), for counting. */
function eachMatch(pairs, ctx, onMatch) {
  const [first] = pairs;
  for (const [rg] of pairs) {
    if (!(rg instanceof Area)) throw E.value;
    if (rg.rows !== first[0].rows || rg.cols !== first[0].cols) throw E.value;
  }
  const rows = first[0].rows;
  const cols = first[0].cols;
  let ri = 0;
  let ci = 0;
  for (const [rg] of pairs) {
    const ext = ctx.extent(rg.si);
    ri = Math.max(ri, Math.min(rows, ext.maxRow - rg.r0 + 1));
    ci = Math.max(ci, Math.min(cols, ext.maxCol - rg.c0 + 1));
  }
  const blanksMatch = pairs.every(([, test]) => test(null));
  if (!blanksMatch && pairs.length === 1 && ri * ci > 4096) {
    // (Only cells that hold something can match: visit those.)
    const [rg, test] = first;
    for (const [r, c, v] of cellsIn(rg, ctx)) if (test(v)) onMatch(r - rg.r0, c - rg.c0);
    return 0;
  }
  for (let i = 0; i < ri; i++) {
    for (let j = 0; j < ci; j++) {
      if (pairs.every(([rg, test]) => test(ctx.cellValue(rg.si, rg.r0 + i, rg.c0 + j)))) onMatch(i, j);
    }
  }
  return blanksMatch ? rows * cols - Math.max(0, ri) * Math.max(0, ci) : 0;
}
const pairsFrom = (a, from, ctx) => {
  const pairs = [];
  for (let i = from; i + 1 < a.length; i += 2) pairs.push([a[i], criterion(one(a[i + 1], ctx), ctx)]);
  return pairs;
};
/** The range to add up or average for SUMIF and AVERAGEIF: the same shape as the criteria range, from its top-left. */
const sameShape = (range, like) => (range instanceof Area ? new Area(range.si, range.r0, range.c0, range.r0 + like.rows - 1, range.c0 + like.cols - 1) : like);

function gather(target, pairs, ctx) {
  if (!(target instanceof Area)) throw E.value;
  const out = [];
  eachMatch(pairs, ctx, (i, j) => {
    const v = ctx.cellValue(target.si, target.r0 + i, target.c0 + j);
    if (v instanceof XlError) throw v;
    if (typeof v === 'number') out.push(v);
  });
  return out;
}
const countIfs = (a, ctx) => {
  let n = 0;
  const past = eachMatch(pairsFrom(a, 0, ctx), ctx, () => n++);
  return n + past;
};
def('COUNTIF COUNTIFS', countIfs);
def('SUMIF', (a, ctx) => {
  if (!(a[0] instanceof Area)) throw E.value;
  return gather(given(a[2]) ? sameShape(a[2], a[0]) : a[0], pairsFrom(a, 0, ctx).slice(0, 1), ctx).reduce((s, x) => s + x, 0);
});
def('SUMIFS', (a, ctx) => gather(a[0], pairsFrom(a, 1, ctx), ctx).reduce((s, x) => s + x, 0));
const mean = (list) => { if (!list.length) throw E.div0; return list.reduce((s, x) => s + x, 0) / list.length; };
def('AVERAGEIF', (a, ctx) => {
  if (!(a[0] instanceof Area)) throw E.value;
  return mean(gather(given(a[2]) ? sameShape(a[2], a[0]) : a[0], pairsFrom(a, 0, ctx).slice(0, 1), ctx));
});
def('AVERAGEIFS', (a, ctx) => mean(gather(a[0], pairsFrom(a, 1, ctx), ctx)));
def('MAXIFS', (a, ctx) => { const n = gather(a[0], pairsFrom(a, 1, ctx), ctx); return n.length ? Math.max(...n) : 0; });
def('MINIFS', (a, ctx) => { const n = gather(a[0], pairsFrom(a, 1, ctx), ctx); return n.length ? Math.min(...n) : 0; });

// ---- Dates and times (serial numbers: days since 1900-01-00, or 1904-01-01 in the 1904 system)

export const serialOf = (v, ctx) => { const n = num(v, ctx); if (n < 0) throw E.num; return n; };
export const parts = (v, ctx) => dateParts(serialOf(v, ctx), ctx.date1904);
const daysIn = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
function nowSerial(ctx) {
  const d = new Date();
  return (d.getTime() - d.getTimezoneOffset() * 60000) / 86400000 + 25569 - (ctx.date1904 ? 1462 : 0);
}
lifted('DATE', (a, ctx) => {
  let y = Math.trunc(num(a[0], ctx));
  if (y < 0 || y > 9999) throw E.num;
  if (y < 1900 && !ctx.date1904) y += 1900;
  else if (y < 1904 && ctx.date1904) y += 1900;
  const n = dateSerial(y, Math.trunc(num(a[1], ctx)), Math.trunc(num(a[2], ctx)), ctx.date1904);
  if (n < 0) throw E.num;
  return n;
});
lifted('TIME', (a, ctx) => {
  const s = Math.trunc(num(a[0], ctx)) * 3600 + Math.trunc(num(a[1], ctx)) * 60 + Math.trunc(num(a[2], ctx));
  if (s < 0) throw E.num;
  // (As a fraction of a day with the whole days taken off: =TIME(25,0,0) is 1.0416... - 1, 0.04166666666666674.)
  const d = s / 86400;
  return d - Math.floor(d);
});
lifted('YEAR', (a, ctx) => parts(a[0], ctx).y);
lifted('MONTH', (a, ctx) => parts(a[0], ctx).m);
lifted('DAY', (a, ctx) => parts(a[0], ctx).d);
lifted('HOUR', (a, ctx) => parts(a[0], ctx).H);
lifted('MINUTE', (a, ctx) => parts(a[0], ctx).M);
lifted('SECOND', (a, ctx) => parts(a[0], ctx).S);
lifted('WEEKDAY', (a, ctx) => {
  const w = parts(a[0], ctx).weekday; // (0 is Sunday)
  const type = Math.trunc(opt(a, 1, ctx, 1));
  if (type === 1) return w + 1;
  if (type === 2) return ((w + 6) % 7) + 1;
  if (type === 3) return (w + 6) % 7;
  throw E.num;
});
const monthsOn = (a, ctx, end) => {
  const p = parts(a[0], ctx);
  const total = p.y * 12 + (p.m - 1) + Math.trunc(num(a[1], ctx));
  const y = Math.floor(total / 12);
  const m = (total % 12) + 1;
  const n = dateSerial(y, m, end ? daysIn(y, m) : Math.min(p.d, daysIn(y, m)), ctx.date1904);
  if (n < 0) throw E.num;
  return n;
};
lifted('EDATE', (a, ctx) => monthsOn(a, ctx, false));
lifted('EOMONTH', (a, ctx) => monthsOn(a, ctx, true));
lifted('DAYS', (a, ctx) => Math.floor(num(a[0], ctx)) - Math.floor(num(a[1], ctx)));
def('TODAY', (a, ctx) => Math.floor(nowSerial(ctx)), { volatile: true });
def('NOW', (a, ctx) => nowSerial(ctx), { volatile: true });
