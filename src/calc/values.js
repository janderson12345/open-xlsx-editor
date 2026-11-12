// Excel's values inside a calculation: numbers, text, TRUE and FALSE, errors, blanks (null), references (Area) and
// arrays (Arr), and how one is turned into another where an operator or a function needs it.
import { decimal } from '../ooxml/numfmt.js';
import { ERROR_CODES } from './tokens.js';

export class XlError {
  constructor(code) { this.code = code; }
  toString() { return this.code; }
}
const BY_CODE = Object.fromEntries(ERROR_CODES.map((c) => [c, new XlError(c)]));
/** The error value for a code ("#N/A"). */
export const errorOf = (code) => BY_CODE[code] || new XlError(code);
export const E = {
  null: BY_CODE['#NULL!'], div0: BY_CODE['#DIV/0!'], value: BY_CODE['#VALUE!'], ref: BY_CODE['#REF!'],
  name: BY_CODE['#NAME?'], num: BY_CODE['#NUM!'], na: BY_CODE['#N/A'], spill: BY_CODE['#SPILL!'], calc: BY_CODE['#CALC!'],
};
export const isError = (v) => v instanceof XlError;

/** Something the engine can't work out (a link to another workbook, a function it doesn't know): the cell keeps the
 *  value Excel last saved, and Excel is asked to recalculate when it opens the file. */
export class Unsupported extends Error {}

/** A rectangle of cells on sheet `si`, corners in order. */
export class Area {
  constructor(si, r0, c0, r1, c1) {
    this.si = si;
    this.r0 = Math.min(r0, r1); this.r1 = Math.max(r0, r1);
    this.c0 = Math.min(c0, c1); this.c1 = Math.max(c0, c1);
  }
  get rows() { return this.r1 - this.r0 + 1; }
  get cols() { return this.c1 - this.c0 + 1; }
  get size() { return this.rows * this.cols; }
  has(si, r, c) { return si === this.si && r >= this.r0 && r <= this.r1 && c >= this.c0 && c <= this.c1; }
}

/** An array of values (an array constant, or what an operator gives over ranges): rows of values. */
export class Arr {
  constructor(rows) { this.rows = rows; }
  get height() { return this.rows.length; }
  get width() { return this.rows[0]?.length || 0; }
}

/** A LAMBDA as a value: its parameters' names (lower case, without _xlpm.), its body, and the names it saw. */
export class Lambda {
  constructor(params, body, scope) { this.params = params; this.body = body; this.scope = scope; }
}
/** A function named where a LAMBDA goes (an eta LAMBDA: =MAP(A1:A3,SQRT), =GROUPBY(...,SUM)), in the file _xleta.SUM. */
export class Builtin extends Lambda {
  constructor(name) { super(null, null, null); this.name = name; }
}
/** LAMBDAs side by side, HSTACK(SUM,AVERAGE): each a column of GROUPBY's and PIVOTBY's answer. */
export class Lambdas {
  constructor(list) { this.list = list; }
}

/** An argument left out (=IF(A1,,2)). */
export const MISSING = Object.freeze({ missing: true });
/** Several areas (=SUM((A1,C1))). */
export class Union {
  constructor(areas) { this.areas = areas; }
}
/** A 3D reference (=SUM(Jan:Apr!A1)): the area on each sheet from the first to the last, in the workbook's order.
 *  Functions of values read it as they read several areas; where one value or an array is wanted, Excel's #REF!. */
export class Ref3D extends Union {}

/** A cell of the model as a value. */
export function cellValue(cell) {
  if (!cell || cell.v == null) return null;
  switch (cell.t) {
    case 'n': return cell.v;
    case 'b': return cell.v;
    case 'e': return errorOf(String(cell.v));
    default: return String(cell.v);
  }
}

/** A reference or array where one value is wanted: a single cell its value; a row or column of cells the one in the
 *  formula's own column or row (implicit intersection, as Excel does in a formula that isn't an array formula), else
 *  #VALUE!; an array its first value. */
export function scalar(v, ctx) {
  // (A range or array cut to one value: what an old formula does where Excel 365 shows an @, legacyat.js.)
  if (ctx.shrunk && ((v instanceof Area && v.size > 1) || (v instanceof Arr && (v.height > 1 || v.width > 1)))) ctx.shrunk(v);
  if (v instanceof Area) {
    if (v.size === 1) return ctx.cellValue(v.si, v.r0, v.c0);
    const { r, c } = ctx.here;
    if (v.rows === 1 && c >= v.c0 && c <= v.c1) return ctx.cellValue(v.si, v.r0, c);
    if (v.cols === 1 && r >= v.r0 && r <= v.r1) return ctx.cellValue(v.si, r, v.c0);
    if (r >= v.r0 && r <= v.r1 && c >= v.c0 && c <= v.c1) return ctx.cellValue(v.si, r, c);
    return E.value;
  }
  if (v instanceof Arr) return v.rows[0][0];
  if (v instanceof Ref3D) return E.ref;
  if (v instanceof Union) return E.value;
  return v;
}

/** The values of a reference, array or single value as rows. */
export function matrix(v, ctx) {
  if (v instanceof Area) {
    const out = [];
    for (let r = v.r0; r <= v.r1; r++) {
      const row = [];
      for (let c = v.c0; c <= v.c1; c++) row.push(ctx.cellValue(v.si, r, c));
      out.push(row);
    }
    return out;
  }
  if (v instanceof Arr) return v.rows;
  if (v instanceof Ref3D) throw E.ref;
  return [[v === MISSING ? null : v]];
}

// Text that reads as a number: signs, thousands separators, a decimal point, an exponent, a trailing %, a $ in front,
// and (5) for -5, with spaces around; dates (2024-01-02, 1/2/2024) and times (10:30, 2:15 PM).
const RE_NUMERIC = /^([+-]?)\(?\$?([+-]?)((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE]([+-]?\d+))?\)?(%?)$/;
const RE_ISO = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;
const RE_US = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/;
const RE_TIME = /^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*([AaPp][Mm])?$/;

/** A date's serial number: 1900 system (with its 29 February 1900) or 1904. */
export function dateSerial(y, m, d, date1904 = false) {
  const ms = Date.UTC(y, m - 1, 1) + (d - 1) * 86400000;
  if (date1904) return Math.round((ms - Date.UTC(1904, 0, 1)) / 86400000);
  const n = Math.round((ms - Date.UTC(1899, 11, 31)) / 86400000);
  return n >= 60 ? n + 1 : n;
}

/** The number some text reads as, or null. */
export function textToNumber(text, date1904 = false) {
  const s = text.trim();
  let m = RE_NUMERIC.exec(s);
  if (m) {
    const paren = s.startsWith('(') || /^[+-]\(/.test(s);
    if (paren !== s.endsWith(')') || (m[1] && m[2])) return null;
    let v = Number(`${m[3].replace(/,/g, '')}${m[4] ? `e${m[4]}` : ''}`);
    if (m[1] === '-' || m[2] === '-' || paren) v = -v;
    if (m[5]) v /= 100;
    return Number.isFinite(v) ? v : null;
  }
  const time = (h, mi, sec, ampm) => {
    let hours = Number(h);
    if (ampm) { if (hours > 12) return null; hours = (hours % 12) + (/p/i.test(ampm) ? 12 : 0); }
    return (hours * 3600 + Number(mi) * 60 + Number(sec || 0)) / 86400;
  };
  if ((m = RE_ISO.exec(s))) {
    const day = dateSerial(Number(m[1]), Number(m[2]), Number(m[3]), date1904);
    return Number(m[2]) <= 12 && Number(m[3]) <= 31 ? day + (m[4] ? time(m[4], m[5], m[6]) : 0) : null;
  }
  if ((m = RE_US.exec(s))) {
    let y = Number(m[3]);
    if (m[3].length === 2) y += y < 30 ? 2000 : 1900;
    return Number(m[1]) <= 12 && Number(m[2]) <= 31 ? dateSerial(y, Number(m[1]), Number(m[2]), date1904) : null;
  }
  if ((m = RE_TIME.exec(s))) return time(m[1], m[2], m[3], m[4]);
  return null;
}

/** A number as text, as Excel turns one into text (=1/3&"" is "0.333333333333333"): 15 significant digits, written out
 *  in full when that takes at most 20 characters (0.000000000000000001, 12345678901234500000), else in scientific
 *  notation (1E-19, 1.23456789012345E-05, 1E+20). Measured on 33 numbers in formulas.xlsx. */
export const numberText = (n) => digitsText(n, 20);
/** A part of a complex number as text: the same, with room for 21 characters (COMPLEX(1E-19,1) is
 *  "0.0000000000000000001+i", COMPLEX(1E20,1) "100000000000000000000+i"). */
export const complexPartText = (n) => digitsText(n, 21);
function digitsText(n, width) {
  if (n === 0) return '0';
  const { digits, exp } = decimal(Math.abs(n));
  const sign = n < 0 ? '-' : '';
  let fixed;
  if (exp >= 0) {
    const frac = digits.slice(exp + 1);
    fixed = `${digits.slice(0, exp + 1).padEnd(exp + 1, '0')}${frac ? `.${frac}` : ''}`;
  } else fixed = `0.${'0'.repeat(-exp - 1)}${digits}`;
  if (fixed.length <= width) return sign + fixed;
  const e = `E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
  return `${sign}${digits[0]}${digits.length > 1 ? `.${digits.slice(1)}` : ''}${e}`;
}

/** x + y or x - y as Excel ends a formula with one: 0 when the two are the same to 15 significant digits
 *  (=0.1+0.2-0.3 is 0, though =(0.1+0.2-0.3) is 5.55E-17; =1+1E-15-1 is 0, =1-1E-15-1 isn't). SUM does it too. */
export function closeToZero(x, y, sum) {
  if (sum === 0 || x === 0 || y === 0) return sum;
  return Number(x.toPrecision(15)) === Number((-y).toPrecision(15)) ? 0 : sum;
}

/** A value as a number, where arithmetic needs one; throws the error to give otherwise. */
export function toNum(v, date1904 = false) {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v == null || v === MISSING) return 0;
  if (typeof v === 'string') {
    const n = textToNumber(v, date1904);
    if (n == null) throw E.value;
    return n;
  }
  if (v instanceof XlError) throw v;
  throw E.value;
}

/** A value as text. */
export function toText(v) {
  if (typeof v === 'string') return v;
  if (typeof v === 'number') return numberText(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v == null || v === MISSING) return '';
  if (v instanceof XlError) throw v;
  throw E.value;
}

/** A value as TRUE or FALSE. */
export function toBool(v) {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v == null || v === MISSING) return false;
  if (typeof v === 'string') {
    const u = v.toUpperCase();
    if (u === 'TRUE') return true;
    if (u === 'FALSE') return false;
    throw E.value;
  }
  if (v instanceof XlError) throw v;
  throw E.value;
}

const typeRank = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
/** Excel's ordering of two values (-1, 0, 1): numbers before text before FALSE before TRUE; text without regard to
 *  case; a blank as 0, "" or FALSE, whichever the other value is. */
export function compare(a, b) {
  if (a == null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b == null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ta = typeRank(a);
  const tb = typeRank(b);
  if (ta !== tb) return ta < tb ? -1 : 1;
  if (ta === 1) {
    const x = a.toLowerCase();
    const y = b.toLowerCase();
    return x === y ? 0 : x.localeCompare(y, 'en-US') < 0 ? -1 : 1;
  }
  if (ta === 0) {
    if (a === b) return 0;
    // (Numbers equal to 15 significant digits compare equal: =0.1+0.2=0.3 is TRUE.)
    if (a !== 0 && b !== 0 && Math.sign(a) === Math.sign(b) && Number(a.toPrecision(15)) === Number(b.toPrecision(15))) return 0;
    return a < b ? -1 : 1;
  }
  return a === b ? 0 : a ? 1 : -1;
}

/** A finished number: infinity and NaN are #NUM!. */
export const finite = (n) => (Number.isFinite(n) ? n : E.num);
