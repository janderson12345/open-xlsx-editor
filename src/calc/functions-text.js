// Text, and dates and times, as Excel defines them: registered into functions.js's table (evaluate.js imports
// this file after it).
import { Area, Arr, Union, XlError, E, matrix, toText, textToNumber, dateSerial } from './values.js';
import { formatValue, dateParts } from '../ooxml/format.js';
import { def, lifted, given, one, num, text, opt, values, bool } from './functions.js';
import { pointsOf, unitAt, pointsBefore, ansiChar, ansiCode } from './charsets.js';

// ---- Text

/** A wildcard pattern (* any run, ? one character, ~ before either to mean itself) as a whole-text regular expression,
 *  case ignored. `one` is what ? matches (SEARCH's doesn't match half of an emoji, nor a whole one: measured). */
export function wildcard(pattern, one = '[\\s\\S]') {
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && '*?~'.includes(pattern[i + 1] || '-')) re += escapeRe(pattern[++i]);
    else if (ch === '*') re += '[\\s\\S]*';
    else if (ch === '?') re += one;
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
    // (Several areas, a 3D reference's sheets: each in turn.)
    const areas = x instanceof Union ? x.areas : [x];
    const list = areas.flatMap((y) => (y instanceof Area || y instanceof Arr ? matrix(y instanceof Area ? clampArea(y, ctx) : y, ctx).flat() : [one(y, ctx)]));
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
// (How Excel 365 counts an emoji, two UTF-16 units: LEFT and RIGHT keep it whole; LEN, MID, FIND, SEARCH and REPLACE
// count it as 2 in a workbook such as this editor writes, as their byte versions do, but as 1, a code point a
// character, where the workbook asks for function compatibility version 2, as one Excel made does (calcmode.js,
// docs/CALC.md). `byUnits` null: as the workbook says, ctx.codePoints.)
const seq = (s, byUnits) => (byUnits ? s : pointsOf(s) ?? s);
const units = (byUnits, ctx) => byUnits ?? !ctx.codePoints;
const cut = (q, from, to) => (typeof q === 'string' ? q.slice(from, to) : q.slice(from, to).join(''));
const len = (byUnits) => (a, ctx) => {
  const s = text(a[0], ctx);
  if (units(byUnits, ctx)) return s.length;
  // (Half an emoji, its first unit alone, at the very end isn't counted where characters are code points: LEN of it
  // is 0, of it and "y" 2 (measured).)
  return seq(s, false).length - (/[\ud800-\udbff]$/.test(s) ? 1 : 0);
};
// (UPPER and LOWER leave letters past U+FFFF as they are: UPPER of Deseret's small long i, U+10428, is itself.)
const cased = (s, f) => (/[\ud800-\udfff]/.test(s) ? s.replace(/[^\ud800-\udfff]+/g, f) : f(s));
const left = (a, ctx) => { const n = opt(a, 1, ctx, 1); if (n < 0) throw E.value; return cut(seq(text(a[0], ctx)), 0, Math.trunc(n)); };
const right = (a, ctx) => {
  const n = Math.trunc(opt(a, 1, ctx, 1));
  if (n < 0) throw E.value;
  const q = seq(text(a[0], ctx));
  return n === 0 ? '' : cut(q, Math.max(0, q.length - n));
};
const mid = (byUnits) => (a, ctx) => {
  const q = seq(text(a[0], ctx), units(byUnits, ctx));
  const start = Math.trunc(num(a[1], ctx));
  const n = Math.trunc(num(a[2], ctx));
  if (start < 1 || n < 0) throw E.value;
  return cut(q, start - 1, start - 1 + n);
};
lifted('LEN', len(null));
lifted('LEFT', left);
lifted('RIGHT', right);
lifted('MID', mid(null));
lifted('UPPER', (a, ctx) => cased(text(a[0], ctx), (x) => x.toUpperCase()));
lifted('LOWER', (a, ctx) => cased(text(a[0], ctx), (x) => x.toLowerCase()));
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
const replace = (byUnits) => (a, ctx) => {
  const q = seq(text(a[0], ctx), units(byUnits, ctx));
  const start = Math.trunc(num(a[1], ctx));
  const n = Math.trunc(num(a[2], ctx));
  if (start < 1 || n < 0) throw E.value;
  return cut(q, 0, start - 1) + text(a[3], ctx) + cut(q, start - 1 + n);
};
lifted('REPLACE', replace(null));
const finder = (insensitive, mode) => (a, ctx) => {
  const byUnits = units(mode, ctx);
  const what = text(a[0], ctx);
  const s = text(a[1], ctx);
  const start = Math.trunc(opt(a, 2, ctx, 1));
  // (FIND may start just past the end, SEARCH not: FIND("","abc",4) is 4, SEARCH("","abc",4) and SEARCH("","")
  // #VALUE!, measured.)
  if (start < 1 || start > seq(s, byUnits).length + (insensitive ? 0 : 1)) throw E.value;
  const from = byUnits ? start - 1 : unitAt(s, start - 1);
  let at;
  if (!insensitive) at = s.indexOf(what, from);
  else {
    // (? is one UTF-16 unit either way: where characters are code points, "??" finds an emoji at 1 (measured).)
    // A match starts only where a character does, so "?b" isn't found in "a"&emoji&"b" there (measured).
    const re = new RegExp(wildcard(what).source.slice(1, -1), byUnits ? 'i' : 'iy');
    if (byUnits) { const m = re.exec(s.slice(from)); at = m ? from + m.index : -1; } else at = boundaryMatch(re, s, from);
  }
  if (at < 0) throw E.value;
  return (byUnits ? at : pointsBefore(s, at)) + 1;
};
/** Where sticky `re` first matches `s` from UTF-16 index `from` on, trying only the starts of code points; -1 if none. */
function boundaryMatch(re, s, from) {
  for (let u = from; u <= s.length; u += s.codePointAt(u) > 0xffff ? 2 : 1) {
    re.lastIndex = u;
    if (re.test(s)) return u;
  }
  return -1;
}
lifted('FIND', finder(false, null));
lifted('SEARCH', finder(true, null));
// (The byte versions, as an English Excel 365 has them (measured, docs/CALC.md): LENB, MIDB, FINDB, SEARCHB and REPLACEB
// count UTF-16 units, an emoji 2; LEFTB and RIGHTB are LEFT and RIGHT, keeping a character whole; ASC and DBCS give the
// text as it is, full-width letters staying full-width.)
lifted('LENB', len(true));
lifted('MIDB', mid(true));
lifted('FINDB', finder(false, true));
lifted('SEARCHB', finder(true, true));
lifted('REPLACEB', replace(true));
lifted('LEFTB', left);
lifted('RIGHTB', right);
lifted('ASC', (a, ctx) => text(a[0], ctx));
lifted('DBCS', (a, ctx) => text(a[0], ctx));
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
lifted('CHAR', (a, ctx) => { const n = Math.trunc(num(a[0], ctx)); if (n < 1 || n > 255) throw E.value; return ansiChar(n); });
lifted('CODE', (a, ctx) => { const s = text(a[0], ctx); if (!s) throw E.value; return ansiCode(s); });

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
