// Excel's regular expressions (PCRE2, with Unicode properties): the escapes and classes that stand for one character,
// read into JavaScript's class syntax (the v flag) for pcre.js, which reads the rest of a pattern, and pcre-match.js,
// which tests each such class one character at a time. What Excel was seen to do, with a hidden Excel: \w, \d, \s and
// \b are Unicode's (é is a letter, a no-break space a space); [[:alpha:]], \Q...\E, \x{..}, \o{..}, \h, \p{Greek},
// \p{Xwd} are read; an unknown escape such as \i, a range from or to a set ([\d-z]) and [[.a.]] are errors.
import { E } from './values.js';

export const WORD = '\\p{L}\\p{N}\\p{Mn}\\p{Pc}';
const SPACE = '\\t\\n\\v\\f\\r\\p{Z}\\u{85}';
const HSPACE = '\\t\\p{Zs}';
export const VSPACE = '\\n\\v\\f\\r\\u{85}\\u{2028}\\u{2029}';
/** Class contents for the escapes that stand for a set (\d, \w, \s, \h, \v), and their negations. */
const SETS = { d: '\\p{Nd}', w: WORD, s: SPACE, h: HSPACE, v: VSPACE };
export const POSIX = {
  alpha: '\\p{L}', digit: '\\p{Nd}', alnum: '\\p{L}\\p{N}', upper: '\\p{Lu}', lower: '\\p{Ll}', space: SPACE, word: WORD,
  punct: '\\p{P}\\u{24}\\u{2b}\\u{3c}\\u{3d}\\u{3e}\\u{5e}\\u{60}\\u{7c}\\u{7e}', xdigit: '0-9A-Fa-f', blank: HSPACE,
  cntrl: '\\p{Cc}', ascii: '\\u{0}-\\u{7F}', graph: '\\p{L}\\p{M}\\p{N}\\p{P}\\p{S}\\p{Co}',
  print: '\\p{L}\\p{M}\\p{N}\\p{P}\\p{S}\\p{Co}\\p{Zs}',
};
const CATEGORIES = new Set(['C', 'Cc', 'Cf', 'Cn', 'Co', 'Cs', 'L', 'Ll', 'Lm', 'Lo', 'Lt', 'Lu', 'M', 'Mc', 'Me', 'Mn', 'N', 'Nd',
  'Nl', 'No', 'P', 'Pc', 'Pd', 'Pe', 'Pf', 'Pi', 'Po', 'Ps', 'S', 'Sc', 'Sk', 'Sm', 'So', 'Z', 'Zl', 'Zp', 'Zs']);
/** A character as JavaScript reads it literally, anywhere: letters and digits as they are, the rest as \u{..}. */
export const lit = (ch) => (/^[A-Za-z0-9]$/.test(ch) ? ch : `\\u{${ch.codePointAt(0).toString(16)}}`);
export const code = (n) => { if (!Number.isFinite(n) || n > 0x10ffff) throw E.value; return `\\u{${n.toString(16)}}`; };

/** \p{..} or \P{..} as JavaScript's: a general category, L&, Any, or else a script. */
function property(name, negated) {
  let n = name;
  let neg = negated;
  if (n.startsWith('^')) { n = n.slice(1); neg = !neg; }
  const P = neg ? 'P' : 'p';
  if (n === 'Any') return neg ? '[]' : '[\\s\\S]';
  if (n === 'L&' || n === 'Lc') return `\\${P}{LC}`;
  // (PCRE2's own: alphanumeric, word, space, and what a universal character name can stand for: measured.)
  const special = { Xan: '\\p{L}\\p{N}', Xwd: '\\p{L}\\p{N}_', Xsp: SPACE, Xps: SPACE, Xuc: '$@`\\u{a0}-\\u{d7ff}\\u{e000}-\\u{10ffff}' }[n];
  if (special) return neg ? `[^${special}]` : `[${special}]`;
  const script = /^(?:Script|sc|Script_Extensions|scx)=([A-Z][A-Za-z_]+)$/.exec(n);
  if (script) return `\\${P}{Script=${script[1]}}`;
  if (CATEGORIES.has(n)) return `\\${P}{${n}}`;
  if (/^[A-Z][A-Za-z_]+$/.test(n)) return `\\${P}{Script=${n}}`;
  throw E.value;
}

/** One pattern being read: `p` its characters, `i` where the reading is. */
export class Reader {
  constructor(pattern, caseless) {
    this.p = [...pattern];
    this.i = 0;
    this.flags = { i: caseless, s: false, m: false, x: false };
    this.groups = []; // (the JavaScript name of each capture group, by its number: "_1" unless the pattern names it)
    this.names = new Set();
    this.pending = []; // (group numbers referred to, checked at the end)
    const p = this.p;
    // (How many capture groups the pattern has: \12 is a reference only if there are 12.)
    this.total = p.filter((ch, k) => ch === '(' && p[k - 1] !== '\\'
      && (p[k + 1] !== '?' || /^\?(P?<[A-Za-z_]|')/.test(p.slice(k + 1, k + 4).join('')))).length;
  }
  peek(k = 0) { return this.p[this.i + k]; }
  next() { return this.p[this.i++]; }
  /** The text up to `close`, after the character at i (an opening brace), moving past it. */
  braced(close = '}') {
    const end = this.p.indexOf(close, this.i + 1);
    if (end < 0) throw E.value;
    const s = this.p.slice(this.i + 1, end).join('');
    this.i = end + 1;
    return s;
  }
  digits(re, max) { let s = ''; while (s.length < max && re.test(this.peek() ?? '')) s += this.next(); return s; }
  backref(n) {
    if (n < 1) throw E.value;
    this.pending.push(n);
    return `\\k<${n <= this.groups.length ? this.groups[n - 1] : `_${n}`}>`;
  }

  /** An escape, read after its backslash; `inClass` gives class contents instead. */
  escape(inClass) {
    const c = this.next();
    if (c === undefined) throw E.value;
    const set = SETS[c.toLowerCase()];
    if (set) return c !== c.toLowerCase() ? `[^${set}]` : inClass ? set : `[${set}]`;
    const simple = { a: '\\u{7}', e: '\\u{1b}', f: '\\f', n: '\\n', r: '\\r', t: '\\t' }[c];
    if (simple) return simple;
    if (c === 'x') return code(parseInt(this.peek() === '{' ? this.braced() : this.digits(/[0-9A-Fa-f]/, 2) || '0', 16));
    if (c === 'o' && this.peek() === '{') return code(parseInt(this.braced(), 8));
    if (c === 'c') { const ch = this.next(); if (!ch) throw E.value; return code(ch.toUpperCase().charCodeAt(0) ^ 0x40); }
    if (c === 'p' || c === 'P') return property(this.peek() === '{' ? this.braced() : this.next(), c === 'P');
    if (/[0-9]/.test(c)) return this.number(c, inClass);
    if (inClass) {
      if (c === 'b') return '\\u{8}';
      if (/[A-Za-z]/.test(c)) throw E.value;
      return lit(c);
    }
    // (Assertions, \R, \N, \g, \k, \K, \G and \X are read by pcre.js before it gets here.)
    if (/[A-Za-z]/.test(c)) throw E.value;
    return lit(c);
  }
  /** \1 to \99 a reference to a group the pattern has; \0, \141 or one in a class a character in octal. */
  number(c, inClass) {
    if (c === '0' || inClass) return code(parseInt(c + this.digits(/[0-7]/, 2), 8));
    const n = c + this.digits(/[0-9]/, 9);
    if (n.length === 1 || Number(n) <= this.total) return this.backref(Number(n));
    this.i -= n.length - 1;
    const oct = c + this.digits(/[0-7]/, 2);
    if (!/^[0-7]+$/.test(oct)) throw E.value;
    return code(parseInt(oct, 8));
  }

  /** A character class, read after its [. */
  charClass(posix = POSIX) {
    let s = '[';
    if (this.peek() === '^') { s += '^'; this.i++; }
    for (let first = true; ; first = false) {
      const c = this.peek();
      if (c === undefined) throw E.value;
      if (c === ']' && !first) { this.i++; break; }
      if (c === '[' && this.peek(1) === ':') {
        const end = this.p.indexOf(']', this.i);
        const m = /^\[:(\^?)([a-z]+):\]$/.exec(this.p.slice(this.i, end + 1).join(''));
        if (m) {
          const set = posix[m[2]];
          this.posixSeen?.push(m[0]);
          if (set === undefined) throw E.value;
          // (An empty entry adds nothing, negated or not: the caseless [:upper:] and [:lower:], tested apart.)
          if (set) s += m[1] ? `[^${set}]` : set;
          this.i = end + 1;
          continue;
        }
      }
      // (Collating elements [[.a.]] and [[=a=]] are errors in PCRE2, so in Excel: measured.)
      if (c === '[' && (this.peek(1) === '.' || this.peek(1) === '=')) throw E.value;
      this.i++;
      if (c === '\\' && this.peek() === 'Q') { this.i++; s += this.quoted().join(''); continue; }
      const piece = c === '\\' ? this.escape(true) : lit(c);
      // (A range: the - between two characters; anywhere else a hyphen.)
      if (this.peek() === '-' && this.peek(1) !== undefined && this.peek(1) !== ']' && /^(\\u\{[0-9a-f]+\}|[A-Za-z0-9]|\\[fnrt])$/.test(piece)) {
        this.i++;
        const to = this.peek() === '\\' ? (this.i++, this.escape(true)) : lit(this.next());
        // ([a-\d], a range up to a set, is an error: measured.)
        if (!/^(\\u\{[0-9a-f]+\}|[A-Za-z0-9]|\\[fnrt])$/.test(to)) throw E.value;
        s += `${piece}-${to}`;
      } else if (c === '\\' && this.peek() === '-' && this.peek(1) !== undefined && this.peek(1) !== ']' && /\\[pP]\{|^\[/.test(piece)) {
        throw E.value; // ([\d-z], a range from a set, is an error: measured.)
      } else s += piece;
    }
    return `${s}]`;
  }
  /** \Q...\E's characters, each literal, read after the \Q. */
  quoted() {
    const out = [];
    while (this.i < this.p.length && !(this.p[this.i] === '\\' && this.p[this.i + 1] === 'E')) out.push(lit(this.next()));
    this.i += 2;
    return out;
  }

  /** A quantifier at i, or null: *, +, ?, {n}, {n,}, {n,m}, {,m}, then ? (lazy) or + (possessive). */
  quantifier() {
    let q;
    const c = this.peek();
    if (c === '*' || c === '+' || c === '?') { q = c; this.i++; } else if (c === '{') {
      const m = /^\{(\d*)(,(\d*))?\}/.exec(this.p.slice(this.i, this.i + 24).join(''));
      if (!m || (m[1] === '' && (m[2] === undefined || m[3] === ''))) return null;
      if (m[3] && Number(m[3]) < Number(m[1] || 0)) throw E.value;
      if (Number(m[1] || 0) > 65535 || Number(m[3] || 0) > 65535) throw E.value; // (PCRE2's largest repeat)
      q = `{${m[1] || 0}${m[2] !== undefined ? `,${m[3] ?? ''}` : ''}}`;
      this.i += m[0].length;
    } else return null;
    if (this.peek() === '?') { this.i++; return { q: `${q}?`, possessive: false }; }
    if (this.peek() === '+') { this.i++; return { q, possessive: true }; }
    return { q, possessive: false };
  }

}
