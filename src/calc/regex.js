// Excel's regular expressions, PCRE2 with Unicode properties, turned into JavaScript's (the v flag), as REGEXTEST,
// REGEXEXTRACT and REGEXREPLACE (functions-regex.js) need them. What Excel was seen to do, with a hidden Excel:
// \w, \d, \s and \b are Unicode's (é is a letter, a no-break space a space); $ matches before a final newline too, and .
// anything but a newline; [[:alpha:]], \Q...\E, (?P<name>...), \x{..}, \o{..}, \h, \R, a{,2}, possessive and atomic
// groups are read; an unknown escape such as \i is an error. What JavaScript has no way to say (\K, recursion,
// conditionals, \G, \X, modifiers past the start of the pattern) is Unsupported: the cell keeps Excel's value.
import { E, Unsupported } from './values.js';

const WORD = '\\p{L}\\p{N}\\p{Mn}\\p{Pc}';
const SPACE = '\\t\\n\\v\\f\\r\\p{Z}\\u{85}';
const HSPACE = '\\t\\p{Zs}';
const VSPACE = '\\n\\v\\f\\r\\u{85}\\u{2028}\\u{2029}';
/** Class contents for the escapes that stand for a set (\d, \w, \s, \h, \v), and their negations. */
const SETS = { d: '\\p{Nd}', w: WORD, s: SPACE, h: HSPACE, v: VSPACE };
const POSIX = {
  alpha: '\\p{L}', digit: '\\p{Nd}', alnum: '\\p{L}\\p{N}', upper: '\\p{Lu}', lower: '\\p{Ll}', space: SPACE, word: WORD,
  punct: '\\p{P}\\u{24}\\u{2b}\\u{3c}\\u{3d}\\u{3e}\\u{5e}\\u{60}\\u{7c}\\u{7e}', xdigit: '0-9A-Fa-f', blank: HSPACE,
  cntrl: '\\p{Cc}', ascii: '\\u{0}-\\u{7F}', graph: '\\p{L}\\p{M}\\p{N}\\p{P}\\p{S}\\p{Co}',
  print: '\\p{L}\\p{M}\\p{N}\\p{P}\\p{S}\\p{Co}\\p{Zs}',
};
const CATEGORIES = new Set(['C', 'Cc', 'Cf', 'Cn', 'Co', 'Cs', 'L', 'Ll', 'Lm', 'Lo', 'Lt', 'Lu', 'M', 'Mc', 'Me', 'Mn', 'N', 'Nd',
  'Nl', 'No', 'P', 'Pc', 'Pd', 'Pe', 'Pf', 'Pi', 'Po', 'Ps', 'S', 'Sc', 'Sk', 'Sm', 'So', 'Z', 'Zl', 'Zp', 'Zs']);
/** A character as JavaScript reads it literally, anywhere: letters and digits as they are, the rest as \u{..}. */
const lit = (ch) => (/^[A-Za-z0-9]$/.test(ch) ? ch : `\\u{${ch.codePointAt(0).toString(16)}}`);
const code = (n) => { if (!Number.isFinite(n) || n > 0x10ffff) throw E.value; return `\\u{${n.toString(16)}}`; };
const B_IN = `(?<=[${WORD}])`;
const AHEAD = `(?=[${WORD}])`;
const ASSERTIONS = {
  b: `(?:${B_IN}(?![${WORD}])|(?<![${WORD}])${AHEAD})`, B: `(?:${B_IN}${AHEAD}|(?<![${WORD}])(?![${WORD}]))`,
  A: '(?<![\\s\\S])', z: '(?![\\s\\S])', Z: '(?=\\n?(?![\\s\\S]))', E: '',
};

/** \p{..} or \P{..} as JavaScript's: a general category, L&, Any, or else a script. */
function property(name, negated) {
  let n = name;
  let neg = negated;
  if (n.startsWith('^')) { n = n.slice(1); neg = !neg; }
  const P = neg ? 'P' : 'p';
  if (n === 'Any') return neg ? '[]' : '[\\s\\S]';
  if (n === 'L&') return `\\${P}{LC}`;
  if (CATEGORIES.has(n)) return `\\${P}{${n}}`;
  if (/^[A-Z][A-Za-z_]+$/.test(n)) return `\\${P}{Script=${n}}`;
  throw E.value;
}

/** One pattern being read: `p` its characters, `i` where the reading is. */
class Reader {
  constructor(pattern, caseless) {
    this.p = [...pattern];
    this.i = 0;
    this.flags = { i: caseless, s: false, m: false, x: false };
    this.groups = []; // (the JavaScript name of each capture group, by its number: "_1" unless the pattern names it)
    this.names = new Set();
    this.atomics = 0;
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
  /** (?>...) and X++ as JavaScript can have them: a lookahead that takes the text, then the text again. */
  atomic(body) { const name = `_a${++this.atomics}`; return `(?=(?<${name}>${body}))\\k<${name}>`; }

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
    if (c in ASSERTIONS) return ASSERTIONS[c];
    if (c === 'R') return `(?:\\r\\n|[${VSPACE}])`;
    if (c === 'N') return '[^\\n]';
    if (c === 'g') return this.groupReference();
    if (c === 'k') {
      const close = { '<': '>', "'": "'", '{': '}' }[this.peek()];
      if (!close) throw E.value;
      return `\\k<${this.braced(close)}>`;
    }
    if ('KGXC'.includes(c)) throw new Unsupported(`\\${c} in a regular expression`);
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
  /** \g1, \g{1}, \g{-1} (the group before), \g{name}. */
  groupReference() {
    const ref = this.peek() === '{' ? this.braced() : this.digits(/[-0-9]/, 9);
    if (/^-\d+$/.test(ref)) return this.backref(this.groups.length + 1 + Number(ref));
    if (/^\d+$/.test(ref)) return this.backref(Number(ref));
    if (this.names.has(ref)) return `\\k<${ref}>`;
    throw E.value;
  }

  /** A character class, read after its [. */
  charClass() {
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
          const set = POSIX[m[2]];
          if (!set) throw E.value;
          s += m[1] ? `[^${set}]` : set;
          this.i = end + 1;
          continue;
        }
      }
      this.i++;
      if (c === '\\' && this.peek() === 'Q') { this.i++; s += this.quoted().join(''); continue; }
      const piece = c === '\\' ? this.escape(true) : lit(c);
      // (A range: the - between two characters; anywhere else a hyphen.)
      if (this.peek() === '-' && this.peek(1) !== undefined && this.peek(1) !== ']' && /^(\\u\{[0-9a-f]+\}|[A-Za-z0-9]|\\[fnrt])$/.test(piece)) {
        this.i++;
        const to = this.peek() === '\\' ? (this.i++, this.escape(true)) : lit(this.next());
        s += `${piece}-${to}`;
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
      q = `{${m[1] || 0}${m[2] !== undefined ? `,${m[3] ?? ''}` : ''}}`;
      this.i += m[0].length;
    } else return null;
    if (this.peek() === '?') { this.i++; return { q: `${q}?`, possessive: false }; }
    if (this.peek() === '+') { this.i++; return { q, possessive: true }; }
    return { q, possessive: false };
  }

  /** A group's opening, read after its (: { open, atomic }, or { ref } for (?P=name), or null for a comment. */
  group() {
    const ahead = this.p.slice(this.i, this.i + 40).join('');
    let m;
    if (this.peek() !== '?') { this.groups.push(`_${this.groups.length + 1}`); return { open: `(?<${this.groups.at(-1)}>` }; }
    if (ahead.startsWith('?#')) { while (this.i < this.p.length && this.p[this.i] !== ')') this.i++; this.i++; return null; }
    if ((m = /^\?(?:P?<([A-Za-z_][A-Za-z0-9_]*)>|'([A-Za-z_][A-Za-z0-9_]*)')/.exec(ahead))) {
      const name = m[1] || m[2];
      this.groups.push(name);
      this.names.add(name);
      this.i += m[0].length;
      return { open: `(?<${name}>` };
    }
    if ((m = /^\?P=([A-Za-z_][A-Za-z0-9_]*)\)/.exec(ahead))) { this.i += m[0].length; return { ref: `\\k<${m[1]}>` }; }
    if ((m = /^\?(:|=|!|<=|<!)/.exec(ahead))) { this.i += m[0].length; return { open: `(${m[0]}` }; }
    if (ahead.startsWith('?>')) { this.i += 2; return { open: '(?:', atomic: true }; }
    throw new Unsupported('this group in a regular expression');
  }

  /** The whole pattern as JavaScript's source. */
  read() {
    const { p, flags } = this;
    // Options at the very start apply to the whole pattern; anywhere else JavaScript has no way to say them.
    for (let m; (m = /^\(\?([imsxn]*)(?:-([imsxn]*))?\)/.exec(p.slice(this.i).join('')));) {
      for (const f of m[1]) flags[f] = true;
      for (const f of m[2] || '') flags[f] = false;
      this.i += m[0].length;
    }
    const out = [];
    const stack = [];
    while (this.i < p.length) {
      const c = this.next();
      if (flags.x && /\s/.test(c)) continue;
      if (flags.x && c === '#') { while (this.i < p.length && p[this.i] !== '\n') this.i++; continue; }
      // (`atom`: where the last thing a quantifier may follow begins in `out`, or -1.)
      let atom = out.length;
      if (c === '\\' && this.peek() === 'Q') { this.i++; const q = this.quoted(); atom = q.length ? out.length + q.length - 1 : -1; out.push(...q); }
      else if (c === '\\') { const e = this.escape(false); out.push(e); if (!/^\\k<|^\[|^\\[pPu]|^\\[fnrt]$|^\(\?:\\r/.test(e)) atom = -1; }
      else if (c === '[') out.push(this.charClass());
      else if (c === '(') {
        const g = this.group();
        if (g?.ref) out.push(g.ref);
        else { if (g) { stack.push({ at: out.length, atomic: g.atomic }); out.push(g.open); } continue; }
      } else if (c === ')') {
        const g = stack.pop();
        if (!g) throw E.value;
        out.push(')');
        if (g.atomic) out.push(this.atomic(out.splice(g.at).join('').slice(3, -1)));
        atom = g.at;
      } else if (c === '.') out.push(flags.s ? '[\\s\\S]' : '[^\\n]');
      else if (c === '^') { out.push(flags.m ? '(?<=^|\\n)' : '(?<![\\s\\S])'); atom = -1; }
      else if (c === '$') { out.push(flags.m ? '(?=\\n|(?![\\s\\S]))' : '(?=\\n?(?![\\s\\S]))'); atom = -1; }
      else if (c === '|') { out.push('|'); atom = -1; }
      else if (c === '*' || c === '+' || c === '?') throw E.value; // (a quantifier with nothing before it)
      else out.push(lit(c));
      const q = atom >= 0 ? this.quantifier() : null;
      if (q) { const body = out.splice(atom).join(''); out.push(q.possessive ? this.atomic(`(?:${body})${q.q}`) : `${body}${q.q}`); }
    }
    if (stack.length) throw E.value;
    for (const n of this.pending) if (n > this.groups.length) throw E.value;
    return out.join('');
  }
}

/** Reads a PCRE2 pattern into { source, flags, groups }: groups the JavaScript name of each capture group by its number
 *  (named "_1", "_2" unless the pattern names them, so that the groups added here for atomic ones don't count). */
export function translate(pattern, caseless) {
  const r = new Reader(pattern, caseless);
  const source = r.read();
  try { new RegExp(source, 'v'); } catch { throw E.value; }
  return { source, flags: `gv${r.flags.i ? 'i' : ''}`, groups: r.groups };
}

/** A regular expression from Excel's pattern; `caseless` for case_sensitivity 1. */
export function regex(pattern, caseless) {
  const t = translate(pattern, caseless);
  return { re: new RegExp(t.source, t.flags), groups: t.groups };
}
