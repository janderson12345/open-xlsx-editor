// Excel's regular expressions (PCRE2, with Unicode properties) read into a tree for pcre-match.js, the matcher
// REGEXTEST, REGEXEXTRACT and REGEXREPLACE use. Escapes and classes that stand for one character are read by
// regex.js into JavaScript's class syntax and tested one character at a time; the rest is here: groups and their
// numbers (branch reset too), \K, recursion and subroutine calls, conditionals, \G, \X, the backtracking verbs,
// options partway through a pattern (i, m, s, x, n, U, J, ^), newline conventions, lookbehind of bounded length. What
// PCRE2 refuses is #VALUE!, as Excel gives it (measured over COM, docs/CALC.md): \C, \u, \U, \L, a lookbehind of
// unbounded length, \K in a lookaround, a name used twice without (?J), callouts.
import { E } from './values.js';
import { Reader, POSIX, VSPACE } from './regex.js';

const NAME = /^[A-Za-z_][A-Za-z0-9_]*/;
const START_OPTIONS = /^\(\*(UTF|UTF8|UCP|NO_AUTO_POSSESS|NO_START_OPT|NO_DOTSTAR_ANCHOR|NO_JIT|NOTEMPTY|NOTEMPTY_ATSTART|LF|CR|CRLF|ANYCRLF|ANY|NUL|BSR_ANYCRLF|BSR_UNICODE|LIMIT_(?:MATCH|HEAP|DEPTH|RECURSION)=\d+)\)/;
const VERBS = { FAIL: 'fail', F: 'fail', ACCEPT: 'accept', COMMIT: 'commit', PRUNE: 'prune', SKIP: 'skip', THEN: 'then' };

/** Node constructors (pcre-match.js compiles them). `char`: one character from a JavaScript class source. */
const char = (src, i) => ({ t: 'char', src, i });
const seq = (items) => ({ t: 'seq', items });

class Parser extends Reader {
  constructor(pattern, caseless) {
    super(pattern, caseless);
    this.count = 0; // (capture groups so far)
    this.names = new Map(); // (name -> [numbers])
    this.refs = []; // (numbers and names referred to, checked at the end)
    this.alts = 0; // (alternations, numbered for (*THEN))
    this.newline = 'LF';
    this.posixSeen = null;
    this.verbs = [];
  }

  parse() {
    for (let m; (m = START_OPTIONS.exec(this.p.slice(this.i).join('')));) {
      if (/^(LF|CR|CRLF|ANYCRLF|ANY|NUL)$/.test(m[1])) this.newline = m[1];
      this.i += m[0].length;
    }
    const flags = { i: this.flags.i, s: false, m: false, x: false, xx: false, n: false, U: false, J: false };
    const body = this.alternation(flags, { look: false, alt: null });
    if (this.i < this.p.length) throw E.value; // (a ) with no group open)
    for (const r of this.refs) {
      if (typeof r === 'number' ? r > this.count : !this.names.has(r)) throw E.value;
    }
    return body;
  }

  /** Alternatives up to the group's ) or the end; `flags` the group's own, changed in place by (?i) and the like. */
  alternation(flags, where) {
    const id = ++this.alts;
    const alts = [];
    const inner = { ...where, alt: id };
    const verbs = this.verbs.length;
    for (;;) {
      alts.push(this.sequence(flags, inner));
      if (this.peek() !== '|') break;
      this.i++;
    }
    if (alts.length > 1) return { t: 'alt', alts, id };
    // ((*THEN) goes back to the innermost alternation that has alternatives: with none here, the one around.)
    for (const v of this.verbs.slice(verbs)) if (v.alt === id) v.alt = where.alt;
    return alts[0];
  }

  sequence(flags, where) {
    const items = [];
    while (this.i < this.p.length && this.peek() !== '|' && this.peek() !== ')') {
      const c = this.next();
      if (flags.x && /\s/.test(c)) continue;
      if (flags.x && c === '#') { while (this.i < this.p.length && this.p[this.i] !== '\n') this.i++; continue; }
      const atom = this.atom(c, flags, where);
      if (atom === null) continue;
      if (atom.t === 'quoted') { items.push(...atom.items); if (atom.items.length) this.repeat(items, flags); continue; }
      items.push(atom);
      if (['assert', 'K', 'verb', 'empty'].includes(atom.t)) {
        if (this.quantifierAhead()) throw E.value;
        continue;
      }
      this.repeat(items, flags);
    }
    return items.length === 1 ? items[0] : seq(items);
  }

  /** A quantifier after the last item, if there is one. */
  repeat(items, flags) {
    for (;;) {
      if (flags.x) while (/\s/.test(this.peek() ?? '')) this.i++;
      const q = this.quantifier();
      if (!q) return;
      const m = /^(\*|\+|\?|\{(\d+)(?:,(\d*))?\})(\?)?$/.exec(q.q);
      const [min, max] = m[1] === '*' ? [0, Infinity] : m[1] === '+' ? [1, Infinity] : m[1] === '?' ? [0, 1]
        : [Number(m[2]), m[3] === undefined ? Number(m[2]) : m[3] === '' ? Infinity : Number(m[3])];
      const lazy = q.possessive ? false : !!m[4] !== flags.U;
      items.push({ t: 'rep', body: items.pop(), min, max, lazy, possessive: q.possessive });
      // (A quantifier on a quantifier is an error: c{2}(?R)?*? is #VALUE!, measured.)
      if (flags.x) while (/\s/.test(this.peek() ?? '')) this.i++;
      if (this.quantifierAhead()) throw E.value;
      return;
    }
  }
  /** Whether a quantifier starts at i: *, +, ?, or a brace one ({2}, {1,}, {,3}; other braces are literal). */
  quantifierAhead() {
    const c = this.peek();
    if (c === '*' || c === '+' || c === '?') return true;
    const m = c === '{' && /^\{(\d*)(,(\d*))?\}/.exec(this.p.slice(this.i, this.i + 24).join(''));
    return !!m && !(m[1] === '' && (m[2] === undefined || m[3] === ''));
  }

  /** One item, after its first character `c`; null for one that adds nothing (a comment, (?i)). */
  atom(c, flags, where) {
    if (c === '\\') return this.backslash(flags, where);
    if (c === '[') return this.klass(flags);
    if (c === '(') return this.group(flags, where);
    if (c === '.') return { t: 'any', s: flags.s, nl: this.newline };
    if (c === '^') return { t: 'assert', k: flags.m ? 'mline^' : '^', nl: this.newline };
    if (c === '$') return { t: 'assert', k: flags.m ? 'mline$' : '$', nl: this.newline };
    if (c === '*' || c === '+' || c === '?') throw E.value;
    // (A brace quantifier with nothing to repeat, as after (?i) or |, is an error too: measured.)
    if (c === '{') { this.i--; const q = this.quantifierAhead(); this.i++; if (q) throw E.value; }
    return char(this.literal(c), flags.i);
  }
  literal(c) { return /^[A-Za-z0-9]$/.test(c) ? c : `\\u{${c.codePointAt(0).toString(16)}}`; }

  backslash(flags, where) {
    const c = this.peek();
    if (c === undefined) throw E.value;
    if (c === 'Q') { this.i++; return { t: 'quoted', items: this.quoted().map((s) => char(s, flags.i)) }; }
    if (c === 'E') { this.i++; return null; }
    if ('bBAzZG'.includes(c)) { this.i++; return { t: 'assert', k: c }; }
    if (c === 'K') { this.i++; if (where.look) throw E.value; return { t: 'K' }; }
    if (c === 'X') { this.i++; return { t: 'X' }; }
    if (c === 'C' || c === 'u' || c === 'U' || c === 'L' || c === 'l') throw E.value;
    if (c === 'R') { this.i++; return { t: 'atomic', body: { t: 'alt', id: ++this.alts, alts: [seq([char('\\r'), char('\\n')]), char(`[${VSPACE}]`)] } }; }
    if (c === 'N') {
      this.i++;
      if (this.peek() === '{') {
        const m = /^U\+([0-9A-Fa-f]+)$/.exec(this.braced());
        if (!m) throw E.value;
        return char(`\\u{${parseInt(m[1], 16).toString(16)}}`, flags.i);
      }
      return { t: 'any', s: false, nl: this.newline };
    }
    if (/[1-9]/.test(c)) {
      this.i++;
      const n = c + this.digits(/[0-9]/, 9);
      if (n.length === 1 || Number(n) <= this.total) return this.bref(Number(n), flags);
      this.i -= n.length;
      return char(this.escape(false), flags.i);
    }
    if (c === 'g') {
      this.i++;
      const open = this.peek();
      if (open === '<' || open === "'") return this.call(this.braced(open === '<' ? '>' : "'"));
      const ref = open === '{' ? this.braced() : this.digits(/[-+0-9]/, 9);
      if (/^-\d+$/.test(ref)) return this.bref(this.count + 1 + Number(ref), flags);
      if (/^\d+$/.test(ref)) return this.bref(Number(ref), flags);
      if (NAME.test(ref)) return this.bref(ref, flags);
      throw E.value;
    }
    if (c === 'k') {
      this.i++;
      const close = { '<': '>', "'": "'", '{': '}' }[this.peek()];
      if (!close) throw E.value;
      return this.bref(this.braced(close), flags);
    }
    // (Caseless, \p{Lt} is any cased letter, as \p{Lu} and \p{Ll} are: REGEXTEST("a","\p{Lt}",1) is TRUE.)
    const src = this.escape(false);
    return char(flags.i ? src.replace('\\p{Lt}', '\\p{LC}') : src, flags.i);
  }
  bref(ref, flags) {
    if (typeof ref === 'number' && ref < 1) throw E.value;
    this.refs.push(ref);
    return { t: 'bref', ref, i: flags.i };
  }
  /** A subroutine call by number (0 the whole pattern), relative number (+1, -1) or name. */
  call(ref) {
    let n = ref;
    if (/^[+-]\d+$/.test(ref)) n = ref[0] === '+' ? this.count + Number(ref) : this.count + 1 + Number(ref);
    else if (/^\d+$/.test(ref)) n = Number(ref);
    else if (!NAME.test(ref)) throw E.value;
    if (typeof n === 'number' && n < 0) throw E.value;
    this.refs.push(n);
    return { t: 'call', ref: n };
  }

  /** A class: one character. Caseless, [:upper:] and [:lower:] keep their case (measured: [[:upper:]] doesn't match
   *  "a" with case_sensitivity 1, though [A-Z] and \p{Lu} do). */
  klass(flags) {
    const from = this.i - 1;
    if (flags.xx) this.stripClassSpaces(from);
    this.posixSeen = [];
    const plain = flags.i ? { ...POSIX, upper: '', lower: '' } : POSIX;
    const src = this.charClass(plain);
    const kept = this.posixSeen.filter((p) => /^\[:\^?(upper|lower):\]$/.test(p));
    this.posixSeen = null;
    if (!flags.i || !kept.length) return char(src, flags.i);
    const own = kept.map((p) => (p[2] === '^' ? `[^${POSIX[p.slice(3, -2)]}]` : POSIX[p.slice(2, -2)])).join('');
    const negated = src.startsWith('[^');
    return { t: 'class2', folded: negated ? `[${src.slice(2)}` : src, kept: `[${own}]`, negated };
  }
  /** (?xx): spaces and tabs in a class are left out. */
  stripClassSpaces(from) {
    let end = from + 1;
    while (end < this.p.length && !(this.p[end] === ']' && end > from + 1)) end++;
    for (let k = end - 1; k > from; k--) if (this.p[k] === ' ' || this.p[k] === '\t') this.p.splice(k, 1);
  }

  /** A group, after its (. */
  group(flags, where) {
    const rest = this.p.slice(this.i, this.i + 60).join('');
    let m;
    if (rest.startsWith('*')) return this.verb(where);
    if (rest[0] !== '?') {
      if (flags.n) return this.inner({ t: 'group', n: null }, flags, where);
      return this.capture(null, flags, where);
    }
    if (rest.startsWith('?#')) { while (this.i < this.p.length && this.p[this.i] !== ')') this.i++; if (this.i >= this.p.length) throw E.value; this.i++; return null; }
    if ((m = /^\?(?:P?<([A-Za-z_][A-Za-z0-9_]*)>|'([A-Za-z_][A-Za-z0-9_]*)')/.exec(rest))) { this.i += m[0].length; return this.capture(m[1] || m[2], flags, where); }
    if ((m = /^\?P=([A-Za-z_][A-Za-z0-9_]*)\)/.exec(rest))) { this.i += m[0].length; return this.bref(m[1], flags); }
    if ((m = /^\?(?:P>|&)([A-Za-z_][A-Za-z0-9_]*)\)/.exec(rest))) { this.i += m[0].length; return this.call(m[1]); }
    if ((m = /^\?(R|[+-]?\d+)\)/.exec(rest))) { this.i += m[0].length; return this.call(m[1] === 'R' ? '0' : m[1]); }
    if (rest.startsWith('?(')) { this.i += 2; return this.conditional(flags, where); }
    if ((m = /^\?(=|!|<=|<!)/.exec(rest))) {
      this.i += m[0].length;
      return this.inner({ t: 'look', behind: m[1][0] === '<', neg: m[1].endsWith('!') }, flags, { ...where, look: true });
    }
    if (rest.startsWith('?>')) { this.i += 2; return this.inner({ t: 'atomic' }, flags, where); }
    if (rest.startsWith('?|')) { this.i += 2; return this.branchReset(flags, where); }
    if ((m = /^\?(\^)?([imnsxUJ]*)(?:-([imnsxUJ]*))?([:)])/.exec(rest))) {
      this.i += m[0].length;
      const set = m[4] === ')' ? flags : { ...flags };
      if (m[1]) Object.assign(set, { i: false, m: false, n: false, s: false, x: false, xx: false });
      for (const f of m[2]) { if (f === 'x' && set.x) set.xx = true; set[f] = true; }
      for (const f of m[3] || '') { set[f] = false; if (f === 'x') set.xx = false; }
      if (m[4] === ')') return { t: 'empty' };
      return this.inner({ t: 'group', n: null }, set, where);
    }
    throw E.value; // ((?C) callouts, (? alone and the rest)
  }
  capture(name, flags, where) {
    const n = ++this.count;
    if (name) {
      const had = this.names.get(name);
      if (had && !had.includes(n) && !flags.J) throw E.value;
      this.names.set(name, [...(had || []), n].filter((x, k, a) => a.indexOf(x) === k));
    }
    return this.inner({ t: 'group', n, name }, flags, where);
  }
  /** The body of a group up to its ), into `node.body`. */
  inner(node, flags, where) {
    node.body = this.alternation({ ...flags }, where);
    if (this.next() !== ')') throw E.value;
    if (node.t === 'look' && node.behind) {
      [node.wmin, node.wmax] = width(node.body);
      if (node.wmax === Infinity) throw E.value; // ((?<=a+)b: measured)
    }
    return node;
  }
  /** (?|...): each alternative numbers its groups from the same one. */
  branchReset(flags, where) {
    const base = this.count;
    let most = base;
    const alts = [];
    for (;;) {
      this.count = base;
      alts.push(this.sequence({ ...flags }, where));
      most = Math.max(most, this.count);
      if (this.peek() !== '|') break;
      this.i++;
    }
    this.count = most;
    if (this.next() !== ')') throw E.value;
    return { t: 'group', n: null, body: alts.length === 1 ? alts[0] : { t: 'alt', alts, id: ++this.alts } };
  }
  /** (*VERB) or (*VERB:name) inside a pattern. */
  verb(where) {
    const end = this.p.indexOf(')', this.i);
    if (end < 0) throw E.value;
    const body = this.p.slice(this.i + 1, end).join('');
    this.i = end + 1;
    const [name, arg] = body.split(':');
    if (name === 'MARK' || name === '') { if (!arg) throw E.value; return { t: 'empty' }; }
    const v = VERBS[name];
    if (!v || (arg !== undefined && (v === 'fail' || v === 'accept' || v === 'commit'))) throw E.value;
    if (v === 'skip' && arg !== undefined) throw E.value; // ((*SKIP:name) needs marks, not kept here)
    const node = { t: 'verb', v, alt: where.alt };
    this.verbs.push(node);
    return node;
  }
  /** (?(condition)yes|no), after the (?(. */
  conditional(flags, where) {
    const rest = this.p.slice(this.i, this.i + 60).join('');
    let cond;
    let m;
    if ((m = /^([+-]?)(\d+)\)/.exec(rest))) {
      const n = m[1] === '+' ? this.count + Number(m[2]) : m[1] === '-' ? this.count + 1 - Number(m[2]) : Number(m[2]);
      if (n < 1) throw E.value;
      this.refs.push(n);
      cond = { set: n };
    } else if ((m = /^(?:<([A-Za-z_]\w*)>|'([A-Za-z_]\w*)')\)/.exec(rest))) { cond = { set: m[1] || m[2] }; this.refs.push(cond.set); }
    else if ((m = /^R\)/.exec(rest))) cond = { R: true };
    else if ((m = /^R(\d+)\)/.exec(rest))) { cond = { R: Number(m[1]) }; this.refs.push(cond.R); }
    else if ((m = /^R&([A-Za-z_]\w*)\)/.exec(rest))) { cond = { R: m[1] }; this.refs.push(m[1]); }
    else if ((m = /^DEFINE\)/.exec(rest))) cond = { define: true };
    else if ((m = /^([A-Za-z_]\w*)\)/.exec(rest))) { cond = { set: m[1] }; this.refs.push(m[1]); }
    else if ((m = /^\?(=|!|<=|<!)/.exec(rest))) {
      this.i += m[0].length;
      cond = { look: this.inner({ t: 'look', behind: m[1][0] === '<', neg: m[1].endsWith('!') }, flags, { ...where, look: true }) };
      m = null;
    } else throw E.value;
    if (m) this.i += m[0].length;
    const body = this.alternation({ ...flags }, where);
    if (this.next() !== ')') throw E.value;
    const branches = body.t === 'alt' ? body.alts : [body];
    if (branches.length > 2 || (cond.define && branches.length > 1)) throw E.value; // ((a)(?(1)b|c|d): measured)
    return { t: 'cond', cond, yes: branches[0], no: branches[1] ?? seq([]) };
  }
}

/** The fewest and most characters a node can match (Infinity when unbounded): lookbehind must be bounded. */
export function width(node) {
  switch (node.t) {
    case 'char': case 'any': case 'class2': return [1, 1];
    case 'seq': return node.items.reduce(([a, b], x) => { const [c, d] = width(x); return [a + c, b + d]; }, [0, 0]);
    case 'alt': { const w = node.alts.map(width); return [Math.min(...w.map((x) => x[0])), Math.max(...w.map((x) => x[1]))]; }
    case 'group': case 'atomic': return width(node.body);
    case 'rep': { const [a, b] = width(node.body); return [a * node.min, node.max === Infinity ? (b === 0 ? 0 : Infinity) : b * node.max]; }
    case 'cond': { const y = width(node.yes); const n = width(node.no); return [Math.min(y[0], n[0]), Math.max(y[1], n[1])]; }
    case 'X': case 'bref': case 'call': return [0, Infinity];
    default: return [0, 0];
  }
}

/** Reads a pattern: { root, count (capture groups), names (name -> [numbers]), newline }. Throws #VALUE! for what
 *  PCRE2 refuses. */
export function parse(pattern, caseless) {
  const p = new Parser(pattern, caseless);
  const root = p.parse();
  return { root, count: p.count, names: p.names, accepts: p.verbs.some((v) => v.v === 'accept') };
}
