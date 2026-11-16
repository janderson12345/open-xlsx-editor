// A backtracking matcher for the tree pcre.js reads, with PCRE2's semantics where Excel was seen to follow them
// (docs/CALC.md) and a limit on the work one match may take: past STEPS steps, or nested too deep, it is #VALUE!, as
// Excel gives at PCRE2's match limit (=REGEXTEST(REPT("a",30)&"b","^(a+)+$") is #VALUE! in Excel, worked out in a few
// milliseconds here). So no pattern, however it backtracks, holds up the sheet: the guard against catastrophic
// backtracking on patterns a workbook brings. Each one-character test is a JavaScript regular expression of one class,
// which can't backtrack. Positions are UTF-16 indexes; a character is a code point.
import { E } from './values.js';
import { parse } from './pcre.js';
import { WORD } from './regex.js';

/** The most steps one match (one try from each start) may take; PCRE2's own default match limit is 10,000,000. */
export const STEPS = 5000000;
/** The most steps all the searches of one REGEXEXTRACT or REGEXREPLACE call (Pattern.all) may take together. */
export const TOTAL_STEPS = 4 * STEPS;
const DEPTH = 2000; // (recursions in progress at once)
const IS_WORD = new RegExp(`[${WORD}]`, 'v');
const NEWLINES = { LF: ['\n'], CR: ['\r'], CRLF: ['\r\n'], ANYCRLF: ['\r\n', '\r', '\n'], NUL: ['\0'],
  ANY: ['\r\n', '\r', '\n', '\v', '\f', '\u0085', ' ', ' '] };
const segmenter = typeof Intl !== 'undefined' && Intl.Segmenter ? new Intl.Segmenter('en', { granularity: 'grapheme' }) : null;
const LIMIT = { limit: true };

const cpLen = (s, p) => (p < s.length && s.charCodeAt(p) >= 0xd800 && s.charCodeAt(p) <= 0xdbff && p + 1 < s.length ? 2 : 1);
const cpBack = (s, p) => (p >= 2 && s.charCodeAt(p - 1) >= 0xdc00 && s.charCodeAt(p - 1) <= 0xdfff && s.charCodeAt(p - 2) >= 0xd800 && s.charCodeAt(p - 2) <= 0xdbff ? 2 : 1);
const fold = (x) => x.toLowerCase();

/** A compiled pattern: exec(s, from, notEmptyAtStart) gives { index, start, end, caps } or null. */
export class Pattern {
  constructor(pattern, caseless) {
    const { root, count, names, accepts } = parse(pattern, caseless);
    this.count = count;
    this.names = names;
    this.S = { s: '', caps: [], steps: 0, kpos: -1, start0: 0, active: [], accept: [], open: [] };
    this.groups = [];
    this.root = this.compile(root);
    this.required = accepts ? null : required(root);
  }

  /** A sticky one-character test for a class's JavaScript source: the length it matches at p, or 0. */
  tester(src, caseless) {
    let re;
    try { re = new RegExp(src, caseless ? 'vyi' : 'vy'); } catch { throw E.value; } // (\p{IsLatin}: #VALUE!, measured)
    return (s, p) => { re.lastIndex = p; return re.test(s) ? re.lastIndex - p : 0; };
  }

  compile(node) {
    const S = this.S;
    const tick = () => { if (++S.steps > STEPS) throw LIMIT; };
    switch (node.t) {
      case 'char': case 'class2': case 'any': {
        const test = this.charTest(node);
        const f = (p, k) => { tick(); const n = test(S.s, p); return n > 0 && k(p + n); };
        f.test = test;
        return f;
      }
      case 'seq': {
        const items = node.items.map((x) => this.compile(x));
        const f = items.reduceRight((next, item) => (p, k) => item(p, (q) => next(q, k)), (p, k) => k(p));
        // (A run of single characters matches one way only: its length at p, or 0, for repeat's loop.)
        if (items.length && items.every((x) => x.test)) {
          const tests = items.map((x) => x.test);
          f.test = (s, p) => { let q = p; for (const t of tests) { tick(); const n = t(s, q); if (!n) return 0; q += n; } return q - p; };
        }
        return f;
      }
      case 'alt': {
        const alts = node.alts.map((x) => this.compile(x));
        return (p, k) => {
          for (const a of alts) {
            try { if (a(p, k)) return true; } catch (e) { if (e?.then !== node.id) throw e; }
          }
          return false;
        };
      }
      case 'group': return this.group(node);
      case 'rep': return this.repeat(node, tick);
      case 'atomic': {
        const body = this.compile(node.body);
        return (p, k) => {
          const snap = S.caps.slice();
          let end = -1;
          if (!body(p, (q) => { end = q; return true; })) return false;
          if (k(end)) return true;
          S.caps = snap;
          return false;
        };
      }
      case 'look': return this.look(node);
      case 'bref': return this.bref(node, tick);
      case 'assert': return this.assertion(node);
      case 'K': return (p, k) => { const was = S.kpos; S.kpos = p; if (k(p)) return true; S.kpos = was; return false; };
      case 'X': return (p, k) => {
        tick();
        if (p >= S.s.length) return false;
        const g = segmenter ? segmenter.segment(S.s.slice(p, p + 64))[Symbol.iterator]().next().value.segment.length : cpLen(S.s, p);
        return k(p + g);
      };
      case 'call': return this.call(node);
      case 'cond': return this.conditional(node);
      case 'verb': return this.verb(node);
      case 'empty': return (p, k) => k(p);
      default: throw E.value;
    }
  }

  charTest(node) {
    if (node.t === 'char') return this.tester(node.src, node.i);
    if (node.t === 'class2') {
      const a = this.tester(node.folded, true);
      const b = this.tester(node.kept, false);
      return (s, p) => {
        if (p >= s.length) return 0;
        const hit = a(s, p) || b(s, p);
        return node.negated ? (hit ? 0 : cpLen(s, p)) : hit;
      };
    }
    const nl = NEWLINES[node.nl] || NEWLINES.LF;
    return (s, p) => (p >= s.length || (!node.s && nl.some((x) => s.startsWith(x, p))) ? 0 : cpLen(s, p));
  }

  group(node) {
    const S = this.S;
    const body = this.compile(node.body);
    if (node.n === null) return body;
    const n = node.n;
    this.groups[n] = body;
    return (p, k) => {
      const entry = { n, p };
      S.open.push(entry);
      const r = body(p, (q) => {
        const was = [S.caps[2 * n], S.caps[2 * n + 1]];
        const at = S.open.lastIndexOf(entry);
        if (at >= 0) S.open.splice(at, 1);
        S.caps[2 * n] = p; S.caps[2 * n + 1] = q;
        if (k(q)) return true;
        [S.caps[2 * n], S.caps[2 * n + 1]] = was;
        S.open.push(entry);
        return false;
      });
      const at = S.open.lastIndexOf(entry);
      if (at >= 0) S.open.splice(at, 1);
      return r;
    };
  }

  repeat(node, tick) {
    const S = this.S;
    const body = this.compile(node.body);
    const { min, max, lazy } = node;
    if (body.test) {
      // (A character, or a run of them, a time, as a loop: no nesting however long the run.)
      const test = body.test;
      const run = (p, k) => {
        const at = [p];
        let q = p;
        while (at.length - 1 < max) { tick(); const n = test(S.s, q); if (!n) break; q += n; at.push(q); if (lazy && at.length - 1 >= min) break; }
        if (at.length - 1 < min) return false;
        if (node.possessive) return k(at[at.length - 1]);
        if (!lazy) { for (let c = at.length - 1; c >= min; c--) if (k(at[c])) return true; return false; }
        // (Lazy: the fewest first, one more each time the rest fails.)
        let c = min;
        q = at[c];
        for (;;) {
          if (k(q)) return true;
          if (c >= max) return false;
          tick();
          const n = test(S.s, q);
          if (!n) return false;
          q += n; c++;
        }
      };
      return run;
    }
    const loop = (count, p, k) => {
      tick();
      const more = () => count < max && body(p, (q) => (q === p ? k(q) : loop(count + 1, q, k)));
      if (lazy) return (count >= min && k(p)) || more();
      return more() || (count >= min && k(p));
    };
    if (!node.possessive) return (p, k) => loop(0, p, k);
    return (p, k) => {
      const snap = S.caps.slice();
      let end = -1;
      if (!loop(0, p, (q) => { end = q; return true; })) return false;
      if (k(end)) return true;
      S.caps = snap;
      return false;
    };
  }

  /** A lookaround: atomic, its captures kept when it holds; a verb inside it only ends it. */
  look(node) {
    const S = this.S;
    const body = this.compile(node.body);
    const holds = (p) => {
      if (!node.behind) return this.inside(() => body(p, () => true));
      // (From the fewest characters back to the most the lookbehind can take.)
      let q = p;
      for (let back = 0; back < node.wmin; back++) { if (q === 0) return false; q -= cpBack(S.s, q); }
      for (let back = node.wmin; back <= node.wmax; back++) {
        const at = q;
        if (this.inside(() => body(at, (e) => e === p))) return true;
        if (q === 0) return false;
        q -= cpBack(S.s, q);
      }
      return false;
    };
    return (p, k) => {
      const snap = S.caps.slice();
      const ok = holds(p);
      if (node.neg) { S.caps = snap; return !ok && k(p); }
      if (!ok) { S.caps = snap; return false; }
      if (k(p)) return true;
      S.caps = snap;
      return false;
    };
  }
  /** Runs `f` as an assertion or recursion runs: (*ACCEPT) ends it, other verbs fail it. */
  inside(f) {
    const S = this.S;
    S.accept.push({ to: () => true, open: S.open.length });
    try { return f(); } catch (e) { if (e === LIMIT || e instanceof Error || e?.code) throw e; return false; } finally { S.accept.pop(); }
  }

  bref(node, tick) {
    const S = this.S;
    return (p, k) => {
      tick();
      const ns = typeof node.ref === 'number' ? [node.ref] : this.names.get(node.ref);
      const n = ns.find((x) => S.caps[2 * x] >= 0 && S.caps[2 * x] !== undefined);
      if (n === undefined) return false; // (a group that took no part matches nothing, as in PCRE2)
      const t = S.s.slice(S.caps[2 * n], S.caps[2 * n + 1]);
      const here = S.s.slice(p, p + t.length);
      if (node.i ? fold(here) !== fold(t) : here !== t) return false;
      return k(p + t.length);
    };
  }

  assertion(node) {
    const S = this.S;
    const nl = NEWLINES[node.nl] || NEWLINES.LF;
    const nlAt = (p) => nl.find((x) => S.s.startsWith(x, p))?.length ?? 0;
    const nlBefore = (p) => nl.some((x) => p >= x.length && S.s.startsWith(x, p - x.length));
    const wordAt = (p) => p >= 0 && p < S.s.length && IS_WORD.test(String.fromCodePoint(S.s.codePointAt(p)));
    const wordBefore = (p) => p > 0 && wordAt(p - cpBack(S.s, p));
    const tests = {
      '^': (p) => p === 0, A: (p) => p === 0, z: (p) => p === S.s.length, G: (p) => p === S.start0,
      $: (p) => p === S.s.length || (nlAt(p) > 0 && p + nlAt(p) === S.s.length), Z: (p) => p === S.s.length || (nlAt(p) > 0 && p + nlAt(p) === S.s.length),
      'mline^': (p) => p === 0 || (p < S.s.length && nlBefore(p)), 'mline$': (p) => p === S.s.length || nlAt(p) > 0,
      b: (p) => wordBefore(p) !== wordAt(p), B: (p) => wordBefore(p) === wordAt(p),
    };
    const t = tests[node.k];
    return (p, k) => t(p) && k(p);
  }

  /** (?R), (?1), (?&name): the group's body again here; its captures put back as they were afterwards. */
  call(node) {
    const S = this.S;
    return (p, k) => {
      const n = typeof node.ref === 'number' ? node.ref : this.names.get(node.ref)[0];
      const body = n === 0 ? this.root : this.groups[n];
      if (S.active.some((a) => a.n === n && a.p === p)) throw E.value; // (a recursion that would loop: (?R) alone)
      if (S.active.length >= DEPTH) throw LIMIT;
      const snap = S.caps.slice();
      const entry = { n, p };
      S.active.push(entry);
      let r;
      try {
        r = this.callBody(body, p, (q) => {
          const inner = S.caps;
          S.active.pop();
          S.caps = snap.slice();
          if (k(q)) return true;
          S.caps = inner;
          S.active.push(entry);
          return false;
        });
      } finally { if (S.active[S.active.length - 1] === entry) S.active.pop(); }
      if (!r) S.caps = snap;
      return r;
    };
  }
  callBody(body, p, k) {
    const S = this.S;
    S.accept.push({ to: k, open: S.open.length });
    try { return body(p, k); } finally { S.accept.pop(); }
  }

  conditional(node) {
    const S = this.S;
    const yes = this.compile(node.yes);
    const no = this.compile(node.no);
    const { cond } = node;
    const look = cond.look ? this.compile(cond.look) : null;
    const isSet = (n) => S.caps[2 * n] >= 0 && S.caps[2 * n] !== undefined;
    return (p, k) => {
      let truth;
      if (cond.define) truth = false;
      else if (cond.set !== undefined) truth = (typeof cond.set === 'number' ? [cond.set] : this.names.get(cond.set)).some(isSet);
      else if (cond.R === true) truth = S.active.length > 0;
      else if (cond.R !== undefined) {
        const n = typeof cond.R === 'number' ? cond.R : this.names.get(cond.R)[0];
        truth = S.active.length > 0 && S.active[S.active.length - 1].n === n;
      } else {
        let held = false;
        if (look(p, () => { held = true; return true; })) truth = held; else truth = false;
      }
      if (cond.define) return k(p);
      return (truth ? yes : no)(p, k);
    };
  }

  verb(node) {
    const S = this.S;
    switch (node.v) {
      case 'fail': return () => false;
      case 'accept': return (p) => {
        const target = S.accept[S.accept.length - 1];
        for (const o of S.open.slice(target.open)) { S.caps[2 * o.n] = o.p; S.caps[2 * o.n + 1] = p; }
        return target.to(p);
      };
      case 'then': return (p, k) => { if (k(p)) return true; throw { then: node.alt ?? null }; };
      default: return (p, k) => { if (k(p)) return true; throw { verb: node.v, p }; };
    }
  }

  /** The first match from `from` on: { index (where the try began), start (after \K), end, caps } or null. */
  exec(s, from, notEmptyAtStart = false) {
    const S = this.S;
    if (this.required && !s.includes(this.required)) return null;
    S.s = s;
    S.start0 = from;
    S.steps = 0;
    for (let p = from; p <= s.length;) {
      S.caps = new Array(2 * (this.count + 1)).fill(-1);
      S.kpos = -1; S.active = []; S.open = [];
      let end = -1;
      const final = (q) => {
        const st = S.kpos >= 0 ? S.kpos : p;
        if (notEmptyAtStart && st === q && st === from) return false;
        end = q;
        return true;
      };
      S.accept = [{ to: final, open: 0 }];
      let next = p + cpLen(s, p);
      try {
        if (this.root(p, final)) return { index: p, start: S.kpos >= 0 ? S.kpos : p, end, caps: S.caps.slice() };
      } catch (e) {
        if (e === LIMIT) throw E.value;
        if (e instanceof RangeError) throw E.value; // (nested too deep for the stack: as at PCRE2's depth limit)
        if (e?.verb === 'commit') return null;
        if (e?.verb === 'skip' && e.p > p) next = e.p;
        else if (!(e?.verb || e?.then !== undefined)) throw e;
      }
      if (p >= s.length) break;
      p = next;
    }
    return null;
  }

  /** Every match, as Excel's REGEXEXTRACT and REGEXREPLACE find them: each search from the last one's end, and after
   *  an empty match first a non-empty one at the same place. */
  all(s) {
    const out = [];
    let from = 0;
    let notEmpty = false;
    // (STEPS bounds one search; a pattern just under it on each of thousands of matches would still hold the sheet up,
    // so all the searches of one call share TOTAL_STEPS, past which it is #VALUE! as one search past STEPS is.)
    let total = 0;
    const spent = () => { total += this.S.steps; if (total > TOTAL_STEPS) throw E.value; };
    while (from <= s.length) {
      // (After an empty match: a search from there on that can't be empty where it starts, and if that finds nothing,
      // a new one a character on, where \G holds again: =REGEXREPLACE("abc","\G","-") is "-a-b-c-" and
      // =REGEXREPLACE("abc","(?!\G)","-") "a-b-c-", measured.)
      let m = this.exec(s, from, notEmpty);
      spent();
      if (!m && notEmpty) {
        if (from >= s.length) break;
        from += cpLen(s, from);
        notEmpty = false;
        m = this.exec(s, from);
        spent();
      }
      if (!m) break;
      out.push(m);
      notEmpty = m.end === m.start;
      if (m.end < from) break;
      from = m.end;
      if (out.length > 2 * s.length + 2) break;
    }
    return out;
  }
}

/** A character every match must have, when the pattern's top level has a plain one (as PCRE2's "last required code
 *  unit": =REGEXEXTRACT(REPT("a",25)&"x","(a*)*b") is #N/A at once in Excel, and so here). */
function required(root) {
  const items = root.t === 'seq' ? root.items : [root];
  for (const x of items) {
    if (x.t === 'char' && !x.i && /^[A-Za-z0-9]$/.test(x.src)) return x.src;
    if (x.t === 'char' && !x.i && /^\\u\{[0-9a-f]+\}$/.test(x.src)) return String.fromCodePoint(parseInt(x.src.slice(3, -1), 16));
  }
  return null;
}

const cache = new Map();
/** A compiled pattern from Excel's text, cached; `caseless` for case_sensitivity 1. */
export function pattern(text, caseless) {
  const key = `${caseless ? 1 : 0}${text}`;
  let p = cache.get(key);
  if (!p) {
    p = new Pattern(text, caseless);
    if (cache.size > 200) cache.clear();
    cache.set(key, p);
  }
  return p;
}
