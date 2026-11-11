// REGEXTEST, REGEXEXTRACT and REGEXREPLACE, with Excel's regular expressions (regex.js) and its replacement text
// (PCRE2's extended substitution): $1, ${1}, ${name}, $0 and $& the match, $$ a dollar; \1 a group too; \n, \t, \x41;
// \u and \l the next character's case, \U and \L until \E; ${1:-default} and ${1:+set:unset}. A group that doesn't
// exist, "$" alone, "\" alone or an unknown escape (\q) is #VALUE!. Measured against Excel with Evaluate and the
// formula probe. Registered into functions.js's table; evaluate.js imports this file.
import { Arr, E } from './values.js';
import { lifted, def, given, num, text } from './functions.js';
import { regex } from './regex.js';

const NAME = /^[A-Za-z_][A-Za-z0-9_]*/;

/** Replacement text as parts: { t } text, { g } a group (number or name), { c } a case change, { if, set, unset }.
 *  `stop` the characters that end it (inside ${n:+set:unset}); returns [parts, where it stopped]. */
function parseReplacement(s, from = 0, stop = '') {
  const parts = [];
  let i = from;
  const text1 = (t) => { if (parts.length && parts[parts.length - 1].t !== undefined) parts[parts.length - 1].t += t; else parts.push({ t }); };
  while (i < s.length) {
    const c = s[i];
    if (stop.includes(c)) return [parts, i];
    i++;
    if (c === '$') {
      const n = s[i];
      if (n === '$') { text1('$'); i++; continue; }
      if (n === '&') { parts.push({ g: 0 }); i++; continue; }
      if (/[0-9]/.test(n ?? '')) { const m = /^\d+/.exec(s.slice(i)); parts.push({ g: Number(m[0]) }); i += m[0].length; continue; }
      if (n === '{') {
        const m = /^\{(\d+|[A-Za-z_][A-Za-z0-9_]*)/.exec(s.slice(i));
        if (!m) throw E.value;
        i += m[0].length;
        const g = /^\d/.test(m[1]) ? Number(m[1]) : m[1];
        if (s[i] === '}') { parts.push({ g }); i++; continue; }
        if (s.startsWith(':-', i)) {
          const [def0, at] = parseReplacement(s, i + 2, '}');
          if (s[at] !== '}') throw E.value;
          parts.push({ if: g, set: [{ g }], unset: def0 });
          i = at + 1;
          continue;
        }
        if (s.startsWith(':+', i)) {
          const [set, at] = parseReplacement(s, i + 2, ':}');
          let unset = [];
          let end = at;
          if (s[at] === ':') [unset, end] = parseReplacement(s, at + 1, '}');
          if (s[end] !== '}') throw E.value;
          parts.push({ if: g, set, unset });
          i = end + 1;
          continue;
        }
        throw E.value;
      }
      const m = NAME.exec(s.slice(i));
      if (!m) throw E.value;
      parts.push({ g: m[0] });
      i += m[0].length;
      continue;
    }
    if (c === '\\') {
      const n = s[i++];
      if (n === undefined) throw E.value;
      if (/[1-9]/.test(n)) { parts.push({ g: Number(n) }); continue; }
      const simple = { n: '\n', t: '\t', r: '\r', f: '\f', a: '\x07', e: '\x1b' }[n];
      if (simple) { text1(simple); continue; }
      if ('uUlLE'.includes(n)) { parts.push({ c: n }); continue; }
      if (n === 'x') {
        let hex;
        if (s[i] === '{') { const end = s.indexOf('}', i); hex = s.slice(i + 1, end); i = end + 1; } else { hex = /^[0-9A-Fa-f]{0,2}/.exec(s.slice(i))[0]; i += hex.length; }
        text1(String.fromCodePoint(parseInt(hex || '0', 16)));
        continue;
      }
      if (n === 'o' && s[i] === '{') { const end = s.indexOf('}', i); text1(String.fromCodePoint(parseInt(s.slice(i + 1, end), 8))); i = end + 1; continue; }
      if (/[A-Za-z0-9]/.test(n)) throw E.value;
      text1(n);
      continue;
    }
    text1(c);
  }
  return [parts, i];
}

/** The replacement for one match: `group(g)` a group's text, or undefined where it didn't take part. */
function replacement(parts, group) {
  let out = '';
  let mode = '';
  let next = '';
  const add = (t) => {
    for (const ch of t) {
      let x = mode === 'U' ? ch.toUpperCase() : mode === 'L' ? ch.toLowerCase() : ch;
      if (next) { x = next === 'u' ? x.toUpperCase() : x.toLowerCase(); next = ''; }
      out += x;
    }
  };
  const run = (ps) => {
    for (const p of ps) {
      if (p.t !== undefined) add(p.t);
      else if (p.g !== undefined) add(group(p.g) ?? '');
      else if (p.c) { if (p.c === 'u' || p.c === 'l') next = p.c; else mode = p.c === 'E' ? '' : p.c; }
      else run(group(p.if) !== undefined ? p.set : p.unset);
    }
  };
  run(parts);
  return out;
}

/** The groups a replacement names, each one the pattern must have. */
function checkGroups(parts, groups) {
  for (const p of parts) {
    const g = p.g ?? p.if;
    if (g !== undefined && g !== 0 && !(typeof g === 'number' ? g <= groups.length : groups.includes(g))) throw E.value;
    if (p.set) { checkGroups(p.set, groups); checkGroups(p.unset, groups); }
  }
}

/** case_sensitivity: 0 (the default) or 1, else #VALUE!. */
function caseless(a, k, ctx) {
  if (!given(a[k])) return false;
  const n = Math.trunc(num(a[k], ctx));
  if (n !== 0 && n !== 1) throw E.value;
  return n === 1;
}
const matches = (re, s) => [...s.matchAll(re)];

lifted('REGEXTEST', (a, ctx) => {
  const { re } = regex(text(a[1], ctx), caseless(a, 2, ctx));
  re.lastIndex = 0;
  return re.test(text(a[0], ctx));
});

// (REGEXEXTRACT isn't worked out for each value of an array: =REGEXEXTRACT({"a1","b2"},"\d") is "1".)
def('REGEXEXTRACT', (a, ctx) => {
  const s = text(a[0], ctx);
  const mode = given(a[2]) ? Math.trunc(num(a[2], ctx)) : 0;
  if (mode < 0 || mode > 2) throw E.value;
  const { re, groups } = regex(text(a[1], ctx), caseless(a, 3, ctx));
  const all = matches(re, s);
  if (mode === 2) {
    // (The groups of the first match, in a row; one that took no part #N/A. No match, or no groups: #VALUE!.)
    if (!all.length || !groups.length) throw E.value;
    return new Arr([groups.map((g) => all[0].groups[g] ?? E.na)]);
  }
  if (!all.length) throw E.na;
  return mode === 0 ? all[0][0] : new Arr([all.map((m) => m[0])]);
});

lifted('REGEXREPLACE', (a, ctx) => {
  const s = text(a[0], ctx);
  const { re, groups } = regex(text(a[1], ctx), caseless(a, 4, ctx));
  const [parts] = parseReplacement(text(a[2], ctx));
  checkGroups(parts, groups);
  const occurrence = given(a[3]) ? Math.trunc(num(a[3], ctx)) : 0;
  const all = matches(re, s);
  const which = occurrence === 0 ? null : occurrence > 0 ? occurrence - 1 : all.length + occurrence;
  let out = '';
  let at = 0;
  all.forEach((m, k) => {
    if (which !== null && k !== which) return;
    const group = (g) => (g === 0 ? m[0] : m.groups[typeof g === 'number' ? groups[g - 1] : g]);
    out += s.slice(at, m.index) + replacement(parts, group);
    at = m.index + m[0].length;
  });
  return out + s.slice(at);
});
