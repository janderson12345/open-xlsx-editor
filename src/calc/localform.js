// A formula as a person with other country settings types and reads it, and back. The file keeps formulas in en-US
// form (a comma between arguments, a dot in numbers); where the decimal separator is a comma Excel shows and takes them
// with a semicolon between arguments and a comma in numbers (=ROUND(1,5;0) for ROUND(1.5,0)), and an array constant's
// columns split by a dot ({1.2;3.4} for {1,2;3,4}). Text in quotes, quoted sheet names and [bracketed] parts are kept
// as they are. Function names stay English (Excel translates them too; not done here).
import { EN_US } from '../ooxml/locale.js';

/** Whether a locale writes formulas as en-US does. */
const plain = (loc) => !loc || loc === EN_US || (loc.listSep === ',' && loc.decimal === '.');

/** The end (index after) of a quoted run starting at i: "..." with "" inside, or '...' with ''. */
function quoted(s, i) {
  const q = s[i];
  let j = i + 1;
  while (j < s.length) {
    if (s[j] === q && s[j + 1] === q) { j += 2; continue; }
    if (s[j] === q) return j + 1;
    j++;
  }
  return j;
}

/** The end of a bracketed run starting at i ([Table1[[#This Row],[Col]]], [1]Sheet1!A1), brackets nested. */
function bracketed(s, i) {
  let depth = 0;
  for (let j = i; j < s.length; j++) {
    if (s[j] === "'" && depth > 0 && s[j + 1] !== undefined) { j++; continue; }
    if (s[j] === '[') depth++;
    else if (s[j] === ']' && --depth === 0) return j + 1;
  }
  return s.length;
}

/** Whether the character before index i makes a digit there part of a name or reference (A1, LOG10, $B$2). */
const inName = (s, i) => i > 0 && /[\p{L}\d_.$\\]/u.test(s[i - 1]);

/** Walk a formula, calling `num(text)` for each number and `sep(ch, inArray)` for each separator character outside
 *  quotes and brackets; `dec` is the decimal separator the formula's numbers use. */
function walk(s, dec, num, sep) {
  let out = '';
  let arr = 0;
  let i = 0;
  while (i < s.length) {
    const ch = s[i];
    if (ch === '"' || ch === "'") { const j = quoted(s, i); out += s.slice(i, j); i = j; continue; }
    if (ch === '[') { const j = bracketed(s, i); out += s.slice(i, j); i = j; continue; }
    if (ch === '{') arr++;
    if (ch === '}') arr = Math.max(0, arr - 1);
    const startsNum = /\d/.test(ch) || (ch === dec && /\d/.test(s[i + 1] || ''));
    if (startsNum && !inName(s, i)) {
      let j = i;
      while (/\d/.test(s[j] || '')) j++;
      if (s[j] === dec && /\d/.test(s[j + 1] || '')) { j++; while (/\d/.test(s[j] || '')) j++; } else if (s[j] === dec && j > i && !arr) j++;
      const e = /^[eE][+-]?\d+/.exec(s.slice(j));
      if (e) j += e[0].length;
      out += num(s.slice(i, j));
      i = j;
      continue;
    }
    out += sep(ch, arr > 0);
    i++;
  }
  return out;
}

/** An en-US formula (without its =, as displayFormula gives it) as `loc` shows it. */
export function toLocalFormula(f, loc) {
  if (plain(loc)) return f;
  return walk(f, '.', (n) => n.replace('.', loc.decimal), (ch, inArray) => {
    if (ch === ',') return inArray ? loc.arrayCol : loc.listSep;
    if (ch === ';' && inArray) return loc.arrayRow;
    return ch;
  });
}

/** A formula typed in `loc`'s form, as en-US writes it. */
export function fromLocalFormula(text, loc) {
  if (plain(loc)) return text;
  return walk(text, loc.decimal, (n) => n.replace(loc.decimal, '.'), (ch, inArray) => {
    if (inArray && ch === loc.arrayCol) return ',';
    if (inArray && ch === loc.arrayRow) return ';';
    if (!inArray && ch === loc.listSep) return ',';
    return ch;
  });
}
