// The Function Arguments box's model, apart from the page: the boxes a function has (its arguments as calc/catalog.js
// names them, a repeated one adding a box as the last is filled), the formula written from what is typed in them, and
// the call a formula's caret is in, for the box to open on.
//
// What Excel takes, measured over COM in a hidden Excel (tools/excel-compare/excel-fnargs.ps1, test/fnargs-excel.json:
// 47 formulas; excel-fnmin.ps1, test/fnmin-excel.json: every listed function):
// - An argument left empty is kept as an empty place: =SUM(A1:A3,,5), =IF(A1>1,,"no"), =VLOOKUP(,,), =PMT(0.01,12,1000,,)
//   are all taken and kept as typed.
// - Too few arguments are refused, empty places counting: =VLOOKUP(2,A1:A3) and =MID(B1,2) are refused, =VLOOKUP(,,)
//   and =MID(B1,2,) taken; =SUM(), =LEFT(), =IF() and =AVERAGE() are refused.
// - The fewest arguments each function takes are the ones its syntax doesn't bracket, for 491 of the 499; the other
//   eight are MIN_ARGS. So the box writes the arguments up to the last one filled, and at least that many, empty ones
//   left as empty places.
import { functionInfo } from './catalog.js';

/** The functions whose fewest arguments, as Excel takes them, aren't their syntax's unbracketed ones (fnmin-excel.json). */
const MIN_ARGS = { IF: 2, DCOUNT: 3, DCOUNTA: 3, PROB: 3, REDUCE: 3, SCAN: 3, BYCOL: 1, BYROW: 1 };
/** Excel takes at most 255 arguments. */
const MOST = 255;

/** A function's arguments: `fixed` [{ name, optional }] as the syntax names them, and `group` the ones that repeat
 *  (the last one, or the last two when both end in the same number: IFS's, COUNTIFS's pairs), or null. */
export function argSlots(info) {
  const fixed = [];
  let depth = 0;
  let dots = -1;
  for (const raw of info.args) {
    if (raw.replace(/[[\]]/g, '') === '...') { dots = fixed.length; depth -= (raw.match(/\]/g) || []).length; continue; }
    const opens = raw.startsWith('[');
    if (opens) depth++;
    fixed.push({ name: raw.replace(/[[\]]/g, ''), optional: depth > 0 });
    depth -= (raw.match(/\]/g) || []).length;
  }
  // (Repeats only at the end: LAMBDA's parameters before its calculation are boxes as named.)
  if (dots !== fixed.length || !fixed.length) return { fixed, group: null };
  const last = fixed.length - 1;
  const digit = (n) => /(\d+)$/.exec(n)?.[1];
  const size = last > 0 && digit(fixed[last].name) && digit(fixed[last].name) === digit(fixed[last - 1].name) ? 2 : 1;
  return { fixed, group: fixed.slice(fixed.length - size).map((a) => a.name) };
}

/** The fewest arguments Excel takes for the function. */
export function minArgs(info) {
  if (MIN_ARGS[info.name] != null) return MIN_ARGS[info.name];
  return argSlots(info).fixed.filter((a) => !a.optional).length;
}

/** The boxes for `values` typed so far: [{ name, optional }], the repeated ones numbered on (number3, number4 ...),
 *  one more group shown once the last is filled, as Excel adds a box. */
export function argBoxes(info, values = []) {
  const { fixed, group } = argSlots(info);
  const boxes = fixed.slice();
  if (!group) return boxes;
  const filled = (from) => values.slice(from, from + group.length).some((v) => String(v ?? '').trim() !== '');
  // (Named on from the group's number: number2 then number3; criteria_range1, criteria1 then 2.)
  // (Arguments read back from a formula each have a box, however many there are.)
  for (let k = 1; boxes.length + group.length <= MOST && (boxes.length < values.length || filled(boxes.length - group.length)); k++) {
    for (const name of group) {
      const n = /(\d+)$/.exec(name);
      boxes.push({ name: n ? name.slice(0, n.index) + (Number(n[1]) + k) : `${name}${k + 1}`, optional: true });
    }
  }
  return boxes;
}

/** The call written from the boxes' `values`: NAME(a,b,...), up to the last one filled and at least the fewest Excel
 *  takes, empty ones kept as empty places. `refused` when Excel would refuse it (too few: =SUM() with nothing). */
export function buildCall(name, values) {
  const info = functionInfo(name);
  const vals = values.map((v) => String(v ?? '').trim());
  let last = vals.length;
  while (last > 0 && vals[last - 1] === '') last--;
  const min = info ? minArgs(info) : 0;
  const count = Math.max(last, min);
  const text = `${info ? info.name : name}(${vals.slice(0, count).concat(Array(Math.max(0, count - vals.length)).fill('')).join(',')})`;
  // (One empty place can't be written: F() is no arguments. With a minimum of two or more, the commas carry them.)
  return { text, refused: min > 0 && last === 0 && min === 1 };
}

/** `text`'s arguments split at its top-level commas: quotes, brackets and array constants passed over. */
export function splitArgs(text) {
  const out = [];
  let depth = 0;
  let from = 0;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"' || ch === "'") {
      let j = i + 1;
      while (j < text.length && !(text[j] === ch && text[j + 1] !== ch)) j += text[j] === ch ? 2 : 1;
      i = j;
    } else if (ch === '(' || ch === '{' || ch === '[') depth++;
    else if (ch === ')' || ch === '}' || ch === ']') depth--;
    else if (ch === ',' && depth === 0) { out.push(text.slice(from, i)); from = i + 1; }
  }
  out.push(text.slice(from));
  return out.length === 1 && out[0].trim() === '' ? [] : out;
}

/** The end of the call whose "(" is at `open` in `text`: the index after its ")", or the text's end if it isn't closed. */
function closeOf(text, open) {
  let depth = 0;
  for (let i = open; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"' || ch === "'") {
      let j = i + 1;
      while (j < text.length && !(text[j] === ch && text[j + 1] !== ch)) j += text[j] === ch ? 2 : 1;
      i = j;
    } else if (ch === '(') depth++;
    else if (ch === ')' && --depth === 0) return { end: i + 1, inner: text.slice(open + 1, i) };
  }
  return { end: text.length, inner: text.slice(open + 1) };
}

/** The listed function a formula's caret is on, for Function Arguments to open on: a name just typed before the caret
 *  (=SUM or =SUM(, what Ctrl+A takes in Excel), or else (`inside`) the innermost listed call the caret is inside.
 *  { name, start, end, args }: the part of `text` from `start` to `end` is the call (name, brackets and arguments),
 *  `args` its arguments; null when there is none. */
export function callForCaret(text, caret, { inside = true } = {}) {
  if (!String(text).startsWith('=')) return null;
  // (The caret inside a name takes the whole name.)
  while (caret < text.length && /[A-Za-z0-9_.]/.test(text[caret]) && /[A-Za-z0-9_.]/.test(text[caret - 1] || '')) caret++;
  const word = /([A-Za-z_][A-Za-z0-9_.]*)\(?$/.exec(text.slice(0, caret));
  const at = (start, name) => {
    const open = start + name.length;
    if (text[open] !== '(') return { name: name.toUpperCase(), start, end: open, args: [] };
    const { end, inner } = closeOf(text, open);
    return { name: name.toUpperCase(), start, end, args: splitArgs(inner) };
  };
  if (word && functionInfo(word[1])) return at(caret - word[0].length, word[1]);
  if (!inside) return null;
  // (Inside a call: the innermost listed one whose brackets hold the caret.)
  const re = /([A-Za-z_][A-Za-z0-9_.]*)\(/g;
  let best = null;
  for (let m; (m = re.exec(text)) && m.index < caret;) {
    if (!functionInfo(m[1])) continue;
    const c = at(m.index, m[1]);
    if (caret > m.index + m[1].length && caret <= c.end) best = c;
  }
  return best;
}
