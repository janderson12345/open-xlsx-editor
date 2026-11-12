// A formula's tokens as a tree, by Excel's precedence: comparisons, then &, + and -, * and /, ^, then %, and negation
// tightest of all (=-2^2 is 4). Operators of one level work left to right (=2^3^2 is 64).
import { tokenize, FormulaError } from './tokens.js';

const BINARY = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };

/** A function's name as the engine knows it: capitals, without the file's _xlfn. or _xlws. prefix. */
export const functionName = (text) => text.toUpperCase().replace(/^(_XL(FN|WS)\.)+/, '');

/** A formula (without "=") as a tree of { k, ... }: num, str, bool, err (`v`); ref (`sheet`, `book`, `area`); name;
 *  fn (`name`, `args`); missing (an argument left out); neg, pos, pct, paren (`a`); bin (`op`, `a`, `b`); isect
 *  (`a`, `b`: intersection); union (`items`); arr (`rows`). A 3D reference's ref has `sheets`, its first and last. Throws a FormulaError when it isn't a formula. */
export function parseFormula(text) {
  const toks = tokenize(text);
  let i = 0;
  const peek = () => toks[i];
  const is = (t, type, s) => t && t.type === type && t.text === s;
  const fail = (msg) => { throw new FormulaError(msg, toks[i] ? toks[i].start : text.length); };

  function expr(min) {
    let left = unary();
    for (;;) {
      const t = peek();
      if (!t || t.type !== 'op' || !(t.text in BINARY) || BINARY[t.text] < min) return left;
      i++;
      left = { k: 'bin', op: t.text, a: left, b: expr(BINARY[t.text] + 1) };
    }
  }
  function unary() {
    const t = peek();
    if (is(t, 'op', '-') || is(t, 'op', '+')) {
      i++;
      const a = unary();
      return { k: t.text === '-' ? 'neg' : 'pos', a };
    }
    // (@: implicit intersection, as the file's SINGLE function has it.)
    if (is(t, 'op', '@')) { i++; return { k: 'fn', name: 'SINGLE', args: [unary()] }; }
    let node = called();
    // (Intersection, a space between references: tighter than negation and %, as Excel has it.)
    while (is(peek(), 'op', ' ')) { i++; node = { k: 'isect', a: node, b: called() }; }
    while (is(peek(), 'op', '%')) { i++; node = { k: 'pct', a: node }; }
    return node;
  }
  function called() {
    let node = primary();
    // (A LAMBDA called where it is written: =LAMBDA(x,x*2)(3).)
    while (is(peek(), 'punct', '(') && (node.k === 'call' || (node.k === 'fn' && node.name === 'LAMBDA'))) node = { k: 'call', f: node, args: argList() };
    return node;
  }
  /** A bracketed list of arguments, from its "(": each an expression, or missing (left out). */
  function argList() {
    i++; // (the "(")
    const args = [];
    if (is(peek(), 'punct', ')')) { i++; return args; }
    for (;;) {
      args.push(is(peek(), 'punct', ',') || is(peek(), 'punct', ')') ? { k: 'missing' } : expr(1));
      const sep = toks[i++];
      if (is(sep, 'punct', ')')) return args;
      if (!is(sep, 'punct', ',')) { i--; return fail('Expected , or ) after an argument'); }
    }
  }
  function constant() {
    let t = toks[i++];
    let sign = 1;
    if (is(t, 'op', '-')) { sign = -1; t = toks[i++]; }
    if (t && t.type === 'num') return t.value * sign;
    if (sign === 1 && t && (t.type === 'str' || t.type === 'bool')) return t.value;
    if (sign === 1 && t && t.type === 'err') return { err: t.value };
    i--;
    return fail('An array can hold only numbers, text, TRUE, FALSE and errors');
  }
  function primary() {
    const t = toks[i++];
    if (!t) { i--; return fail('The formula ends too soon'); }
    switch (t.type) {
      // (`src`: a number and a reference as typed, for Evaluate Formula to show: stepper.js.)
      case 'num': return { k: 'num', v: t.value, src: t.text };
      case 'str': case 'bool': return { k: t.type, v: t.value };
      case 'err': return { k: 'err', v: t.value };
      case 'ref': {
        if (!t.area) return { k: 'err', v: '#REF!' };
        const ref = { k: 'ref', sheet: t.sheet, book: t.book, area: t.area, src: t.text };
        if (t.sheets) ref.sheets = t.sheets;
        return t.spill ? { k: 'fn', name: 'ANCHORARRAY', args: [ref] } : ref;
      }
      case 'name': return { k: 'name', name: t.text };
      case 'struct': return { k: 'struct', table: t.table, spec: t.spec, src: t.text };
      case 'func': return { k: 'fn', name: functionName(t.text), args: argList() };
      case 'punct':
        if (t.text === '(') {
          const items = [expr(1)];
          while (is(peek(), 'punct', ',')) { i++; items.push(expr(1)); }
          if (!is(toks[i++], 'punct', ')')) { i--; return fail('Expected )'); }
          return items.length === 1 ? { k: 'paren', a: items[0] } : { k: 'union', items };
        }
        if (t.text === '{') {
          const rows = [[constant()]];
          for (;;) {
            const sep = toks[i++];
            if (is(sep, 'punct', '}')) break;
            if (is(sep, 'punct', ',')) rows[rows.length - 1].push(constant());
            else if (is(sep, 'punct', ';')) rows.push([constant()]);
            else { i--; return fail('Expected , ; or } in an array'); }
          }
          if (rows.some((r) => r.length !== rows[0].length)) return fail('The rows of an array must be as long as each other');
          return { k: 'arr', rows };
        }
    }
    i--;
    return fail(`Unexpected "${t.text}"`);
  }

  if (!toks.length) fail('The formula is empty');
  const tree = expr(1);
  if (i < toks.length) fail(`Unexpected "${toks[i].text}"`);
  return tree;
}
