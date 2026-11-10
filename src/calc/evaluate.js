// Working out a formula's tree: operators, references, names and function calls. A formula that isn't an array formula
// works on single values: a range given to an operator is cut down to the formula's own row or column (implicit
// intersection). The arguments of SUMPRODUCT and its kind, and array formulas, work on whole arrays, element by element.
import { Area, Arr, Union, MISSING, XlError, E, errorOf, Unsupported, isError, scalar, matrix, toNum, toText, compare, finite,
  closeToZero } from './values.js';
import { FUNCTIONS } from './functions.js';
// (More functions, registered into FUNCTIONS once it is made.)
import './functions-more.js';

/** x^y as Excel has it: 0^0 is #NUM!, 0 to a negative power #DIV/0!, a negative number to a fractional power #NUM!
 *  unless it is an odd root (=(-8)^(1/3) is -2). */
export function power(x, y) {
  if (x === 0 && y === 0) return E.num;
  if (x === 0 && y < 0) return E.div0;
  if (x < 0 && !Number.isInteger(y)) {
    const root = Math.round(1 / y);
    if (root % 2 && Math.abs(1 / y - root) < 1e-12) return finite(-((-x) ** y));
    return E.num;
  }
  return finite(x ** y);
}

/** An operator on two single values. `last`: it is the formula's last operation (closeToZero). */
function binary(op, a, b, ctx, last = false) {
  if (a instanceof XlError) return a;
  if (b instanceof XlError) return b;
  try {
    switch (op) {
      case '&': return toText(a) + toText(b);
      case '=': return compare(a, b) === 0;
      case '<>': return compare(a, b) !== 0;
      case '<': return compare(a, b) < 0;
      case '>': return compare(a, b) > 0;
      case '<=': return compare(a, b) <= 0;
      case '>=': return compare(a, b) >= 0;
    }
    const x = toNum(a, ctx.date1904);
    const y = toNum(b, ctx.date1904);
    switch (op) {
      case '+': return finite(last ? closeToZero(x, y, x + y) : x + y);
      case '-': return finite(last ? closeToZero(x, -y, x - y) : x - y);
      case '*': return finite(x * y);
      case '/': return y === 0 ? E.div0 : finite(x / y);
      case '^': return power(x, y);
    }
  } catch (e) {
    if (e instanceof XlError) return e;
    throw e;
  }
  throw new Error(`no operator ${op}`);
}

function unary(k, a, ctx) {
  if (a instanceof XlError) return a;
  try {
    const x = toNum(a, ctx.date1904);
    return k === 'neg' ? -x : k === 'pct' ? x / 100 : x;
  } catch (e) {
    if (e instanceof XlError) return e;
    throw e;
  }
}

/** An operator's operand: in a formula that isn't an array formula a range is cut to one value (implicit
 *  intersection) while an array constant stays whole (=SUM({1,2,3}*2) is 12, =SUM(A1:A3*2) is A-row*2). */
function operand(v, ctx, mode) {
  if (v instanceof Area) return mode === 'array' && v.size > 1 ? v : scalar(v, ctx);
  return v;
}

/** A reference to more than one cell, or an array. */
export const isMulti = (v) => (v instanceof Area && v.size > 1) || v instanceof Arr;

const dim = (p, q) => (p === 1 ? q : q === 1 ? p : Math.max(p, q));
/** Element (i, j) of an array spread to a larger size: a single row or column repeats, beyond its end is #N/A. */
export function pick(m, i, j) {
  const r = m.length === 1 ? 0 : i;
  const c = m[0].length === 1 ? 0 : j;
  return r < m.length && c < m[0].length ? m[r][c] : E.na;
}

/** `fn` element by element over several values spread to one size (Excel's broadcasting of arrays). */
export function elementwise(values, ctx, fn) {
  const ms = values.map((v) => matrix(v, ctx));
  const rows = ms.reduce((n, m) => dim(n, m.length), 1);
  const cols = ms.reduce((n, m) => dim(n, m[0].length), 1);
  const out = [];
  for (let i = 0; i < rows; i++) {
    const row = [];
    for (let j = 0; j < cols; j++) row.push(fn(ms.map((m) => pick(m, i, j))));
    out.push(row);
  }
  return new Arr(out);
}

function call(node, ctx, mode) {
  const def = FUNCTIONS[node.name];
  if (!def) throw new Unsupported(`the function ${node.name}`);
  if (def.volatile) ctx.volatile = true;
  const argMode = def.array ? 'array' : mode;
  const args = node.args.map((n) => evaluate(n, ctx, argMode));
  const run = (list) => {
    try {
      const v = def.fn(list, ctx, argMode);
      return v === undefined ? E.value : v;
    } catch (e) {
      if (e instanceof XlError) return e;
      throw e;
    }
  };
  // (A function of single values given arrays in an array formula gives an array: =SUMPRODUCT(--ISNUMBER(A1:A9)).)
  if (def.lift && (args.some((a) => a instanceof Arr) || (argMode === 'array' && args.some(isMulti)))) return elementwise(args, ctx, run);
  return run(args);
}

/** The value of a formula's tree (a number, text, boolean, error, null for a blank, or an Area, Arr or Union). */
export function evaluate(node, ctx, mode = 'scalar') {
  switch (node.k) {
    case 'num': case 'str': case 'bool': return node.v;
    case 'err': return errorOf(node.v);
    case 'missing': return MISSING;
    case 'paren': return evaluate(node.a, ctx, mode);
    case 'ref': {
      if (node.book != null) throw new Unsupported('a link to another workbook');
      const si = node.sheet == null ? ctx.here.si : ctx.sheetIndex(node.sheet);
      if (si < 0) return E.ref;
      const a = node.area;
      return new Area(si, a.r0, a.c0, a.r1, a.c1);
    }
    case 'union': {
      const items = node.items.map((n) => evaluate(n, ctx, mode));
      const err = items.find(isError);
      if (err) return err;
      return items.every((x) => x instanceof Area) ? new Union(items) : E.value;
    }
    case 'arr': return new Arr(node.rows.map((row) => row.map((v) => (v && v.err ? errorOf(v.err) : v))));
    case 'name': return ctx.name(node.name, mode);
    case 'neg': case 'pos': case 'pct': {
      const a = operand(evaluate(node.a, ctx, mode), ctx, mode);
      if (a instanceof Arr) return elementwise([a], ctx, ([x]) => unary(node.k, x, ctx));
      return unary(node.k, a, ctx);
    }
    case 'bin': {
      const a = operand(evaluate(node.a, ctx, mode), ctx, mode);
      const b = operand(evaluate(node.b, ctx, mode), ctx, mode);
      if (a instanceof Union || b instanceof Union) return E.value;
      if (a instanceof Arr || b instanceof Arr || a instanceof Area || b instanceof Area) {
        return elementwise([a, b], ctx, ([x, y]) => binary(node.op, x, y, ctx));
      }
      return binary(node.op, a, b, ctx, node === ctx.root);
    }
    case 'fn': return call(node, ctx, mode);
  }
  throw new Error(`no node ${node.k}`);
}

/** A formula's result as a cell holds it: one value (a blank is 0, as =A1 shows for an empty A1), or for an array
 *  formula its rows. */
export function result(tree, ctx, array = false) {
  ctx.root = tree;
  const v = evaluate(tree, ctx, array ? 'array' : 'scalar');
  const one = (x) => (x == null || x === MISSING ? 0 : x);
  if (array) return matrix(v instanceof Union ? E.value : v, ctx).map((row) => row.map(one));
  return one(v instanceof Union ? E.value : scalar(v, ctx));
}
