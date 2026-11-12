// Working out a formula's tree: operators, references, names and function calls. A formula that isn't an array formula
// works on single values: a range given to an operator is cut down to the formula's own row or column (implicit
// intersection). The arguments of SUMPRODUCT and its kind, and array formulas, work on whole arrays, element by element.
import { Builtin } from './values.js';
import { Area, Arr, Union, Ref3D, MISSING, XlError, E, errorOf, Unsupported, isError, scalar, matrix, toNum, toText, compare, finite,
  closeToZero, Lambda } from './values.js';
import { FUNCTIONS } from './functions.js';
import { threeD, intersect, refused3D, answer3D } from './threed.js';
// (More functions, registered into FUNCTIONS once it is made.)
import './functions-text.js';
import './functions-lookup.js';
import './functions-more.js';
import './functions-dynamic.js';
import './functions-arrays.js';
import './functions-finance.js';
import './functions-stats.js';
import './functions-dist.js';
import './functions-agg.js';
import './functions-eng.js';
import './functions-db.js';
import './functions-bonds.js';
import './functions-complex.js';
import './functions-convert.js';
import './functions-regex.js';
import './functions-amort.js';
import './functions-bessel.js';
import './functions-oddbonds.js';
import './functions-groupby.js';
import './functions-pivotby.js';

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

/** A LET or LAMBDA name as the scope keeps it: lower case, without the file's _xlpm. */
const scoped = (name) => name.replace(/^_xlpm\./i, '').toLowerCase();
/** What a name stands for in the LETs and LAMBDAs around it, or undefined. */
function bound(ctx, name) {
  const key = scoped(name);
  for (let k = (ctx.scope || []).length - 1; k >= 0; k--) if (ctx.scope[k].has(key)) return ctx.scope[k].get(key);
  return undefined;
}
/** Work out `node` with names bound (`binds`, a Map) on top of `scope`. */
function within(scope, binds, node, ctx, mode) {
  const was = ctx.scope;
  ctx.scope = [...(scope || []), binds];
  try { return evaluate(node, ctx, mode); } finally { ctx.scope = was; }
}
/** LET(name1, value1, ..., calculation): each name bound in turn, the last argument worked out with them. */
function letOf(node, ctx, mode) {
  const args = node.args;
  if (args.length < 3 || args.length % 2 === 0) return E.value;
  const binds = new Map();
  const was = ctx.scope;
  ctx.scope = [...(was || []), binds];
  try {
    for (let k = 0; k < args.length - 1; k += 2) {
      if (args[k].k !== 'name') return E.name;
      binds.set(scoped(args[k].name), evaluate(args[k + 1], ctx, mode));
    }
    return evaluate(args[args.length - 1], ctx, mode);
  } finally { ctx.scope = was; }
}
/** A LAMBDA called with arguments (worked out where it is called). */
function apply(fn, argNodes, ctx, mode) {
  if (isError(fn)) return fn;
  if (fn instanceof Builtin) return call({ k: 'fn', name: fn.name, args: argNodes }, ctx, mode);
  if (!(fn instanceof Lambda)) return E.value;
  if (argNodes.length !== fn.params.length) return E.value;
  const binds = new Map(fn.params.map((p, k) => [p, evaluate(argNodes[k], ctx, mode)]));
  return within(fn.scope, binds, fn.body, ctx, mode);
}

/** A LAMBDA called with values already worked out (MAP, REDUCE, SCAN, BYROW...: functions-arrays.js). */
function applyValues(fn, vals, ctx) {
  if (isError(fn)) return fn;
  if (fn instanceof Builtin) return call({ k: 'fn', name: fn.name, args: vals.map((v) => ({ k: 'val', v })) }, ctx, 'array');
  if (!(fn instanceof Lambda) || vals.length !== fn.params.length) return E.value;
  const binds = new Map(fn.params.map((p, k) => [p, vals[k]]));
  const v = within(fn.scope, binds, fn.body, ctx, 'array');
  return v instanceof Lambda ? E.calc : v;
}

function call(node, ctx, mode) {
  ctx.callLambda ||= (fn, vals) => applyValues(fn, vals, ctx);
  if (node.name === 'LET') return letOf(node, ctx, mode);
  if (node.name === 'LAMBDA') {
    const params = node.args.slice(0, -1);
    if (!node.args.length || params.some((p) => p.k !== 'name')) return E.value;
    return new Lambda(params.map((p) => scoped(p.name)), node.args[node.args.length - 1], [...(ctx.scope || [])]);
  }
  // (A name LET bound to a LAMBDA, called: =LET(f, LAMBDA(x, x*2), f(3)).)
  const local = bound(ctx, node.name);
  if (local !== undefined) return apply(local, node.args, ctx, mode);
  const def = FUNCTIONS[node.name];
  // (Not a function: a defined name holding a LAMBDA, called with these arguments, =Twice(21), as Excel 365 has it; a
  // function of the same name comes first, as =Fact(5) is FACT's.)
  if (!def && ctx.hasName?.(node.name)) return apply(ctx.name(node.name, mode), node.args, ctx, mode);
  if (!def) throw new Unsupported(`the function ${node.name}`);
  if (def.volatile) ctx.volatile = true;
  const argMode = def.array ? 'array' : mode;
  const args = node.args.map((n) => evaluate(n, ctx, argMode));
  const refused = refused3D(node.name, args);
  if (refused) return refused;
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
    // (A value put in a part's place: Evaluate Formula, stepper.js.)
    case 'val': return node.v;
    case 'paren': return evaluate(node.a, ctx, mode);
    case 'ref': {
      if (node.book != null) throw new Unsupported('a link to another workbook');
      if (node.sheets) return threeD(node, ctx);
      const si = node.sheet == null ? ctx.here.si : ctx.sheetIndex(node.sheet);
      if (si < 0) return E.ref;
      const a = node.area;
      return new Area(si, a.r0, a.c0, a.r1, a.c1);
    }
    case 'isect': return intersect(evaluate(node.a, ctx, mode), evaluate(node.b, ctx, mode));
    case 'union': {
      const items = node.items.map((n) => evaluate(n, ctx, mode));
      const err = items.find(isError);
      if (err) return err;
      return items.every((x) => x instanceof Area) ? new Union(items) : E.value;
    }
    case 'arr': return new Arr(node.rows.map((row) => row.map((v) => (v && v.err ? errorOf(v.err) : v))));
    case 'name': {
      const local = bound(ctx, node.name);
      if (local !== undefined) return local;
      // (A function named on its own, SUM or the file's _xleta.SUM: an eta LAMBDA, unless a defined name is called so.)
      const eta = node.name.replace(/^_xleta\./i, '').toUpperCase();
      if (FUNCTIONS[eta] && (/^_xleta\./i.test(node.name) || !ctx.hasName?.(node.name))) return new Builtin(eta);
      return ctx.name(node.name, mode);
    }
    case 'call': return apply(evaluate(node.f, ctx, mode), node.args, ctx, mode);
    case 'struct': {
      // (A table reference, as the range it stands for from this cell: ooxml/tables.js.)
      const a = ctx.tableArea ? ctx.tableArea(node.table, node.spec) : { error: '#REF!' };
      return a.error ? errorOf(a.error) : new Area(a.si, a.r0, a.c0, a.r1, a.c1);
    }
    case 'neg': case 'pos': case 'pct': {
      const a = operand(evaluate(node.a, ctx, mode), ctx, mode);
      if (a instanceof Arr) return elementwise([a], ctx, ([x]) => unary(node.k, x, ctx));
      return unary(node.k, a, ctx);
    }
    case 'bin': {
      const a = operand(evaluate(node.a, ctx, mode), ctx, mode);
      const b = operand(evaluate(node.b, ctx, mode), ctx, mode);
      if (a instanceof Ref3D || b instanceof Ref3D) return E.ref;
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
  let v = evaluate(tree, ctx, array ? 'array' : 'scalar');
  if (v instanceof Lambda) v = E.calc; // (a LAMBDA not called: Excel's #CALC!)
  const one = (x) => (x == null || x === MISSING ? 0 : x);
  // (A 3D reference as a cell's answer is #REF! in Excel; several areas #VALUE!.)
  if (v instanceof Ref3D) v = answer3D(tree);
  if (array) return matrix(v instanceof Union ? E.value : v, ctx).map((row) => row.map(one));
  return one(v instanceof Union ? E.value : scalar(v, ctx));
}

/** A dynamic array formula's result (Excel 365's, entered or saved with its metadata): { rows } when it gives an array
 *  (even of one value, as SEQUENCE(1) does), to spill from its cell; else { value }. */
/** The functions whose answer is an array to Excel, whatever its size (a formula of them is saved as an array). */
const ARRAY_FUNCTIONS = new Set(['SEQUENCE', 'SORT', 'SORTBY', 'UNIQUE', 'FILTER', 'TRANSPOSE']);

export function dynamicResult(tree, ctx) {
  ctx.root = tree;
  let v = evaluate(tree, ctx, 'array');
  if (v instanceof Lambda) v = E.calc;
  const one = (x) => (x == null || x === MISSING ? 0 : x);
  if (v instanceof Ref3D) return { value: answer3D(tree) };
  if (v instanceof Union) return { value: E.value };
  if (v instanceof Arr || (v instanceof Area && v.size > 1)) return { rows: matrix(v, ctx).map((row) => row.map(one)) };
  // (A function that gives arrays, or a LAMBDA called, is an array formula to Excel even when its answer is one value,
  // FILTER's if_empty or its #CALC!: kept as one of one, as Excel saves it.)
  const value = one(scalar(v, ctx));
  if (tree.k === 'call' || (tree.k === 'fn' && ARRAY_FUNCTIONS.has(tree.name))) return { rows: [[value]] };
  return { value };
}
