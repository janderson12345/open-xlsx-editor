// Formulas > Evaluate Formula: a formula worked out a step at a time, as Excel's box goes. The next part to work out is
// underlined; Evaluate puts its value in its place, shown in italics. A reference on its own, or as an operand, is a
// step (its value); a reference handed to a function goes with the function (=AVERAGE(D2:D5) to 40 at once, as
// Microsoft's example has it). IF works out its condition, then only the argument it returns (=IF(FALSE,SUM(E2:E5),0)
// to 0 without SUM); IFERROR its value, then its value_if_error only for an error. Brackets go with what is in them.
import { parseFormula } from './parse.js';
import { evaluate } from './evaluate.js';
import { Area, Arr, Union, XlError, MISSING, numberText, toBool, matrix } from './values.js';

const LITERAL = new Set(['num', 'str', 'bool', 'err', 'missing', 'arr', 'val']);
/** References: a step of their own as an operand, part of the function as its argument. */
const REFERENCE = new Set(['ref', 'struct', 'name', 'union']);

/** A value as Evaluate Formula writes it: "text" quoted, TRUE, #DIV/0!, an array {1,2;3,4}, a blank 0. */
export function valueText(v, ctx) {
  if (v instanceof XlError) return v.code;
  if (v == null || v === MISSING) return '0';
  if (typeof v === 'number') return numberText(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'string') return `"${v.replace(/"/g, '""')}"`;
  if (v instanceof Area || v instanceof Arr) {
    const rows = matrix(v, ctx);
    if (rows.length * (rows[0]?.length || 0) > 1000) return '{...}';
    if (rows.length === 1 && rows[0].length === 1) return valueText(rows[0][0], ctx);
    return `{${rows.map((row) => row.map((x) => valueText(x, ctx)).join(',')).join(';')}}`;
  }
  if (v instanceof Union) return '#VALUE!';
  return String(v);
}

/** The formula `text` of a cell worked out a step at a time in `ctx` (the engine's context for the cell). */
export class Stepper {
  constructor(text, ctx) {
    this.text = text;
    this.ctx = ctx;
    this.restart();
  }
  restart() {
    this.tree = parseFormula(this.text);
    this.last = null;
  }
  /** Whether it is worked out to its result. */
  get done() { return this.tree.k === 'val' || (LITERAL.has(this.tree.k) && this.tree.k !== 'missing'); }

  /** The next part to work out: { node, put(newNode) }, or null when it is done. */
  next() {
    return this._next(this.tree, (n) => { this.tree = n; }, false);
  }
  _next(node, put, asArg) {
    if (LITERAL.has(node.k)) return null;
    if (REFERENCE.has(node.k)) return asArg ? null : { node, put };
    const at = (key, i = null) => (n) => { if (i == null) node[key] = n; else node[key][i] = n; };
    switch (node.k) {
      case 'paren': return this._next(node.a, at('a'), asArg) || { node, put };
      case 'neg': case 'pos': case 'pct': return this._next(node.a, at('a'), false) || { node, put };
      case 'bin': return this._next(node.a, at('a'), false) || this._next(node.b, at('b'), false) || { node, put };
      // (Intersection and the range operator: one reference, as A1:A3 is; its parts aren't stepped into, where INDEX's
      // value would stand for the reference it gives.)
      case 'isect': case 'range': return asArg ? null : { node, put };
      case 'fn': {
        if (node.name === 'IF' || node.name === 'IFERROR') {
          const first = this._next(node.args[0], at('args', 0), false);
          if (first) return first;
          const v = this._value(node.args[0]);
          const pick = node.name === 'IF' ? (this._truthy(v) ? 1 : 2) : v instanceof XlError ? 1 : -1;
          const branch = pick >= 0 ? node.args[pick] : null;
          return (branch && this._next(branch, at('args', pick), false)) || { node, put };
        }
        for (let i = 0; i < node.args.length; i++) { const r = this._next(node.args[i], at('args', i), true); if (r) return r; }
        return { node, put };
      }
      case 'call': {
        for (let i = 0; i < node.args.length; i++) { const r = this._next(node.args[i], at('args', i), true); if (r) return r; }
        return { node, put };
      }
      default: return { node, put };
    }
  }
  _truthy(v) {
    try { return toBool(v instanceof Area || v instanceof Arr ? matrix(v, this.ctx)[0][0] : v); } catch { return false; }
  }
  _value(node) {
    this.ctx.root = null;
    return node.k === 'val' ? node.v : evaluate(node, this.ctx, 'scalar');
  }

  /** Where Step In goes: the cell the underlined part refers to ({ si, r, c }), when it is one cell's reference;
   *  else null. (Excel's Step In: the referenced cell's own formula worked out beneath; Step Out puts its value back.) */
  stepInTarget() {
    const n = this.next();
    if (!n || n.node.k !== 'ref') return null;
    this.ctx.root = null;
    let v;
    try { v = evaluate(n.node, this.ctx, 'scalar'); } catch { return null; }
    return v instanceof Area && v.r0 === v.r1 && v.c0 === v.c1 ? { si: v.si, r: v.r0, c: v.c0 } : null;
  }

  /** Work out the underlined part: its value in its place. */
  step() {
    const n = this.next();
    if (!n) return false;
    this.ctx.root = n.node === this.tree ? this.tree : null;
    let v = n.node.k === 'paren' && n.node.a.k === 'val' ? n.node.a.v : evaluate(n.node, this.ctx, 'scalar');
    // (One cell's value for one cell; a range keeps its cells, shown as an array.)
    if (v instanceof Area && v.size === 1) v = this.ctx.cellValue(v.si, v.r0, v.c0);
    const val = { k: 'val', v };
    n.put(val);
    this.last = val;
    // (Brackets round a value go: =(55+35)/2 is =90/2 next.)
    const unwrap = (node, put) => {
      if (!node || typeof node !== 'object') return;
      if (node.k === 'paren' && node.a === val) { put(val); return; }
      for (const [key, x] of Object.entries(node)) {
        if (Array.isArray(x)) x.forEach((y, i) => unwrap(y, (m) => { x[i] = m; }));
        else if (x && typeof x === 'object' && x.k) unwrap(x, (m) => { node[key] = m; });
      }
    };
    unwrap(this.tree, (m) => { this.tree = m; });
    return true;
  }

  /** The formula as it stands, in pieces [{ text, under, last }]: `under` for the part underlined, `last` for the value
   *  just worked out (which may be in the part underlined). Pieces alike are joined. */
  pieces() {
    const next = this.next()?.node || null;
    const out = [];
    let under = false;
    let last = false;
    const put = (text) => {
      if (!text) return;
      const prev = out[out.length - 1];
      if (prev && prev.under === under && prev.last === last) prev.text += text;
      else out.push({ text, under, last });
    };
    const print = (node) => {
      const was = [under, last];
      if (node === next) under = true;
      if (node === this.last) last = true;
      walk(node);
      [under, last] = was;
    };
    const walk = (node) => {
      switch (node.k) {
        case 'num': return put(node.src ?? numberText(node.v));
        case 'str': return put(`"${node.v.replace(/"/g, '""')}"`);
        case 'bool': return put(node.v ? 'TRUE' : 'FALSE');
        case 'err': return put(node.v);
        case 'missing': return null;
        case 'val': return put(valueText(node.v, this.ctx));
        case 'ref': case 'struct': return put(node.src ?? '#REF!');
        case 'name': return put(node.name);
        case 'arr': return put(`{${node.rows.map((row) => row.map((v) => (v && v.err ? v.err : typeof v === 'string' ? `"${v}"` : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : numberText(v))).join(',')).join(';')}}`);
        case 'paren': put('('); print(node.a); return put(')');
        case 'neg': put('-'); return print(node.a);
        case 'pos': put('+'); return print(node.a);
        case 'pct': print(node.a); return put('%');
        case 'bin': print(node.a); put(node.op); return print(node.b);
        case 'isect': print(node.a); put(' '); return print(node.b);
        case 'range': print(node.a); put(':'); return print(node.b);
        case 'union': put('('); node.items.forEach((x, i) => { if (i) put(','); print(x); }); return put(')');
        case 'fn': {
          if (node.name === 'SINGLE') { put('@'); return print(node.args[0]); }
          if (node.name === 'ANCHORARRAY' && node.args[0]?.k === 'ref') { print(node.args[0]); return put('#'); }
          put(`${node.name}(`);
          node.args.forEach((x, i) => { if (i) put(','); print(x); });
          return put(')');
        }
        case 'call': print(node.f); put('('); node.args.forEach((x, i) => { if (i) put(','); print(x); }); return put(')');
        default: return null;
      }
    };
    print(this.tree);
    return [{ text: '=', under: false, last: false }, ...out];
  }
}

