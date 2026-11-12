// XPath 1.0 over a parsed XML document, for FILTERXML: location paths with every axis but namespace, name, * and node
// type tests, predicates, the operators (| or and = != < <= > >= + - * div mod), and the core functions on node sets,
// strings, numbers and booleans. Element names are matched as written, case and prefix included, as Excel's MSXML
// matches them. A query it can't read throws.

const AXES = new Set(['ancestor', 'ancestor-or-self', 'attribute', 'child', 'descendant', 'descendant-or-self', 'following',
  'following-sibling', 'parent', 'preceding', 'preceding-sibling', 'self']);
const TYPES = new Set(['node', 'text', 'comment', 'processing-instruction']);

/** The query's tokens: { t: 'name' | 'str' | 'num' | 'op', v }. */
function tokenize(q) {
  const out = [];
  let i = 0;
  // (XPath's rule: after a token that isn't @, ::, (, [, a comma or an operator, * multiplies and and, or, div and mod
  // are operators; elsewhere * and those words are names.)
  const OPERATORS = ['@', '::', '(', '[', ',', '/', '//', '|', '+', '-', '=', '!=', '<', '<=', '>', '>=', '*', 'and', 'or', 'mod', 'div'];
  const after = () => { const p = out[out.length - 1]; return !!p && !(p.t === 'op' && OPERATORS.includes(p.v)); };
  while (i < q.length) {
    const ch = q[i];
    if (/\s/.test(ch)) { i++; continue; }
    const two = q.slice(i, i + 2);
    if (['//', '::', '!=', '<=', '>=', '..'].includes(two)) { out.push({ t: 'op', v: two }); i += 2; continue; }
    if (ch === '"' || ch === "'") {
      const end = q.indexOf(ch, i + 1);
      if (end < 0) throw new Error('literal');
      out.push({ t: 'str', v: q.slice(i + 1, end) }); i = end + 1; continue;
    }
    if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(q[i + 1] || ''))) {
      const m = /^[0-9]*\.?[0-9]*/.exec(q.slice(i))[0];
      out.push({ t: 'num', v: Number(m) }); i += m.length; continue;
    }
    if ('/()[]@,|+-=<>.$'.includes(ch)) { out.push({ t: 'op', v: ch }); i++; continue; }
    if (ch === '*') { out.push({ t: 'op', v: after() ? '*' : '*n' }); i++; continue; }
    const m = /^[A-Za-z_][\w.-]*(?::(?:[A-Za-z_][\w.-]*|\*))?/.exec(q.slice(i));
    if (!m) throw new Error('token');
    const v = m[0];
    const isOperator = after() && ['and', 'or', 'div', 'mod'].includes(v);
    out.push(isOperator ? { t: 'op', v } : { t: 'name', v });
    i += v.length;
  }
  return out;
}

/** The query parsed into a tree. */
function parse(q) {
  const toks = tokenize(q);
  let k = 0;
  const peek = (o = 0) => toks[k + o];
  const isOp = (v, o = 0) => peek(o)?.t === 'op' && peek(o).v === v;
  const take = (v) => { if (!isOp(v)) throw new Error(`expected ${v}`); k++; };
  const binary = (next, ops) => () => {
    let left = next();
    while (peek()?.t === 'op' && ops.includes(peek().v)) { const op = toks[k++].v; left = { op, left, right: next() }; }
    return left;
  };
  const step = () => {
    if (isOp('.')) { k++; return { axis: 'self', test: { type: 'node' }, preds: [] }; }
    if (isOp('..')) { k++; return { axis: 'parent', test: { type: 'node' }, preds: [] }; }
    let axis = 'child';
    if (isOp('@')) { k++; axis = 'attribute'; } else if (peek()?.t === 'name' && isOp('::', 1)) {
      axis = toks[k].v; k += 2;
      if (!AXES.has(axis)) throw new Error('axis');
    }
    let test;
    if (isOp('*n') || isOp('*')) { k++; test = { name: '*' }; } else if (peek()?.t === 'name') {
      const n = toks[k++].v;
      if (TYPES.has(n) && isOp('(')) { k++; if (peek()?.t === 'str') k++; take(')'); test = { type: n }; } else test = { name: n };
    } else throw new Error('step');
    const preds = [];
    while (isOp('[')) { k++; preds.push(expr()); take(']'); }
    return { axis, test, preds };
  };
  const relative = (steps) => {
    steps.push(step());
    while (isOp('/') || isOp('//')) {
      if (toks[k++].v === '//') steps.push({ axis: 'descendant-or-self', test: { type: 'node' }, preds: [] });
      steps.push(step());
    }
    return steps;
  };
  const startsStep = () => { const p = peek(); return p && ((p.t === 'name' && !isOp('(', 1)) || (p.t === 'name' && TYPES.has(p.v)) || ['@', '.', '..', '*n', '*'].includes(p.t === 'op' ? p.v : '')); };
  const primary = () => {
    const p = peek();
    if (!p) throw new Error('end');
    if (isOp('(')) { k++; const e = expr(); take(')'); return e; }
    if (p.t === 'str') { k++; return { lit: p.v }; }
    if (p.t === 'num') { k++; return { num: p.v }; }
    if (p.t === 'name' && isOp('(', 1) && !TYPES.has(p.v)) {
      k += 2;
      const args = [];
      if (!isOp(')')) { args.push(expr()); while (isOp(',')) { k++; args.push(expr()); } }
      take(')');
      return { fn: p.v, args };
    }
    throw new Error('primary');
  };
  const path = () => {
    if (isOp('/')) {
      k++;
      return { path: startsStep() ? relative([]) : [], absolute: true };
    }
    if (isOp('//')) { k++; return { path: relative([{ axis: 'descendant-or-self', test: { type: 'node' }, preds: [] }]), absolute: true }; }
    if (startsStep()) return { path: relative([]), absolute: false };
    let e = primary();
    const preds = [];
    while (isOp('[')) { k++; preds.push(expr()); take(']'); }
    if (preds.length) e = { filter: e, preds };
    if (isOp('/') || isOp('//')) {
      const steps = [];
      if (toks[k++].v === '//') steps.push({ axis: 'descendant-or-self', test: { type: 'node' }, preds: [] });
      e = { from: e, path: relative(steps) };
    }
    return e;
  };
  const union = binary(path, ['|']);
  const unary = () => { if (isOp('-')) { k++; return { neg: unary() }; } return union(); };
  const mul = binary(unary, ['*', 'div', 'mod']);
  const add = binary(mul, ['+', '-']);
  const rel = binary(add, ['<', '<=', '>', '>=']);
  const eq = binary(rel, ['=', '!=']);
  const and = binary(eq, ['and']);
  const expr = binary(and, ['or']);
  const tree = expr();
  if (k < toks.length) throw new Error('trailing');
  return tree;
}

// ---- Evaluation

const isNodes = (v) => Array.isArray(v);
const kidsOf = (n) => (n.childNodes ? [...n.childNodes] : []);
const attrsOf = (n) => (n.attributes ? [...n.attributes].filter((a) => a.name !== 'xmlns' && !a.name.startsWith('xmlns:')) : []);

/** A node's string-value: an element's or document's text below it, a text or attribute node's own. */
export function stringOf(n) {
  if (n.nodeType === 1 || n.nodeType === 9) { let s = ''; for (const c of kidsOf(n)) if (c.nodeType === 3 || c.nodeType === 4) s += c.data; else if (c.nodeType === 1) s += stringOf(c); return s; }
  return n.nodeValue ?? n.data ?? '';
}
const str = (v) => (isNodes(v) ? (v.length ? stringOf(v[0]) : '') : typeof v === 'number' ? numText(v) : typeof v === 'boolean' ? String(v) : v);
const numText = (x) => (Number.isNaN(x) ? 'NaN' : Number.isInteger(x) ? String(x) : String(x));
const num = (v) => { if (typeof v === 'number') return v; if (typeof v === 'boolean') return v ? 1 : 0; const s = str(v).trim(); return s === '' || !/^-?(\d+\.?\d*|\.\d+)$/.test(s) ? NaN : Number(s); };
const bool = (v) => (isNodes(v) ? v.length > 0 : typeof v === 'number' ? v !== 0 && !Number.isNaN(v) : typeof v === 'string' ? v.length > 0 : v);

function axis(n, name) {
  const out = [];
  const down = (x) => { for (const c of kidsOf(x)) { out.push(c); down(c); } };
  const parentOf = (x) => x.ownerElement || x.parentNode;
  switch (name) {
    case 'child': return kidsOf(n);
    case 'attribute': return attrsOf(n);
    case 'self': return [n];
    case 'parent': return parentOf(n) ? [parentOf(n)] : [];
    case 'descendant': down(n); return out;
    case 'descendant-or-self': out.push(n); down(n); return out;
    case 'ancestor': case 'ancestor-or-self': {
      if (name === 'ancestor-or-self') out.push(n);
      for (let p = parentOf(n); p; p = p.parentNode) out.push(p);
      return out;
    }
    case 'following-sibling': for (let s = n.nextSibling; s; s = s.nextSibling) out.push(s); return out;
    case 'preceding-sibling': for (let s = n.previousSibling; s; s = s.previousSibling) out.push(s); return out;
    case 'following': for (let x = n; x; x = x.parentNode) for (let s = x.nextSibling; s; s = s.nextSibling) { out.push(s); down(s); } return out;
    case 'preceding': {
      const anc = new Set(axis(n, 'ancestor'));
      const all = [];
      const walk = (x) => { for (const c of kidsOf(x)) { all.push(c); walk(c); } };
      walk(n.ownerDocument || n);
      return all.slice(0, all.indexOf(n)).filter((x) => !anc.has(x)).reverse();
    }
    default: throw new Error('axis');
  }
}

function matchesTest(n, test, ax) {
  if (test.type) {
    if (test.type === 'node') return true;
    if (test.type === 'text') return n.nodeType === 3 || n.nodeType === 4;
    if (test.type === 'comment') return n.nodeType === 8;
    return n.nodeType === 7;
  }
  const principal = ax === 'attribute' ? 2 : 1;
  if (n.nodeType !== principal) return false;
  if (test.name === '*') return true;
  if (test.name.endsWith(':*')) return (n.prefix || '') === test.name.slice(0, -2);
  return n.nodeName === test.name;
}

/** Document order of the nodes, duplicates gone. */
function inOrder(nodes, doc) {
  const order = new Map();
  let i = 0;
  const walk = (x) => { order.set(x, i++); for (const a of attrsOf(x)) order.set(a, i++); for (const c of kidsOf(x)) walk(c); };
  walk(doc);
  return [...new Set(nodes)].sort((a, b) => (order.get(a) ?? 0) - (order.get(b) ?? 0));
}


function applyPreds(nodes, preds, doc, vars) {
  let cur = nodes;
  for (const p of preds) {
    const size = cur.length;
    cur = cur.filter((n, i) => {
      const v = ev(p, { node: n, pos: i + 1, size, doc, vars });
      return typeof v === 'number' ? v === i + 1 : bool(v);
    });
  }
  return cur;
}

function runPath(starts, steps, doc, vars) {
  let cur = starts;
  for (const s of steps) {
    const next = [];
    for (const n of cur) {
      let found = axis(n, s.axis).filter((x) => matchesTest(x, s.test, s.axis));
      // (Predicates count along the axis: a reverse one (ancestors, preceding) lists its nodes nearest first.)
      found = applyPreds(found, s.preds, doc, vars);
      next.push(...found);
    }
    cur = inOrder(next, doc);
  }
  return cur;
}

function compare(a, op, b) {
  const test = (x, y) => {
    if (op === '=' || op === '!=') {
      const eq = typeof x === 'boolean' || typeof y === 'boolean' ? bool(x) === bool(y) : typeof x === 'number' || typeof y === 'number' ? num(x) === num(y) : str(x) === str(y);
      return op === '=' ? eq : !eq;
    }
    const l = num(x);
    const r = num(y);
    return op === '<' ? l < r : op === '<=' ? l <= r : op === '>' ? l > r : l >= r;
  };
  // (A node set against anything: true when one of its nodes' string-values is; against a boolean, as a boolean.)
  if (typeof a === 'boolean' || typeof b === 'boolean') return test(bool(a), bool(b));
  const each = (v) => (isNodes(v) ? v.map((n) => stringOf(n)) : [v]);
  return each(a).some((x) => each(b).some((y) => test(x, y)));
}

const FNS = {
  last: (c) => c.size,
  position: (c) => c.pos,
  count: (c, [v]) => { if (!isNodes(v)) throw new Error('count'); return v.length; },
  'local-name': (c, [v]) => { const n = v ? v[0] : c.node; return n ? (n.localName || n.nodeName || '') : ''; },
  name: (c, [v]) => { const n = v ? v[0] : c.node; return n && n.nodeType !== 9 ? n.nodeName : ''; },
  string: (c, [v]) => (v === undefined ? stringOf(c.node) : str(v)),
  concat: (c, args) => args.map(str).join(''),
  'starts-with': (c, [a, b]) => str(a).startsWith(str(b)),
  contains: (c, [a, b]) => str(a).includes(str(b)),
  'substring-before': (c, [a, b]) => { const s = str(a); const i = s.indexOf(str(b)); return i < 0 ? '' : s.slice(0, i); },
  'substring-after': (c, [a, b]) => { const s = str(a); const t = str(b); const i = s.indexOf(t); return i < 0 ? '' : s.slice(i + t.length); },
  substring: (c, [a, b, n]) => {
    const s = [...str(a)];
    const from = Math.round(num(b));
    const to = n === undefined ? Infinity : from + Math.round(num(n));
    return s.filter((_, i) => i + 1 >= from && i + 1 < to).join('');
  },
  'string-length': (c, [v]) => [...(v === undefined ? stringOf(c.node) : str(v))].length,
  'normalize-space': (c, [v]) => (v === undefined ? stringOf(c.node) : str(v)).trim().replace(/\s+/g, ' '),
  translate: (c, [a, b, t]) => { const from = [...str(b)]; const to = [...str(t)]; return [...str(a)].map((ch) => { const i = from.indexOf(ch); return i < 0 ? ch : (to[i] ?? ''); }).join(''); },
  boolean: (c, [v]) => bool(v),
  not: (c, [v]) => !bool(v),
  true: () => true,
  false: () => false,
  number: (c, [v]) => num(v === undefined ? [c.node] : v),
  sum: (c, [v]) => { if (!isNodes(v)) throw new Error('sum'); return v.reduce((s, n) => s + num(stringOf(n)), 0); },
  floor: (c, [v]) => Math.floor(num(v)),
  ceiling: (c, [v]) => Math.ceil(num(v)),
  round: (c, [v]) => Math.round(num(v)),
};

function ev(e, c) {
  if ('lit' in e) return e.lit;
  if ('num' in e) return e.num;
  if (e.neg) return -num(ev(e.neg, c));
  if (e.fn) {
    const f = FNS[e.fn];
    if (!f) throw new Error(`function ${e.fn}`);
    return f(c, e.args.map((a) => ev(a, c)));
  }
  if (e.path) {
    if (e.from) { const base = ev(e.from, c); if (!isNodes(base)) throw new Error('path'); return runPath(base, e.path, c.doc, c.vars); }
    return runPath([e.absolute ? c.doc : c.node], e.path, c.doc, c.vars);
  }
  if (e.filter) { const base = ev(e.filter, c); if (!isNodes(base)) throw new Error('filter'); return applyPreds(base, e.preds, c.doc, c.vars); }
  const l = ev(e.left, c);
  switch (e.op) {
    case 'or': return bool(l) || bool(ev(e.right, c));
    case 'and': return bool(l) && bool(ev(e.right, c));
    case '|': { const r = ev(e.right, c); if (!isNodes(l) || !isNodes(r)) throw new Error('union'); return inOrder([...l, ...r], c.doc); }
    case '=': case '!=': case '<': case '<=': case '>': case '>=': return compare(l, e.op, ev(e.right, c));
    case '+': return num(l) + num(ev(e.right, c));
    case '-': return num(l) - num(ev(e.right, c));
    case '*': return num(l) * num(ev(e.right, c));
    case 'div': return num(l) / num(ev(e.right, c));
    case 'mod': return num(l) % num(ev(e.right, c));
    default: throw new Error(`op ${e.op}`);
  }
}

/** `query` over document `doc`: a node set (an array of nodes, in document order), a string, a number or a boolean. */
export function xpath(doc, query) {
  return ev(parse(query), { node: doc, pos: 1, size: 1, doc, vars: {} });
}
