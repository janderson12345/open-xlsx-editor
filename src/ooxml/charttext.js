// A chart part's text taken apart without parsing it whole, so that an edit rewrites only the element it changes and
// the rest of the part is kept character for character (chartformat.js, chartcombo.js). Elements are found by their
// place: { name, start, inner (past the open tag), close (where the close tag starts), end (past it) }, offsets into
// the text. Chart parts hold no comments or CDATA; quoted attribute values may hold '>', and are skipped as such.

/** Where the tag opening at `lt` ends (the index of its '>'), quoted values skipped. */
function tagEnd(x, lt) {
  let q = null;
  for (let i = lt + 1; i < x.length; i++) {
    const ch = x[i];
    if (q) { if (ch === q) q = null; } else if (ch === '"' || ch === "'") q = ch;
    else if (ch === '>') return i;
  }
  throw new Error('Unclosed tag in the chart part');
}

/** The element whose open tag starts at `at`. */
export function elementAt(x, at) {
  const gt = tagEnd(x, at);
  const name = /^<([^\s/>]+)/.exec(x.slice(at, gt + 1))[1];
  if (x[gt - 1] === '/') return { name, start: at, inner: gt + 1, close: gt + 1, end: gt + 1, empty: true };
  let depth = 1;
  let i = gt + 1;
  for (;;) {
    const lt = x.indexOf('<', i);
    if (lt < 0) throw new Error(`Unclosed <${name}> in the chart part`);
    const e = tagEnd(x, lt);
    if (x[lt + 1] === '/') { if (--depth === 0) return { name, start: at, inner: gt + 1, close: lt, end: e + 1 }; }
    else if (x[e - 1] !== '/' && x[lt + 1] !== '?') depth++;
    i = e + 1;
  }
}

/** The direct children of element `el`. */
export function kidsOf(x, el) {
  const out = [];
  if (!el || el.empty) return out;
  let i = el.inner;
  for (;;) {
    const lt = x.indexOf('<', i);
    if (lt < 0 || lt >= el.close) return out;
    const k = elementAt(x, lt);
    out.push(k);
    i = k.end;
  }
}
/** The first child of `el` named `name` (with its prefix: 'c:title'), or null. */
export const kidOf = (x, el, name) => kidsOf(x, el).find((k) => k.name === name) || null;
/** The element at the end of a path of names from `el`, or null. */
export function pathOf(x, el, ...names) {
  let at = el;
  for (const n of names) { at = at && kidOf(x, at, n); if (!at) return null; }
  return at;
}
/** The part's root, its c:chartSpace. */
export function rootOf(x) {
  const at = x.indexOf('<c:chartSpace');
  if (at < 0) throw new Error("This chart can't be edited here yet.");
  return elementAt(x, at);
}
/** An element's text. */
export const textOf = (x, el) => x.slice(el.start, el.end);
/** A `val` attribute of an element's open tag, or null. */
export function valOf(x, el, attr = 'val') {
  if (!el) return null;
  const m = new RegExp(`\\s${attr}="([^"]*)"`).exec(x.slice(el.start, el.inner));
  return m ? m[1] : null;
}
/** The text with [start, end) replaced by `text`. */
export const splice = (x, start, end, text) => x.slice(0, start) + text + x.slice(end);

/** Where to put a new child of `el` so it comes before the first child named in `before` (else at the end). */
export function placeIn(x, el, before) {
  const k = kidsOf(x, el).find((c) => before.includes(c.name));
  return k ? k.start : el.close;
}

/** The text with a self-closing element opened: `<a:x/>` as `<a:x></a:x>`. */
export function opened(t) {
  const m = /^(<[^\s/>]+)([^>]*?)\/>$/.exec(t);
  return m ? `${m[1]}${m[2]}></${m[1].slice(1)}>` : t;
}

/** An open tag's attribute set (added after `after` when it has none, or straight after the name). */
export function setAttr(tag, name, value, after = []) {
  const re = new RegExp(`(\\s${name}=")[^"]*(")`);
  if (re.test(tag)) return tag.replace(re, `$1${value}$2`);
  for (const a of after) {
    const m = new RegExp(`\\s${a}="[^"]*"`).exec(tag);
    if (m) return tag.slice(0, m.index + m[0].length) + ` ${name}="${value}"` + tag.slice(m.index + m[0].length);
  }
  const n = /^<[^\s/>]+/.exec(tag)[0].length;
  return `${tag.slice(0, n)} ${name}="${value}"${tag.slice(n)}`;
}

// (xml.js's, for the chart modules that take it from here.)
export { esc } from './xml.js';
