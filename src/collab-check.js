// What another editor put in the room, checked before any of it reaches this workbook (collab-room.js): the room is
// written by every editor in it and comes over the network, so an entry that isn't what this version writes is left
// out rather than written into the workbook's XML. Names are XML names, attribute values text, counts whole numbers
// within the sheet's size, colours hex.
import { MAX_ROWS, MAX_COLS } from './ooxml/refs.js';

const NAME = /^[A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?$/;
const TOKEN = /^(?:\d+|[\w.-]+:\d+)$/;
const ID = /^[\w.-]+$/;
const HEX = /^[0-9A-Fa-f]{6}(?:[0-9A-Fa-f]{2})?$/;
const whole = (n, max) => Number.isInteger(n) && n >= 0 && n <= max;

/** Attributes { name: text } with XML names (none of `skip`, the element's own place), or null. */
export function checkAttrs(a, skip = new Set()) {
  if (!a || typeof a !== 'object' || Array.isArray(a)) return null;
  const out = {};
  for (const [k, v] of Object.entries(a)) {
    if (!NAME.test(k) || k.startsWith('xmlns') || skip.has(k) || (typeof v !== 'string' && typeof v !== 'number')) return null;
    out[k] = String(v);
  }
  return out;
}

/** An element as collab-styles.js describes one ([name, attrs, children]), no deeper than `depth`, or false. */
export function checkElement(e, depth = 6) {
  return Array.isArray(e) && e.length === 3 && typeof e[0] === 'string' && NAME.test(e[0]) && checkAttrs(e[1]) !== null
    && Array.isArray(e[2]) && e[2].length <= 64 && depth > 0 && e[2].every((c) => checkElement(c, depth - 1));
}

/** A format's description, as describeStyle writes it. */
export function checkStyle(d) {
  if (!d || typeof d !== 'object') return false;
  if (d.xf === null) return true;
  if (!checkElement(d.xf) || d.xf[0] !== 'xf') return false;
  for (const k of ['font', 'fill', 'border']) if (d[k] !== null && !checkElement(d[k])) return false;
  const n = d.numFmt;
  return !!n && (typeof n.code === 'string' ? n.code.length < 256 : whole(n.id, 163));
}

/** A line operation, as collab-capture.js writes one. */
export function checkOp(op) {
  if (!op || typeof op !== 'object' || !ID.test(op.id || '') || typeof op.sheet !== 'string' || !['row', 'col'].includes(op.axis)) return false;
  if (!Number.isInteger(op.cid) || !Number.isInteger(op.seq) || !Number.isInteger(op.ts)) return false;
  const max = op.axis === 'row' ? MAX_ROWS : MAX_COLS;
  if (op.kind === 'ins') return whole(op.n, max) && op.n > 0 && (op.after === null || (typeof op.after === 'string' && TOKEN.test(op.after)));
  if (op.kind === 'del') return Array.isArray(op.runs) && op.runs.every((r) => r && typeof r.id === 'string' && ID.test(r.id) && whole(r.start, max) && whole(r.len, max));
  return false;
}

/** Cell contents: { f } (a formula's text) or { v } (a number, text, true or false, or null). */
export const checkContents = (v) => !!v && typeof v === 'object' && (typeof v.f === 'string' || v.v === null || ['number', 'string', 'boolean'].includes(typeof v.v));

/** A tab colour as setTabColor takes it, or null. */
export const checkColor = (c) => c === null || (!!c && typeof c === 'object' && (typeof c.rgb === 'string' ? HEX.test(c.rgb) : Number.isInteger(c.theme) && c.theme >= 0 && c.theme < 64 && (c.tint == null || (typeof c.tint === 'number' && c.tint >= -1 && c.tint <= 1))));

export const isToken = (t) => typeof t === 'string' && TOKEN.test(t);
