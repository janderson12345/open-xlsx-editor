// Cell formats in a co-edited workbook (collab.js), shared by what they are rather than by their index: each editor
// adds formats to its own styles part as it goes (stylesedit.js: an equal one reused, a new one at the end), so the
// same format can have a different index in two editors. A format goes into the room as its description (the cell
// format with its font, fill, border and number format written out) under a key made from that description; an
// editor taking a cell's key looks for an equal format of its own, and adds one at the end when it has none, as
// formatting the cell there would.
import { kids, kid, attr } from './ooxml/xml.js';
import { formatsOf } from './ooxml/stylesedit.js';
import { readFont, readFill, readBorder, readXf } from './ooxml/styles.js';

const LISTS = [['fontId', 'fonts', 'font'], ['fillId', 'fills', 'fill'], ['borderId', 'borders', 'border']];

/** An element as plain data, its attributes in name order: [name, { attr: value }, children]. */
function canon(el) {
  const attrs = {};
  for (const a of [...el.attributes].sort((p, q) => (p.name < q.name ? -1 : 1))) if (!a.name.startsWith('xmlns')) attrs[a.name] = a.value;
  return [el.localName, attrs, [...el.childNodes].filter((n) => n.nodeType === 1).map(canon)];
}

/** canon()'s data back as an element of `doc` (in its root's namespace and prefix). */
function build(doc, [name, attrs, children]) {
  const root = doc.documentElement;
  const el = doc.createElementNS(root.namespaceURI, root.prefix ? `${root.prefix}:${name}` : name);
  for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, v);
  for (const c of children) el.appendChild(build(doc, c));
  return el;
}

/** FNV-1a, twice (two seeds), as 16 hex digits: short keys for descriptions that may run to hundreds of characters. */
function hash(text) {
  let a = 0x811c9dc5;
  let b = 0x01000193 ^ 0x5bd1e995;
  for (let i = 0; i < text.length; i++) {
    const ch = text.charCodeAt(i);
    a = Math.imul(a ^ ch, 0x01000193) >>> 0;
    b = Math.imul(b ^ ch, 0x5bd1e995) >>> 0;
  }
  return a.toString(16).padStart(8, '0') + b.toString(16).padStart(8, '0');
}

const items = (root, list, item) => kids(kid(root, list, root.namespaceURI), item, root.namespaceURI);

/** Cell format `s` of `wb` as a description: { xf (without its indexes), font, fill, border, numFmt: { id } for a
 *  built-in number format, { code } for the workbook's own }. */
export function describeStyle(wb, s) {
  const doc = wb.stylesDoc;
  if (!doc) return { xf: null };
  const root = doc.documentElement;
  const xfEl = items(root, 'cellXfs', 'xf')[s] || items(root, 'cellXfs', 'xf')[0];
  if (!xfEl) return { xf: null };
  const out = { xf: canon(xfEl) };
  for (const [key, list, item] of LISTS) {
    delete out.xf[1][key];
    const el = items(root, list, item)[Number(attr(xfEl, key) || 0)];
    out[item] = el ? canon(el) : null;
  }
  const id = Number(attr(xfEl, 'numFmtId') || 0);
  delete out.xf[1].numFmtId;
  const code = wb.styles.numFmts?.[id];
  out.numFmt = code != null ? { code } : { id };
  return out;
}

/** Keys and descriptions of a workbook's formats, worked out once per format (a workbook's formats never change in
 *  place: a change is a new one at the end). */
export function styleKeys(wb) {
  const byIndex = new Map();
  const byKey = new Map();
  return {
    /** The key of format `s`, and its description. */
    keyOf(s) {
      if (!byIndex.has(s)) {
        const desc = describeStyle(wb, s);
        const key = hash(JSON.stringify(desc));
        byIndex.set(s, { key, desc });
        if (!byKey.has(key)) byKey.set(key, s);
      }
      return byIndex.get(s);
    },
    /** The index of the format `desc` describes (`key`): this workbook's own equal one, or a new one added. */
    indexOf(key, desc) {
      if (byKey.has(key)) return byKey.get(key);
      const n = wb.stylesDoc ? items(wb.stylesDoc.documentElement, 'cellXfs', 'xf').length : 0;
      for (let s = 0; s < n; s++) if (this.keyOf(s).key === key) return s;
      if (!desc || !desc.xf) return 0;
      const s = addStyle(wb, desc);
      byIndex.delete(s);
      byKey.set(key, s);
      return s;
    },
  };
}

/** A format described by `desc` added to wb's styles part (its font, fill, border and number format found or added
 *  first): its index. */
function addStyle(wb, desc) {
  const f = formatsOf(wb);
  const ns = f.ns;
  const styles = wb.styles;
  const read = { fonts: (x, i) => { styles.fonts[i] = readFont(x, ns); }, fills: (x, i) => { styles.fills[i] = readFill(x, ns); }, borders: (x, i) => { styles.borders[i] = readBorder(x, ns); } };
  // (Its indexes first, in the order Excel writes an xf's attributes; the rest after.)
  const [name, attrs, children] = desc.xf;
  const xf = build(f.doc, [name, {}, children]);
  xf.setAttribute('numFmtId', String(desc.numFmt.code != null ? f.numFmt(desc.numFmt.code) : desc.numFmt.id));
  for (const [key, list, item] of LISTS) xf.setAttribute(key, String(desc[item] ? f.add(list, build(f.doc, desc[item]), read[list]) : 0));
  for (const k of ['xfId', ...Object.keys(attrs).filter((k) => k !== 'xfId')]) if (attrs[k] != null) xf.setAttribute(k, attrs[k]);
  return f.add('cellXfs', xf, (x, i) => { styles.cellXfs[i] = readXf(x, ns, styles); });
}
