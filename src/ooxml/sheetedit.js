// Changing cells in a worksheet part's XML in place. Only the `row` and `c` elements an edit needs are touched;
// everything else in the part (and the other cells' markup) stays as it came.
import { kids, kid, attr } from './xml.js';
import { parseRef, refName, parseRange, rangeName } from './refs.js';

const isOn = (v) => v === '1' || v === 'true';
/** A new element in the part's namespace, under the root's prefix when the part uses one (`x:c`). */
const make = (doc, ns, name) => doc.createElementNS(ns, doc.documentElement.prefix ? `${doc.documentElement.prefix}:${name}` : name);

/** Give every row and cell of sheetData an explicit `r` where the file left it implied, so inserting one doesn't move
 *  the ones after it. (Allowed by the schema and kept by Excel; changes nothing a reader sees.) */
function pinRefs(sheetData, ns) {
  let r = -1;
  for (const rowEl of kids(sheetData, 'row', ns)) {
    r = attr(rowEl, 'r') ? Number(attr(rowEl, 'r')) - 1 : r + 1;
    if (!attr(rowEl, 'r')) rowEl.setAttribute('r', String(r + 1));
    let c = -1;
    for (const cEl of kids(rowEl, 'c', ns)) {
      const at = parseRef(attr(cEl, 'r'));
      c = at ? at.c : c + 1;
      if (!at) cEl.setAttribute('r', refName(r, c));
    }
  }
}

/** The `row` element for row r, made (in order) if there isn't one. */
function rowElement(doc, sheetData, ns, r) {
  for (const rowEl of kids(sheetData, 'row', ns)) {
    const n = Number(attr(rowEl, 'r')) - 1;
    if (n === r) return rowEl;
    if (n > r) return sheetData.insertBefore(newRow(doc, ns, r), rowEl);
  }
  return sheetData.appendChild(newRow(doc, ns, r));
}
const newRow = (doc, ns, r) => {
  const el = make(doc, ns, 'row');
  el.setAttribute('r', String(r + 1));
  return el;
};

/** The `c` element for column c of a row, made (in order, with the row's style if it has one) if there isn't one. */
function cellElement(doc, rowEl, ns, r, c) {
  const cell = () => {
    const el = make(doc, ns, 'c');
    el.setAttribute('r', refName(r, c));
    if (isOn(attr(rowEl, 'customFormat')) && attr(rowEl, 's')) el.setAttribute('s', attr(rowEl, 's'));
    // (`spans`, "1:6", is only a hint for loading; drop it rather than leave it wrong.)
    const spans = /^(\d+):(\d+)$/.exec(attr(rowEl, 'spans') || '');
    if (spans && (c + 1 < Number(spans[1]) || c + 1 > Number(spans[2]))) rowEl.removeAttribute('spans');
    return el;
  };
  for (const cEl of kids(rowEl, 'c', ns)) {
    const n = parseRef(attr(cEl, 'r')).c;
    if (n === c) return cEl;
    if (n > c) return rowEl.insertBefore(cell(), cEl);
  }
  return rowEl.appendChild(cell());
}

/** Grow the sheet's `dimension` (its used range) to take in a cell; the rest of the element is left alone. */
function growDimension(root, ns, r, c) {
  const dim = kid(root, 'dimension', ns);
  const range = dim && parseRange(attr(dim, 'ref'));
  if (!range) return;
  const grown = { r0: Math.min(range.r0, r), c0: Math.min(range.c0, c), r1: Math.max(range.r1, r), c1: Math.max(range.c1, c) };
  if (grown.r0 !== range.r0 || grown.c0 !== range.c0 || grown.r1 !== range.r1 || grown.c1 !== range.c1) dim.setAttribute('ref', rangeName(grown));
}

/** Set a cell to a number, a boolean, a string (written inline, so the shared string table isn't rewritten) or
 *  null (emptied; its style stays). Formula cells are refused until formulas and the calculation chain are modelled:
 *  dropping a formula leaves calcChain.xml pointing at a plain cell, and a shared formula's other cells depend on it. */
export function setCellXml(doc, r, c, value) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const sheetData = kid(root, 'sheetData', ns);
  if (!sheetData) throw new Error('worksheet without sheetData');
  pinRefs(sheetData, ns);
  const cEl = cellElement(doc, rowElement(doc, sheetData, ns, r), ns, r, c);
  if (kid(cEl, 'f', ns)) throw new Error(`${refName(r, c)} holds a formula; changing formula cells is not supported yet`);
  for (const k of [...cEl.childNodes]) cEl.removeChild(k);
  cEl.removeAttribute('t');
  const add = (parent, name, text) => {
    const el = parent.appendChild(make(doc, ns, name));
    if (text != null) el.appendChild(doc.createTextNode(text));
    return el;
  };
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new Error('not a finite number');
    add(cEl, 'v', String(value));
  } else if (typeof value === 'boolean') {
    cEl.setAttribute('t', 'b');
    add(cEl, 'v', value ? '1' : '0');
  } else if (value != null) {
    cEl.setAttribute('t', 'inlineStr');
    const t = add(add(cEl, 'is'), 't', String(value));
    // (Leading or trailing spaces, tabs and line breaks are kept only with xml:space="preserve".)
    if (/^\s|\s$|[\t\n]/.test(String(value))) t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
  }
  growDimension(root, ns, r, c);
}
