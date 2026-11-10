// Changing cells in a worksheet part's XML in place. Only the `row` and `c` elements an edit needs are touched;
// everything else in the part (and the other cells' markup) stays as it came.
import { kids, kid, attr } from './xml.js';
import { parseRef, refName, parseRange, rangeName } from './refs.js';

const isOn = (v) => v === '1' || v === 'true';
/** A new element in the part's namespace, under the root's prefix when the part uses one (`x:c`). */
const make = (doc, ns, name) => doc.createElementNS(ns, doc.documentElement.prefix ? `${doc.documentElement.prefix}:${name}` : name);

/** Give every row and cell of sheetData an explicit `r` where the file left it implied, so inserting one doesn't move
 *  the ones after it. (Allowed by the schema and kept by Excel; changes nothing a reader sees.) */
export function pinRefs(sheetData, ns) {
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

// The row elements of each edited part by row, made at its first edit (after pinRefs): a recalculation writes many
// cells, and finding each one's row by walking sheetData would cost a pass over the sheet per cell.
const rowIndexes = new WeakMap();
/** Forget a part's row index (its rows moved). */
export const forgetRows = (doc) => rowIndexes.delete(doc);

/** The `row` element for row r, made (in order) if there isn't one. */
function rowElement(doc, sheetData, ns, r) {
  let index = rowIndexes.get(doc);
  if (!index) {
    pinRefs(sheetData, ns);
    index = new Map(kids(sheetData, 'row', ns).map((el) => [Number(attr(el, 'r')) - 1, el]));
    rowIndexes.set(doc, index);
  }
  if (index.has(r)) return index.get(r);
  const el = newRow(doc, ns, r);
  const after = kids(sheetData, 'row', ns).find((x) => Number(attr(x, 'r')) - 1 > r);
  if (after) sheetData.insertBefore(el, after);
  else sheetData.appendChild(el);
  index.set(r, el);
  return el;
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

/** The `c` element of a cell, made (with its row) if there isn't one. */
function cellEl(doc, r, c) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const sheetData = kid(root, 'sheetData', ns);
  if (!sheetData) throw new Error('worksheet without sheetData');
  return cellElement(doc, rowElement(doc, sheetData, ns, r), ns, r, c);
}
const adder = (doc, ns) => (parent, name, text) => {
  const el = parent.appendChild(make(doc, ns, name));
  if (text != null) el.appendChild(doc.createTextNode(text));
  return el;
};

/** Set a cell to a number, a boolean, a string (written inline, so the shared string table isn't rewritten), an error
 *  ({ code }) or null (emptied; its style stays). A formula the cell held goes (the caller takes it out of a shared
 *  formula's group first, and out of calcChain.xml). Returns the cell's element. */
export function setCellXml(doc, r, c, value) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const cEl = cellEl(doc, r, c);
  for (const k of [...cEl.childNodes]) cEl.removeChild(k);
  cEl.removeAttribute('t');
  const add = adder(doc, ns);
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new Error('not a finite number');
    add(cEl, 'v', String(value));
  } else if (typeof value === 'boolean') {
    cEl.setAttribute('t', 'b');
    add(cEl, 'v', value ? '1' : '0');
  } else if (value != null && typeof value === 'object') {
    // (An error value, { code: '#DIV/0!' }, as Excel writes a constant one.)
    cEl.setAttribute('t', 'e');
    add(cEl, 'v', value.code);
  } else if (value != null) {
    cEl.setAttribute('t', 'inlineStr');
    const t = add(add(cEl, 'is'), 't', String(value));
    // (Leading or trailing spaces, tabs and line breaks are kept only with xml:space="preserve".)
    if (/^\s|\s$|[\t\n]/.test(String(value))) t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
  }
  growDimension(root, ns, r, c);
  return cEl;
}

/** Give a cell cell format `s` (0, the default, is written by leaving `s` out, as Excel does). */
export function setStyleXml(doc, r, c, s) {
  const root = doc.documentElement;
  const cEl = cellEl(doc, r, c);
  if (s) cEl.setAttribute('s', String(s));
  else cEl.removeAttribute('s');
  growDimension(root, root.namespaceURI, r, c);
  return cEl;
}

/** The elements that follow mergeCells in a worksheet (CT_Worksheet's order), for placing it when it is new. */
const AFTER_MERGES = ['phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup',
  'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing',
  'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

/** Write the sheet's merged ranges ([{ r0, c0, r1, c1 }]): its mergeCells element made, rewritten or taken out. */
export function setMergesXml(doc, merges) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  let el = kid(root, 'mergeCells', ns);
  if (!merges.length) { if (el) root.removeChild(el); return; }
  if (!el) {
    el = make(doc, ns, 'mergeCells');
    const next = [...root.childNodes].find((n) => n.nodeType === 1 && AFTER_MERGES.includes(n.localName));
    root.insertBefore(el, next || null);
  }
  while (el.firstChild) el.removeChild(el.firstChild);
  for (const m of merges) {
    const mc = make(doc, ns, 'mergeCell');
    mc.setAttribute('ref', rangeName(m));
    el.appendChild(mc);
  }
  el.setAttribute('count', String(merges.length));
}

/** Give a cell a formula (its text as the file has it, without "="); its value comes with setResultXml. */
export function setFormulaXml(doc, r, c, formula) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const cEl = cellEl(doc, r, c);
  for (const k of [...cEl.childNodes]) cEl.removeChild(k);
  cEl.removeAttribute('t');
  adder(doc, ns)(cEl, 'f', formula);
  growDimension(root, ns, r, c);
  return cEl;
}

/** Write a formula's result into its cell, keeping the formula: a number, text (t="str"), TRUE or FALSE, or an error
 *  ({ code }). A cell of an array formula's range that has no element yet gets one. */
export function setResultXml(doc, r, c, value) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  const cEl = cellEl(doc, r, c);
  for (const k of [...cEl.childNodes]) if (k.nodeType === 1 && (k.localName === 'v' || k.localName === 'is')) cEl.removeChild(k);
  const add = adder(doc, ns);
  if (typeof value === 'number') { cEl.removeAttribute('t'); add(cEl, 'v', String(value)); }
  else if (typeof value === 'boolean') { cEl.setAttribute('t', 'b'); add(cEl, 'v', value ? '1' : '0'); }
  else if (typeof value === 'string') {
    cEl.setAttribute('t', 'str');
    const v = add(cEl, 'v', value);
    if (/^\s|\s$|[\t\n]/.test(value)) v.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
  } else { cEl.setAttribute('t', 'e'); add(cEl, 'v', value.code); }
  return cEl;
}

/** Write out every cell of shared formula `si` with its own text (`texts`: "r:c" -> text), so that one of them can
 *  change without the others, which read the group's first cell. */
export function unshareXml(doc, si, texts) {
  const root = doc.documentElement;
  const ns = root.namespaceURI;
  for (const rowEl of kids(kid(root, 'sheetData', ns), 'row', ns)) {
    const r = Number(attr(rowEl, 'r')) - 1;
    for (const cEl of kids(rowEl, 'c', ns)) {
      const f = kid(cEl, 'f', ns);
      if (!f || attr(f, 't') !== 'shared' || attr(f, 'si') !== String(si)) continue;
      const text = texts.get(`${r}:${parseRef(attr(cEl, 'r')).c}`);
      if (text == null) continue;
      for (const name of ['t', 'si', 'ref']) f.removeAttribute(name);
      while (f.firstChild) f.removeChild(f.firstChild);
      f.appendChild(doc.createTextNode(text));
    }
  }
}
