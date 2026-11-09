// Small XML helpers shared by the reader and writer.
// The DOMParser / XMLSerializer implementations are injectable so the same code
// runs in the browser (native) and in Node (@xmldom/xmldom) for tests.

export const NS = {
  // SpreadsheetML, the default for the lookups below.
  x: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  // (The same vocabulary in Strict OOXML, ISO/IEC 29500 Strict; read by taking the root's namespace.)
  xStrict: 'http://purl.oclc.org/ooxml/spreadsheetml/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
};

export const REL_TYPES = {
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  worksheet: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
  chartsheet: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet',
  sharedStrings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  calcChain: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain',
  drawing: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
  comments: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
  table: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
};

let impl = {
  DOMParser: typeof DOMParser !== 'undefined' ? DOMParser : null,
  XMLSerializer: typeof XMLSerializer !== 'undefined' ? XMLSerializer : null,
};

/** Override the XML implementation (used by Node tests). */
export function setXmlImpl({ DOMParser: P, XMLSerializer: S }) {
  impl = { DOMParser: P, XMLSerializer: S };
}

export function parseXml(text) {
  const doc = new impl.DOMParser().parseFromString(text, 'application/xml');
  const err = doc.getElementsByTagName('parsererror')[0];
  if (err) throw new Error('Invalid XML: ' + err.textContent);
  return doc;
}

export function serializeXml(node) {
  return new impl.XMLSerializer().serializeToString(node);
}

export const isEl = (n) => n && n.nodeType === 1;

/** Element children, optionally filtered by local name in a namespace (SpreadsheetML by default). */
export function kids(el, name, ns = NS.x) {
  const out = [];
  if (!el) return out;
  for (let c = el.firstChild; c; c = c.nextSibling) {
    if (!isEl(c)) continue;
    if (name && (c.localName !== name || c.namespaceURI !== ns)) continue;
    out.push(c);
  }
  return out;
}

export const kid = (el, name, ns = NS.x) => kids(el, name, ns)[0] || null;

/** Read an attribute in a namespace, falling back to the unprefixed name (SpreadsheetML's attributes are unprefixed). */
export function attr(el, name, ns = NS.x) {
  if (!el) return null;
  if (el.hasAttributeNS && el.hasAttributeNS(ns, name)) return el.getAttributeNS(ns, name);
  return el.hasAttribute && el.hasAttribute(name) ? el.getAttribute(name) : null;
}

/** OOXML on/off properties: <b/> is on, <b val="0|false"/> is off. */
export function onOff(el) {
  if (!el) return null;
  const v = attr(el, 'val');
  return !(v === '0' || v === 'false' || v === 'off');
}

export function esc(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Serialize an element and drop xmlns declarations that the document root already makes. */
export function outerXml(el, rootNs) {
  let s = serializeXml(el);
  s = s.replace(/\s+xmlns(?::([\w.-]+))?="([^"]*)"/g, (m, pfx, uri) =>
    rootNs[pfx || ''] === uri ? '' : m,
  );
  return s;
}

/** Map of prefix -> namespace URI declared on an element. */
export function nsDecls(el) {
  const out = {};
  for (const a of Array.from(el.attributes)) {
    if (a.name === 'xmlns') out[''] = a.value;
    else if (a.name.startsWith('xmlns:')) out[a.name.slice(6)] = a.value;
  }
  return out;
}

export const textOf = (el) => (el ? el.textContent || '' : '');

/** An element's child elements, whatever their namespace. */
export const elementKids = (el) => (el ? [...el.childNodes].filter((n) => n.nodeType === 1) : []);
/** Its first child element with this local name (any namespace), or null. */
export const childNamed = (el, name) => elementKids(el).find((n) => n.localName === name) || null;
/** Its child elements with this local name (any namespace). */
export const childrenNamed = (el, name) => elementKids(el).filter((n) => n.localName === name);

/** An attribute value's entities decoded (&quot; &apos; &lt; &gt; &amp;), as the XML parser would. */
export const unescapeXml = (s) => (s == null ? s : s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&'));

/** SpreadsheetML's escaped characters in strings (ST_Xstring): `_x000D_` is a carriage return, `_x005F_` a literal
 *  underscore so that `_x005F_x000D_` reads as the text `_x000D_`. */
export const decodeXstring = (s) => (s && s.includes('_x') ? s.replace(/_x([0-9A-Fa-f]{4})_/g, (m, h) => String.fromCharCode(parseInt(h, 16))) : s);
