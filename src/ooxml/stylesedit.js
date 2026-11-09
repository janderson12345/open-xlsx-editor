// Changing cell formats as Excel writes them (docs/FORMATS.md): a cell format (cellXfs) is the cell's old one with a
// change, and the font, fill or border it points to is the old one changed the same way. One that matches an entry
// already there is reused; a new one goes at the end, so indexes cells already use never move. Percent, comma and
// currency are Excel's built-in cell styles (Percent, Comma, Currency), added when the workbook hasn't got them.
import { kids, kid, attr, serializeXml, NS, REL_TYPES } from './xml.js';
import { readStyles, readFont, readFill, readBorder, readXf } from './styles.js';
import { BUILTIN_FORMATS } from './format.js';

/** Excel's order of a font's children, and of a border's sides. */
const FONT_ORDER = ['b', 'i', 'strike', 'condense', 'extend', 'outline', 'shadow', 'u', 'vertAlign', 'sz', 'color', 'name', 'family', 'charset', 'scheme'];
const SIDE_ORDER = ['start', 'left', 'end', 'right', 'top', 'bottom', 'diagonal', 'vertical', 'horizontal'];
/** Excel's built-in cell styles the Number group applies: builtinId and number format. */
export const CELL_STYLES = { Comma: { builtinId: 3, numFmtId: 43 }, Currency: { builtinId: 4, numFmtId: 44 }, Percent: { builtinId: 5, numFmtId: 9 } };
/** Built-in formats whose codes depend on the computer's locale: Excel writes their codes out (numFmts) when it uses
 *  them, so another locale reads the same thing. */
const LOCALE_BUILTINS = new Set([5, 6, 7, 8, 41, 42, 43, 44]);

/** A minimal styles part, as a workbook without one gets when it is first formatted: Excel's defaults (Calibri 11, no
 *  fill and the gray125 fill Excel always keeps second, no border, the Normal style). */
const EMPTY_STYLES = `<styleSheet xmlns="${NS.x}"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

const make = (doc, name) => {
  const root = doc.documentElement;
  return doc.createElementNS(root.namespaceURI, root.prefix ? `${root.prefix}:${name}` : name);
};

/** A colour element's attributes from { rgb: 'FFRRGGBB' } or { theme, tint } or { indexed }. */
function setColor(el, color) {
  for (const a of ['rgb', 'theme', 'tint', 'indexed', 'auto']) el.removeAttribute(a);
  if (color.rgb) el.setAttribute('rgb', color.rgb.length === 6 ? `FF${color.rgb}` : color.rgb);
  else if (color.theme != null) {
    el.setAttribute('theme', String(color.theme));
    if (color.tint) el.setAttribute('tint', String(color.tint));
  } else if (color.indexed != null) el.setAttribute('indexed', String(color.indexed));
  return el;
}

/** Put `child` among `parent`'s children in `order` (replacing one of the same name). */
function placeChild(parent, child, order) {
  const name = child.localName;
  const old = [...parent.childNodes].find((n) => n.nodeType === 1 && n.localName === name);
  if (old === child) return child;
  if (old) return parent.replaceChild(child, old);
  const after = order.indexOf(name);
  const next = [...parent.childNodes].find((n) => n.nodeType === 1 && order.indexOf(n.localName) > after);
  parent.insertBefore(child, next || null);
}
const removeChild = (parent, name) => {
  for (const n of [...parent.childNodes]) if (n.nodeType === 1 && n.localName === name) parent.removeChild(n);
};

/** The styles part of a workbook, made when it has none. */
export function stylesPart(wb) {
  if (wb.stylesDoc) return wb.stylesDoc;
  const path = 'xl/styles.xml';
  wb.pkg.setText(path, EMPTY_STYLES);
  wb.pkg.addOverride(path, 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml');
  wb.pkg.addRel(REL_TYPES.styles, 'styles.xml');
  wb.stylesPath = path;
  wb.stylesDoc = wb.pkg.xml(path);
  // (The model now reads the new part: the same defaults, with its second fill.)
  Object.assign(wb.styles, readStyles(wb.stylesDoc), { palette: wb.styles.palette });
  return wb.stylesDoc;
}

/**
 * Writes changed cell formats into a workbook's styles part and keeps `wb.styles` (styles.js) in step. Made once per
 * workbook (`formatsOf`).
 */
class Formats {
  constructor(wb) {
    this.wb = wb;
    this.doc = stylesPart(wb);
    this.root = this.doc.documentElement;
    this.ns = this.root.namespaceURI;
    // (Each list's entries as text, to find one that matches.)
    this.seen = {};
    for (const [list, item] of [['fonts', 'font'], ['fills', 'fill'], ['borders', 'border'], ['cellXfs', 'xf'], ['cellStyleXfs', 'xf']]) {
      this.seen[list] = new Map(this.items(list, item).map((el, i) => [serializeXml(el), i]).reverse());
    }
  }

  list(name) {
    let el = kid(this.root, name, this.ns);
    if (!el) {
      el = make(this.doc, name);
      const ORDER = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst'];
      placeChild(this.root, el, ORDER);
    }
    return el;
  }

  items(list, item) { return kids(kid(this.root, list, this.ns), item, this.ns); }

  /** The index of `el` in `list`: an equal entry's, or a new one at the end. */
  add(list, el, read) {
    const key = serializeXml(el);
    const seen = this.seen[list];
    if (seen.has(key)) return seen.get(key);
    const parent = this.list(list);
    parent.appendChild(el);
    const i = this.items(list, el.localName).length - 1;
    parent.setAttribute('count', String(i + 1));
    seen.set(key, i);
    if (read) read(el, i);
    this.wb.dirty.add(this.wb.stylesPath);
    return i;
  }

  /** A copy of entry i of a list (the first when there's none). */
  copy(list, item, i) {
    const el = this.items(list, item)[i] || this.items(list, item)[0];
    return el ? el.cloneNode(true) : make(this.doc, item);
  }

  /** A font changed by { b, i, strike, u ('single', 'double' or null), size, name, color ({ rgb } / { theme, tint },
   *  null: automatic) }; returns its index. */
  font(fontId, change) {
    const el = this.copy('fonts', 'font', fontId);
    const toggle = (name, on) => (on ? placeChild(el, make(this.doc, name), FONT_ORDER) : removeChild(el, name));
    if ('b' in change) toggle('b', change.b);
    if ('i' in change) toggle('i', change.i);
    if ('strike' in change) toggle('strike', change.strike);
    if ('u' in change) {
      removeChild(el, 'u');
      if (change.u) {
        const u = make(this.doc, 'u');
        if (change.u !== 'single') u.setAttribute('val', change.u);
        placeChild(el, u, FONT_ORDER);
      }
    }
    if (change.size != null) {
      const sz = make(this.doc, 'sz');
      sz.setAttribute('val', String(change.size));
      placeChild(el, sz, FONT_ORDER);
    }
    if (change.name != null) {
      const name = make(this.doc, 'name');
      name.setAttribute('val', change.name);
      placeChild(el, name, FONT_ORDER);
      // (A font chosen by name is no longer the theme's: Excel drops the scheme, and the family it doesn't know.)
      removeChild(el, 'scheme');
      removeChild(el, 'family');
      removeChild(el, 'charset');
    }
    if ('color' in change) {
      if (change.color) placeChild(el, setColor(make(this.doc, 'color'), change.color), FONT_ORDER);
      else placeChild(el, setColor(make(this.doc, 'color'), { theme: 1 }), FONT_ORDER);
    }
    return this.add('fonts', el, (x, i) => { this.wb.styles.fonts[i] = readFont(x, this.ns); });
  }

  /** A solid fill of a colour, or none (null: fill 0, which every workbook has as "none"). */
  fill(color) {
    if (!color) return 0;
    const el = make(this.doc, 'fill');
    const p = el.appendChild(make(this.doc, 'patternFill'));
    p.setAttribute('patternType', 'solid');
    p.appendChild(setColor(make(this.doc, 'fgColor'), color));
    p.appendChild(setColor(make(this.doc, 'bgColor'), { indexed: 64 }));
    return this.add('fills', el, (x, i) => { this.wb.styles.fills[i] = readFill(x, this.ns); });
  }

  /** A border with some sides changed: { left, right, top, bottom }, each { style, color } or null (no line). */
  border(borderId, sides) {
    const el = this.copy('borders', 'border', borderId);
    for (const [name, side] of Object.entries(sides)) {
      // (A right-to-left file may name the sides start and end.)
      if (name === 'left') removeChild(el, 'start');
      if (name === 'right') removeChild(el, 'end');
      const s = make(this.doc, name);
      if (side) {
        s.setAttribute('style', side.style);
        s.appendChild(setColor(make(this.doc, 'color'), side.color || { indexed: 64 }));
      }
      placeChild(el, s, SIDE_ORDER);
    }
    for (const name of ['left', 'right', 'top', 'bottom', 'diagonal']) if (!kid(el, name, this.ns) && !(name === 'left' && kid(el, 'start', this.ns)) && !(name === 'right' && kid(el, 'end', this.ns))) placeChild(el, make(this.doc, name), SIDE_ORDER);
    return this.add('borders', el, (x, i) => { this.wb.styles.borders[i] = readBorder(x, this.ns); });
  }

  /** The id of a number format code: a built-in one's, the workbook's own, or a new one (164 up, as Excel numbers
   *  them). */
  numFmt(code) {
    const styles = this.wb.styles;
    for (const [id, c] of Object.entries(styles.numFmts)) if (c === code) return Number(id);
    for (const [id, c] of Object.entries(BUILTIN_FORMATS)) if (c === code && !LOCALE_BUILTINS.has(Number(id))) return Number(id);
    const builtin = Object.entries(BUILTIN_FORMATS).find(([, c]) => c === code);
    const id = builtin ? Number(builtin[0]) : Math.max(163, ...Object.keys(styles.numFmts).map(Number)) + 1;
    this.writeNumFmt(id, code);
    return id;
  }

  writeNumFmt(id, code) {
    const el = make(this.doc, 'numFmt');
    el.setAttribute('numFmtId', String(id));
    el.setAttribute('formatCode', code);
    const parent = this.list('numFmts');
    parent.appendChild(el);
    parent.setAttribute('count', String(kids(parent, 'numFmt', this.ns).length));
    this.wb.styles.numFmts[id] = code;
    this.wb.dirty.add(this.wb.stylesPath);
  }

  /** The cellStyleXfs index of one of Excel's built-in cell styles (CELL_STYLES), added with its cellStyle entry when
   *  the workbook hasn't got it. */
  cellStyle(name) {
    const { builtinId, numFmtId } = CELL_STYLES[name];
    const found = kids(kid(this.root, 'cellStyles', this.ns), 'cellStyle', this.ns).find((el) => attr(el, 'builtinId') === String(builtinId));
    if (found) return Number(attr(found, 'xfId') || 0);
    if (LOCALE_BUILTINS.has(numFmtId) && this.wb.styles.numFmts[numFmtId] == null) this.writeNumFmt(numFmtId, BUILTIN_FORMATS[numFmtId]);
    // (As Excel writes a style's own format: the Normal font, no fill or border, and only its number format applied.)
    const xf = make(this.doc, 'xf');
    for (const [k, v] of [['numFmtId', numFmtId], ['fontId', 0], ['fillId', 0], ['borderId', 0], ['applyFont', 0], ['applyFill', 0], ['applyBorder', 0], ['applyAlignment', 0], ['applyProtection', 0]]) xf.setAttribute(k, String(v));
    const parent = this.list('cellStyleXfs');
    parent.appendChild(xf);
    const xfId = kids(parent, 'xf', this.ns).length - 1;
    parent.setAttribute('count', String(xfId + 1));
    const styles = this.list('cellStyles');
    const cs = make(this.doc, 'cellStyle');
    cs.setAttribute('name', name);
    cs.setAttribute('xfId', String(xfId));
    cs.setAttribute('builtinId', String(builtinId));
    styles.appendChild(cs);
    styles.setAttribute('count', String(kids(styles, 'cellStyle', this.ns).length));
    this.wb.dirty.add(this.wb.stylesPath);
    return xfId;
  }

  /**
   * Cell format `s` with a change: { font: {...} (see font()), fill: colour or null, border: { side: ... } (see
   * border()), align: { horizontal, vertical, wrap, indent }, numFmt: code, style: 'Percent' | 'Comma' | 'Currency' }.
   * Returns the index of the cell format.
   */
  xf(s, change) {
    const styles = this.wb.styles;
    const old = styles.cellXfs[s] || styles.cellXfs[0];
    const el = this.copy('cellXfs', 'xf', s);
    const set = (k, v, apply) => { el.setAttribute(k, String(v)); if (apply) el.setAttribute(apply, '1'); };
    if (change.font) set('fontId', this.font(old.fontId, change.font), 'applyFont');
    if ('fill' in change) set('fillId', this.fill(change.fill), 'applyFill');
    if (change.border) set('borderId', this.border(old.borderId, change.border), 'applyBorder');
    if (change.style) {
      set('xfId', this.cellStyle(change.style));
      set('numFmtId', CELL_STYLES[change.style].numFmtId, 'applyNumberFormat');
    } else if (change.numFmt != null) set('numFmtId', this.numFmt(change.numFmt), 'applyNumberFormat');
    if (change.align) {
      const a = kid(el, 'alignment', this.ns) || make(this.doc, 'alignment');
      const { horizontal, vertical, wrap, indent } = change.align;
      const put = (k, v, dflt) => (v == null || v === dflt ? a.removeAttribute(k) : a.setAttribute(k, String(v)));
      if (horizontal !== undefined) put('horizontal', horizontal, 'general');
      if (vertical !== undefined) put('vertical', vertical, 'bottom');
      if (wrap !== undefined) put('wrapText', wrap ? 1 : null, null);
      if (indent !== undefined) put('indent', indent || null, null);
      if (a.attributes.length) placeChild(el, a, ['alignment', 'protection', 'extLst']);
      else removeChild(el, 'alignment');
      el.setAttribute('applyAlignment', '1');
    }
    if (!el.getAttribute('xfId')) el.setAttribute('xfId', '0');
    return this.add('cellXfs', el, (x, i) => { styles.cellXfs[i] = readXf(x, this.ns, styles); });
  }
}

const formats = new WeakMap();
/** The format writer of a workbook. */
export function formatsOf(wb) {
  if (!formats.has(wb)) formats.set(wb, new Formats(wb));
  return formats.get(wb);
}

