// An .xlsx workbook: its sheets in order, the shared strings, and each sheet's model read when first asked for.
// Saving rewrites only the sheets that were changed; every other part goes back byte for byte.
import { OpcPackage } from './package.js';
import { kids, kid, attr, serializeXml, NS } from './xml.js';
import { readSharedStrings } from './strings.js';
import { readSheet } from './sheet.js';
import { readStyles } from './styles.js';
import { readTheme } from './theme.js';
import { setCellXml } from './sheetedit.js';
import { parseRef } from './refs.js';

/** A part's XML text with the original's XML declaration in front (serialisers drop or rewrite it, and Excel writes
 *  `standalone="yes"`). */
function partText(doc, original) {
  const text = original && original.charCodeAt(0) === 0xfeff ? original.slice(1) : original || ''; // (a byte order mark)
  const decl = /^(<\?xml[^?]*\?>)/.exec(text);
  const body = serializeXml(doc).replace(/^<\?xml[^?]*\?>\s*/, '');
  return decl ? `${decl[1]}\r\n${body}` : body;
}

export class Workbook {
  /** @param {OpcPackage} pkg */
  constructor(pkg) {
    this.pkg = pkg;
    const doc = pkg.xml(pkg.mainPath);
    const root = doc.documentElement;
    this.ns = root.namespaceURI;
    if (this.ns !== NS.x && this.ns !== NS.xStrict) throw new Error('Not a workbook: the main part is ' + root.localName);
    const pr = kid(root, 'workbookPr', this.ns);
    /** Dates count from 1904-01-01 instead of 1900-01-00 (older Mac workbooks). */
    this.date1904 = !!pr && (attr(pr, 'date1904') === '1' || attr(pr, 'date1904') === 'true');
    const rels = pkg.rels;
    this.sheets = kids(kid(root, 'sheets', this.ns), 'sheet', this.ns).map((el) => {
      // (The relationship id is r:id in Transitional, and in Strict's own relationships namespace.)
      const rId = [...el.attributes].find((a) => a.localName === 'id' && a.namespaceURI !== null)?.value || null;
      const rel = rId ? rels.find((x) => x.id === rId) : null;
      return {
        name: attr(el, 'name'),
        sheetId: Number(attr(el, 'sheetId')),
        state: attr(el, 'state') || 'visible',
        rId,
        path: rel && !rel.external ? rel.target : null,
        kind: rel ? rel.type.slice(rel.type.lastIndexOf('/') + 1) : null, // worksheet, chartsheet, dialogsheet...
      };
    });
    const sstPath = rels.find((x) => x.type.endsWith('/sharedStrings'))?.target;
    this.sharedStrings = readSharedStrings(sstPath ? pkg.xml(sstPath) : null);
    const part = (kind) => rels.find((x) => x.type.endsWith('/' + kind) && !x.external && pkg.has(x.target))?.target;
    /** Fonts and cell formats (styles.js), and the theme's fonts and colours (theme.js; Excel's default without one). */
    this.styles = readStyles(part('styles') ? pkg.xml(part('styles')) : null);
    this.theme = readTheme(part('theme') ? pkg.xml(part('theme')) : null);
    /** Parsed sheets by index: { doc, model }. */
    this._parsed = new Map();
    /** Paths of the parts changed since opening (the only ones rewritten on save). */
    this.dirty = new Set();
  }

  static fromBytes(bytes) {
    const pkg = OpcPackage.fromBytes(bytes);
    if (!/\.xml$/.test(pkg.mainPath) || !pkg.has(pkg.mainPath)) throw new Error('Not a workbook: main part missing');
    return new Workbook(pkg);
  }

  /** The index of the sheet a workbook opens on: the active tab if it is a visible worksheet, else the first one. */
  get activeSheet() {
    const view = kid(kid(this.pkg.xml(this.pkg.mainPath).documentElement, 'bookViews', this.ns), 'workbookView', this.ns);
    const active = Number(attr(view, 'activeTab') || 0);
    const usable = (i) => this.sheets[i] && this.sheets[i].kind === 'worksheet' && this.sheets[i].state === 'visible' && this.sheets[i].path;
    if (usable(active)) return active;
    const i = this.sheets.findIndex((_, k) => usable(k));
    return i < 0 ? 0 : i;
  }

  _part(i) {
    if (!this._parsed.has(i)) {
      const info = this.sheets[i];
      if (!info) throw new Error(`no sheet ${i}`);
      if (info.kind !== 'worksheet' || !info.path || !this.pkg.has(info.path)) throw new Error(`sheet "${info.name}" is not a worksheet`);
      const doc = this.pkg.xml(info.path);
      this._parsed.set(i, { doc, model: readSheet(doc, this.sharedStrings) });
    }
    return this._parsed.get(i);
  }

  /** The model of sheet i (see readSheet). */
  sheet(i) {
    return this._part(i).model;
  }

  /** Set one cell ("B2", or { r, c } 0-based) of sheet i; see setCellXml for what a value can be. */
  setCell(i, ref, value) {
    const at = typeof ref === 'string' ? parseRef(ref) : ref;
    if (!at) throw new Error(`not a cell reference: ${ref}`);
    const part = this._part(i);
    setCellXml(part.doc, at.r, at.c, value);
    part.model = readSheet(part.doc, this.sharedStrings);
    this.dirty.add(this.sheets[i].path);
  }

  /** The workbook as .xlsx bytes: changed sheets serialised, every other part as it was read. */
  toBytes() {
    for (const [i, part] of this._parsed) {
      const path = this.sheets[i].path;
      if (this.dirty.has(path)) this.pkg.setText(path, partText(part.doc, this.pkg.text(path)));
    }
    return this.pkg.toBytes();
  }
}
