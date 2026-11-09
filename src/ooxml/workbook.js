// An .xlsx workbook: its sheets in order, the shared strings, and each sheet's model read when first asked for.
// Saving rewrites only the sheets that were changed; every other part goes back byte for byte.
import { OpcPackage } from './package.js';
import { kids, kid, attr, serializeXml, NS } from './xml.js';
import { readSharedStrings } from './strings.js';
import { readSheet, readCell } from './sheet.js';
import { readStyles } from './styles.js';
import { readTheme } from './theme.js';
import { setCellXml, setFormulaXml, setResultXml, unshareXml } from './sheetedit.js';
import { parseRef } from './refs.js';
import { Calc } from '../calc/engine.js';
import { parseFormula } from '../calc/parse.js';

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
    /** Defined names: { name, sheet (the index of the sheet it belongs to, or null for the whole workbook), text }. */
    this.definedNames = kids(kid(root, 'definedNames', this.ns), 'definedName', this.ns).map((el) => ({
      name: attr(el, 'name'),
      sheet: attr(el, 'localSheetId') == null ? null : Number(attr(el, 'localSheetId')),
      text: el.textContent,
    }));
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

  /** The index of a worksheet by name (case ignored, as in a formula), or -1. */
  sheetIndexOf(name) {
    const n = String(name).toLowerCase();
    return this.sheets.findIndex((x) => x.kind === 'worksheet' && x.path && x.name.toLowerCase() === n);
  }

  /** A defined name as a formula sees it from sheet si: the sheet's own name first, then the workbook's. */
  definedName(name, si) {
    const n = name.toLowerCase();
    return this.definedNames.find((d) => d.sheet === si && d.name.toLowerCase() === n) ||
      this.definedNames.find((d) => d.sheet == null && d.name.toLowerCase() === n) || null;
  }

  /** The calculation engine (calc/engine.js), made at the first change. */
  get calc() {
    if (!this._calc) this._calc = new Calc(this);
    return this._calc;
  }

  /** Set one cell ("B2", or { r, c } 0-based) of sheet i; see setCellXml for what a value can be. */
  setCell(i, ref, value) {
    return this.apply(i, [{ ...this._at(ref), value }]);
  }

  /** Give a cell a formula, as the file has it ("SUM(A1:A3)", without "="). */
  setFormula(i, ref, formula) {
    return this.apply(i, [{ ...this._at(ref), formula }]);
  }

  _at(ref) {
    const at = typeof ref === 'string' ? parseRef(ref) : ref;
    if (!at) throw new Error(`not a cell reference: ${ref}`);
    return { r: at.r, c: at.c };
  }

  /** Change cells of sheet i, [{ r, c, value } or { r, c, formula }], then work out again the formulas that read
   *  them. A cell of an array formula is refused, as Excel refuses changing part of an array. Returns { circular }:
   *  formulas that read themselves (left as they were). */
  apply(i, changes) {
    const part = this._part(i);
    const model = part.model;
    const ns = part.doc.documentElement.namespaceURI;
    for (const ch of changes) {
      const inArray = model.arrays.find((a) => ch.r >= a.r0 && ch.r <= a.r1 && ch.c >= a.c0 && ch.c <= a.c1);
      if (inArray) throw new Error("You can't change part of an array");
      if (ch.formula != null) parseFormula(ch.formula); // (throws a FormulaError before anything changes)
    }
    let formulas = false;
    for (const { r, c, value, formula } of changes) {
      const old = model.cells.get(r)?.get(c);
      if (old && old.shared != null) this._unshare(i, old.shared);
      if (old && old.f != null) { this.calc.remove(i, r, c); formulas = true; }
      const el = formula != null ? setFormulaXml(part.doc, r, c, formula) : setCellXml(part.doc, r, c, value);
      // Only this cell of the model changes (re-reading the sheet would cost a pass over all its cells per keystroke).
      if (!model.cells.has(r)) model.cells.set(r, new Map());
      const cell = readCell(el, ns, this.sharedStrings);
      model.cells.get(r).set(c, cell);
      model.maxRow = Math.max(model.maxRow, r);
      model.maxCol = Math.max(model.maxCol, c);
      if (formula != null) { this.calc.add(i, r, c, cell); formulas = true; }
    }
    this.dirty.add(this.sheets[i].path);
    this.edited = true;
    // (calcChain.xml lists the formula cells in the order Excel last worked them out; one naming a cell that no longer
    // holds a formula makes Excel repair the file. Without it Excel works the order out again.)
    if (formulas) this._dropCalcChain();
    const rects = changes.map(({ r, c }) => ({ si: i, r0: r, c0: c, r1: r, c1: c }));
    return this.calc.recalc(rects, changes.filter((ch) => ch.formula != null).map(({ r, c }) => ({ si: i, r, c })));
  }

  /** Work out every formula of the workbook again (Excel's Ctrl+Alt+F9). */
  calculate() {
    return this.calc.calculateAll();
  }

  /** Write a formula's result into a cell (from the engine); nothing changes if it is the value the cell has. */
  writeResult(i, r, c, value) {
    const part = this._part(i);
    const model = part.model;
    const cell = model.cells.get(r)?.get(c);
    const t = typeof value === 'number' ? 'n' : typeof value === 'boolean' ? 'b' : typeof value === 'string' ? 'str' : 'e';
    const v = t === 'e' ? value.code : value;
    if (cell && cell.t === t && cell.v === v) return;
    setResultXml(part.doc, r, c, t === 'e' ? value : v);
    if (cell) { cell.t = t; cell.v = v; }
    else {
      if (!model.cells.has(r)) model.cells.set(r, new Map());
      model.cells.get(r).set(c, { t, v, s: 0, f: null });
    }
    model.maxRow = Math.max(model.maxRow, r);
    model.maxCol = Math.max(model.maxCol, c);
    this.dirty.add(this.sheets[i].path);
  }

  /** Write each cell of shared formula `si` on sheet i with its own formula. */
  _unshare(i, si) {
    const model = this._part(i).model;
    const texts = new Map();
    for (const [r, line] of model.cells) for (const [c, cell] of line) if (cell.shared === si) { texts.set(`${r}:${c}`, cell.f); delete cell.shared; }
    unshareXml(this._part(i).doc, si, texts);
  }

  _dropCalcChain() {
    const rel = this.pkg.rels.find((x) => x.type.endsWith('/calcChain') && !x.external);
    if (rel) this.pkg.removePart(rel.target);
  }

  /** The workbook as .xlsx bytes: changed sheets serialised, every other part as it was read. */
  toBytes() {
    // (A formula the engine couldn't work out may read a changed cell: have Excel recalculate the workbook on opening.)
    if (this.edited && this._calc && this._calc.incomplete) this._fullCalcOnLoad();
    for (const [i, part] of this._parsed) {
      const path = this.sheets[i].path;
      if (this.dirty.has(path)) this.pkg.setText(path, partText(part.doc, this.pkg.text(path)));
    }
    return this.pkg.toBytes();
  }

  _fullCalcOnLoad() {
    const path = this.pkg.mainPath;
    const original = this.pkg.text(path);
    const doc = this.pkg.xml(path);
    const root = doc.documentElement;
    let pr = kid(root, 'calcPr', this.ns);
    if (pr && attr(pr, 'fullCalcOnLoad') === '1') return;
    if (!pr) {
      pr = doc.createElementNS(this.ns, root.prefix ? `${root.prefix}:calcPr` : 'calcPr');
      // (calcPr comes after definedNames and before oleSize, the custom views, pivotCaches and the rest.)
      const AFTER = ['oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'];
      const next = [...root.childNodes].find((n) => n.nodeType === 1 && AFTER.includes(n.localName));
      root.insertBefore(pr, next || null);
    }
    pr.setAttribute('fullCalcOnLoad', '1');
    this.pkg.setText(path, partText(doc, original));
  }
}
