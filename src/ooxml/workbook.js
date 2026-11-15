// An .xlsx workbook: its sheets in order, the shared strings, and each sheet's model read when first asked for.
// Saving rewrites only the sheets that were changed; every other part goes back byte for byte.
import { OpcPackage } from './package.js';
import { fileType, setFileType } from './filetype.js';
import { kids, kid, attr, serializeXml, parseXml, NS } from './xml.js';
import { readSharedStrings } from './strings.js';
import { readSheet, readCell } from './sheet.js';
import { readStyles } from './styles.js';
import { readTheme, resolveColor } from './theme.js';
import { setCellXml, setFormulaXml, setResultXml, setStyleXml, setMergesXml, unshareXml } from './sheetedit.js';
import { setDynamicXml, clearSpilledXml, RICH_ERRORS } from './spillxml.js';
import { recalcAfter } from './calcmode.js';
import { formatsOf } from './stylesedit.js';
import { parseRef } from './refs.js';
import { withEl, addSheet, renameSheet, moveSheet, setSheetHidden, setTabColor, deleteSheet } from './sheetops.js';
import { changeLines } from './rowcol.js';
import { setColumns, setRows, setLineStyles } from './sizes.js';
import { setFreeze, setZoom, setShowFormulas, setShown } from './view.js';
import { applyAutoFilter, filterName } from './filterops.js';
import { applyAdvancedFilter } from './advfilter.js';
import { Calc } from '../calc/engine.js';
import { parseFormula } from '../calc/parse.js';
import { isCompoundFile, decryptWorkbook, encryptWorkbook } from './encryption.js';
import { setValidation as setValidationXml, addToRule, moveValidation } from './dvedit.js';
import { addConditional, clearConditional } from './cfedit.js';

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
    /** The workbook part, parsed once; written back on save only when changed (`bookDirty`). */
    this.doc = doc;
    this.bookDirty = false;
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
      return withEl({
        name: attr(el, 'name'),
        sheetId: Number(attr(el, 'sheetId')),
        state: attr(el, 'state') || 'visible',
        rId,
        path: rel && !rel.external ? rel.target : null,
        kind: rel ? rel.type.slice(rel.type.lastIndexOf('/') + 1) : null, // worksheet, chartsheet, dialogsheet...
      }, el);
    });
    const sstPath = rels.find((x) => x.type.endsWith('/sharedStrings'))?.target;
    this.sharedStrings = readSharedStrings(sstPath ? pkg.xml(sstPath) : null);
    const part = (kind) => rels.find((x) => x.type.endsWith('/' + kind) && !x.external && pkg.has(x.target))?.target;
    /** Fonts and cell formats (styles.js), and the theme's fonts and colours (theme.js; Excel's default without one). */
    this.stylesPath = part('styles') || null;
    /** The styles part, parsed once; formatting cells adds to it (stylesedit.js). */
    this.stylesDoc = this.stylesPath ? pkg.xml(this.stylesPath) : null;
    this.styles = readStyles(this.stylesDoc);
    this.theme = readTheme(part('theme') ? pkg.xml(part('theme')) : null);
    /** Defined names: { name, sheet (the index of the sheet it belongs to, or null for the whole workbook), text }. */
    this.definedNames = kids(kid(root, 'definedNames', this.ns), 'definedName', this.ns).map((el) => withEl({
      name: attr(el, 'name'),
      sheet: attr(el, 'localSheetId') == null ? null : Number(attr(el, 'localSheetId')),
      text: el.textContent,
    }, el));
    /** Parsed sheets by their entry in `sheets` (indexes change when sheets are added, moved or deleted): { doc, model }. */
    this._parsed = new Map();
    /** The number the next new sheet tries (Sheet1, Sheet2...: Excel counts from 1 each time it opens a workbook). */
    this.nextSheetNumber = 1;
    /** Paths of the parts changed since opening (the only ones rewritten on save). */
    this.dirty = new Set();
  }

  /** Open a workbook's bytes. One saved with a password (encryption.js) needs `password`; without it, or with a
   *  wrong one, a PasswordError says so. It is saved again with the same password (`setPassword` changes it). */
  static fromBytes(bytes, { password = null } = {}) {
    let encryption = null;
    if (isCompoundFile(bytes)) {
      const { bytes: plain, streams } = decryptWorkbook(bytes, password);
      encryption = { password, streams };
      bytes = plain;
    }
    const pkg = OpcPackage.fromBytes(bytes);
    if (!/\.xml$/.test(pkg.mainPath) || !pkg.has(pkg.mainPath)) throw new Error('Not a workbook: main part missing');
    const wb = new Workbook(pkg);
    wb.encryption = encryption;
    return wb;
  }

  /** The password the workbook is saved with: a new one (Encrypt with Password), or null to save it unencrypted. */
  setPassword(password) {
    this.encryption = password ? { password, streams: this.encryption?.streams ?? null } : null;
    this.edited = true;
  }

  /** Whether the workbook is saved with a password. */
  get encrypted() { return !!this.encryption; }

  /** The index of the sheet a workbook opens on: the active tab if it is a visible worksheet, else the first one. */
  get activeSheet() {
    const view = kid(kid(this.doc.documentElement, 'bookViews', this.ns), 'workbookView', this.ns);
    const active = Number(attr(view, 'activeTab') || 0);
    if (this.canShow(active)) return active;
    const i = this.sheets.findIndex((_, k) => this.canShow(k));
    return i < 0 ? 0 : i;
  }

  _part(i) {
    const info = this.sheets[i];
    if (!info) throw new Error(`no sheet ${i}`);
    if (!this._parsed.has(info)) {
      if (info.kind !== 'worksheet' || !info.path || !this.pkg.has(info.path)) throw new Error(`sheet "${info.name}" is not a worksheet`);
      const doc = this.pkg.xml(info.path);
      this._parsed.set(info, { doc, model: readSheet(doc, this.sharedStrings) });
    }
    return this._parsed.get(info);
  }

  /** The model of sheet i (see readSheet). */
  sheet(i) {
    return this._part(i).model;
  }

  /** Whether sheet i can be shown: a worksheet whose tab isn't hidden. */
  canShow(i) {
    const info = this.sheets[i];
    return !!info && info.kind === 'worksheet' && info.state === 'visible' && !!info.path && this.pkg.has(info.path);
  }

  /** Sheet i's tab colour as '#RRGGBB', or null. Its `sheetPr` comes first in the part, so only the part's start is
   *  read (a tab strip shouldn't parse every sheet). */
  tabColor(i) {
    const info = this.sheets[i];
    let at;
    if (this._parsed.has(info)) {
      // (Parsed, it may have been given a colour since.)
      const root = this._parsed.get(info).doc.documentElement;
      const el = kid(kid(root, 'sheetPr', root.namespaceURI), 'tabColor', root.namespaceURI);
      if (!el) return null;
      at = Object.fromEntries([...el.attributes].map((a) => [a.name, a.value]));
    } else {
      const bytes = info?.path && this.pkg.files[info.path];
      if (!bytes) return null;
      const head = new TextDecoder().decode(bytes.subarray(0, 8192)).split(/<(?:\w+:)?sheetData\b/)[0];
      const el = /<(?:\w+:)?tabColor\b([^>]*)>/.exec(head);
      if (!el) return null;
      at = Object.fromEntries([...el[1].matchAll(/([\w:]+)="([^"]*)"/g)].map((m) => [m[1], m[2]]));
    }
    const num = (name) => (at[name] != null ? Number(at[name]) : null);
    return resolveColor({ rgb: at.rgb || null, theme: num('theme'), indexed: num('indexed'), auto: at.auto === '1', tint: num('tint') || 0 }, this.theme, this.styles.palette);
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
    // (The formulas known before any is changed: one replaced here may have spilled into cells the file filled.)
    this.calc.build();
    for (const { r, c, value, formula, legacy } of changes) {
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
      // (A formula entered here is Excel 365's: an array result spills from it; spill.js.)
      if (formula != null) { cell.dynamic = !legacy; this.calc.add(i, r, c, cell); formulas = true; }
    }
    this.dirty.add(this.sheets[i].path);
    this.edited = true;
    // (calcChain.xml lists the formula cells in the order Excel last worked them out; one naming a cell that no longer
    // holds a formula makes Excel repair the file. Without it Excel works the order out again.)
    if (formulas) this._dropCalcChain();
    const rects = changes.map(({ r, c }) => ({ si: i, r0: r, c0: c, r1: r, c1: c }));
    // (In Manual calculation the entered formulas alone: calcmode.js.)
    return recalcAfter(this, rects, changes.filter((ch) => ch.formula != null).map(({ r, c }) => ({ si: i, r, c })));
  }

  /** Cell format `s` with a change (stylesedit.js: font, fill, border, alignment, number format, cell style): the
   *  index of the cell format that has it, added to the styles part if there's none. */
  styleWith(s, change) {
    return formatsOf(this).xf(s, change);
  }

  /** Give cells of sheet i the cell formats [{ r, c, s }]. A cell that doesn't exist is made, holding only its format
   *  (as Excel keeps a formatted empty cell). */
  setStyles(i, cells) {
    const part = this._part(i);
    const model = part.model;
    for (const { r, c, s } of cells) {
      setStyleXml(part.doc, r, c, s);
      if (!model.cells.has(r)) model.cells.set(r, new Map());
      const cell = model.cells.get(r).get(c);
      if (cell) cell.s = s;
      else model.cells.get(r).set(c, { t: 'n', v: null, s, f: null });
      model.maxRow = Math.max(model.maxRow, r);
      model.maxCol = Math.max(model.maxCol, c);
    }
    this.dirty.add(this.sheets[i].path);
    this.edited = true;
  }

  /** Sheet i's merged ranges, all of them, as [{ r0, c0, r1, c1 }]. */
  setMerges(i, merges) {
    const part = this._part(i);
    setMergesXml(part.doc, merges);
    part.model.merges = merges.map((m) => ({ ...m }));
    for (const m of merges) { part.model.maxRow = Math.max(part.model.maxRow, m.r1); part.model.maxCol = Math.max(part.model.maxCol, m.c1); }
    this.dirty.add(this.sheets[i].path);
    this.edited = true;
  }

  /** Insert `n` rows ('row') or columns ('col') of sheet i before index `at`, or delete -n from `at`, or only the cells of
   *  `band` (rowcol.js; a refusal throws a SheetError with Excel's words). */
  changeLines(i, axis, at, n, band = null, opts = {}) { changeLines(this, i, axis, at, n, band, opts); }

  /** Rows ('row') or columns ('col') a to b of sheet i resized, hidden or shown (sizes.js setRows and setColumns: { width },
   *  { height }, { hidden }, { bestFit }); the sheet read again. */
  setLines(i, axis, a, b, change) {
    const part = this._part(i);
    if (axis === 'col') setColumns(part.doc, a, b, change, part.model.format);
    else setRows(part.doc, a, b, change);
    this._reread(i, part);
    this.resetCalc();
  }

  /** Sheet i (its `part`) read again after its XML was changed, and marked changed. */
  _reread(i, part) {
    this._parsed.set(this.sheets[i], { doc: part.doc, model: readSheet(part.doc, this.sharedStrings) });
    this.dirty.add(this.sheets[i].path);
    this.edited = true;
  }

  /** Sheet i's rows' and columns' own formats: `cols` [{ c, s }], `rows` [{ r, s }] (sizes.js, as Excel writes a
   *  formatted column and row); neighbours with one format written as one element; the sheet read again once. */
  setLineStyles(i, { cols = [], rows = [] }) {
    const part = this._part(i);
    setLineStyles(part.doc, { cols, rows }, part.model.format);
    this._reread(i, part);
  }

  /**
   * Sheet i's AutoFilter set to `filter` ({ r0, c0, r1, c1, columns }, as the model reads it) or taken off (null), as
   * Excel does it (filterops.js): the autoFilter element written, the rows it doesn't let through hidden and the rest
   * of its range shown, and its range kept in the sheet's hidden _xlnm._FilterDatabase name (left when it's off, as
   * Excel leaves it).
   */
  /** Add a conditional format `rule` (cfedit.js) for cells `ranges` of sheet i, or clear the rules of those cells
   *  (rule null; ranges null: of the whole sheet). */
  setConditional(i, ranges, rule) {
    const part = this._part(i);
    if (rule) addConditional(this, i, ranges, rule); else clearConditional(this, i, ranges);
    this._reread(i, part);
  }

  /** A change to sheet i's conditional formats made by `fn()` (cfmanage.js: the Rules Manager's), the sheet read again. */
  changeConditional(i, fn) {
    const part = this._part(i);
    fn();
    this._reread(i, part);
  }

  /** Give cells `ranges` of sheet i the data validation `rule` (dvedit.js), or none (null: Clear All). */
  /** A cut's validation moved with it on sheet i (dvedit.js moveValidation). */
  moveValidation(i, rect, dr, dc) { moveValidation(this._part(i).doc, rect, dr, dc); this._reread(i, this._part(i)); }
  /** Cells added to the rule they were copied from (dvedit.js addToRule): false when it isn't there. */
  addToValidation(i, match, ranges) { return addToRule(this._part(i).doc, match, ranges) && (this._reread(i, this._part(i)), true); }
  setValidation(i, ranges, rule) {
    const part = this._part(i);
    setValidationXml(part.doc, ranges, rule, this.date1904);
    this._reread(i, part);
  }

  /** Sheet i's AutoFilter on (`filter`: its area and columns' filters) or off (null), rows hidden or shown (filterops.js). */
  setAutoFilter(i, filter) { applyAutoFilter(this, i, filter); }

  /**
   * Data > Advanced: the rows of `list` on sheet i passing `criteria` (advfilter.js), without later duplicates when
   * `unique`; in place, the others hidden; with `copyTo` ({ r, c }) the heading and the rows passing written there
   * (their values and formats). Excel's names written as it writes them: _xlnm._FilterDatabase (hidden) for the list,
   * _xlnm.Criteria for the criteria, _xlnm.Extract for the copy's heading row. Returns null, or Excel's words.
   */
  advancedFilter(i, spec) { return applyAdvancedFilter(this, i, spec); }

  /** The hidden name Excel keeps a sheet's filter range in: _xlnm._FilterDatabase, local to the sheet (filterops.js). */
  _filterName(i, area, name = '_xlnm._FilterDatabase', hidden = true) { filterName(this, i, area, name, hidden); }

  /** Sheet i's view changed (view.js): { freeze: { rows, cols }, active: { r, c }, topLeft }, { zoom } (percent) or
   *  { showFormulas }, { showGridLines }, { showHeadings }. */
  setView(i, { freeze, active, topLeft, zoom, showFormulas, showGridLines, showHeadings }) {
    const part = this._part(i);
    if (freeze) setFreeze(part.doc, freeze, active, topLeft);
    if (zoom != null) setZoom(part.doc, zoom);
    part.model.freeze = freeze ? { rows: freeze.rows, cols: freeze.cols } : part.model.freeze;
    if (zoom != null) part.model.zoom = Math.round(Math.min(400, Math.max(10, zoom)));
    if (showFormulas != null) { setShowFormulas(part.doc, showFormulas); part.model.showFormulas = !!showFormulas; }
    if (showGridLines != null) { setShown(part.doc, 'showGridLines', showGridLines); part.model.showGridLines = !!showGridLines; }
    if (showHeadings != null) { setShown(part.doc, 'showRowColHeaders', showHeadings); part.model.showHeadings = !!showHeadings; }
    this.dirty.add(this.sheets[i].path);
    this.edited = true;
  }

  /** The workbook as it is now, to go back to (restore): the parts, the sheets read so far and the workbook part. */
  snapshot() {
    return {
      files: { ...this.pkg.files }, rels: this.pkg.rels, book: serializeXml(this.doc), dirty: new Set(this.dirty), bookDirty: this.bookDirty,
      // (The sheets as they were, names, order and all, and the parsed ones by their ids: sheets added, renamed, moved
      // or deleted since are put back too.)
      sheets: this.sheets.map((info) => ({ ...info })),
      parsed: [...this._parsed].map(([info, part]) => [info.sheetId, serializeXml(part.doc)]),
    };
  }

  /** Back to a snapshot: its parts, its sheets (their list read again) and its defined names. Cell formats added
   *  since stay (no cell of the snapshot uses them). */
  restore(snap) {
    this.pkg.files = { ...snap.files };
    this.pkg.rels = snap.rels;
    this.doc = parseXml(snap.book);
    const root = this.doc.documentElement;
    const els = kids(kid(root, 'sheets', this.ns), 'sheet', this.ns);
    this.sheets = snap.sheets.map((info) => withEl({ ...info }, els.find((el) => Number(attr(el, 'sheetId')) === info.sheetId) || null));
    this._parsed = new Map(snap.parsed.map(([id, text]) => {
      const doc = parseXml(text);
      return [this.sheets.find((s) => s.sheetId === id), { doc, model: readSheet(doc, this.sharedStrings) }];
    }).filter(([info]) => info));
    this.definedNames = kids(kid(root, 'definedNames', this.ns), 'definedName', this.ns).map((el) => withEl({
      name: attr(el, 'name'),
      sheet: attr(el, 'localSheetId') == null ? null : Number(attr(el, 'localSheetId')),
      text: el.textContent,
    }, el));
    this.dirty = new Set(snap.dirty);
    this._dynamicCm = null;
    this.bookDirty = true;
    this.edited = true;
    this.resetCalc();
  }

  /** A dynamic array formula's cell in the file's form (spillxml.js), and a cell its spill no longer reaches emptied. */
  setDynamic(i, r, c, rect) { setDynamicXml(this, i, r, c, rect); }
  clearSpilled(i, r, c) { clearSpilledXml(this, i, r, c); }
  /** Whether a dynamic formula spilled into (r, c) of sheet i (the value there isn't the cell's own). */
  isSpilled(i, r, c) { return !!this._calc?.spilledBy?.get(`${i}:${r}:${c}`); }

  /** Sheets added, renamed, moved, hidden or shown, coloured and deleted as Excel does it (sheetops.js; a refusal
   *  throws a SheetError with Excel's words). `addSheet` returns the new sheet's index. */
  addSheet(at, name) { return addSheet(this, at, name); }
  renameSheet(i, name) { renameSheet(this, i, name); }
  moveSheet(from, to) { moveSheet(this, from, to); }
  setSheetHidden(i, hidden) { setSheetHidden(this, i, hidden); }
  setTabColor(i, color) { setTabColor(this, i, color); }
  deleteSheet(i) { deleteSheet(this, i); }

  /** Start the calculation state again (sheets' indexes or names changed), remembering that some formula couldn't be
   *  worked out here. */
  resetCalc() {
    if (this._calc && this._calc.incomplete) this._incomplete = true;
    this._calc = null;
  }

  /** Work out every formula of the workbook again (Excel's Ctrl+Alt+F9). */
  calculate() {
    return this.calc.calculateAll();
  }

  /** Work out the formulas the file came without a result for (as programs other than Excel may save them), and those
   *  that read them, as Excel works them out on opening such a file; every other formula keeps the result saved with
   *  it. The results are shown, and saved with a sheet once it is edited: a workbook opened and saved unchanged still
   *  comes back byte for byte. Returns how many had none. */
  calculateMissing() {
    const also = [];
    this.sheets.forEach((info, si) => {
      if (info.kind !== 'worksheet' || !info.path) return;
      for (const [r, line] of this.sheet(si).cells) for (const [c, cell] of line) if (cell.f != null && cell.v == null) also.push({ si, r, c });
    });
    if (!also.length) return 0;
    const dirty = new Set(this.dirty);
    this.calc.recalc([], also);
    this.dirty = dirty;
    return also.length;
  }

  /** Write a formula's result into a cell (from the engine); nothing changes if it is the value the cell has. */
  writeResult(i, r, c, value) {
    const part = this._part(i);
    const model = part.model;
    const cell = model.cells.get(r)?.get(c);
    const t = typeof value === 'number' ? 'n' : typeof value === 'boolean' ? 'b' : typeof value === 'string' ? 'str' : 'e';
    const v = t === 'e' ? value.code : value;
    if (cell && cell.t === t && cell.v === v) return;
    // (#SPILL! and #CALC! are rich values in Excel's file: #VALUE! there, worked out again when Excel opens it.)
    if (t === 'e' && RICH_ERRORS.has(v)) this._incomplete = true;
    setResultXml(part.doc, r, c, t === 'e' ? (RICH_ERRORS.has(v) ? { code: '#VALUE!' } : value) : v);
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
  /** The file's kind: 'xlsx', 'xltx' (a template), 'xlsm' or 'xltm' (filetype.js). */
  get fileType() { return fileType(this.pkg); }
  /** Made a file of kind `type` when saved, as Excel's Save As (filetype.js setFileType). */
  setFileType(type) { setFileType(this.pkg, type); }

  toBytes() {
    // (A formula the engine couldn't work out may read a changed cell: have Excel recalculate the workbook on opening.)
    if (this.edited && ((this._calc && this._calc.incomplete) || this._incomplete)) this._fullCalcOnLoad();
    for (const [info, part] of this._parsed) {
      if (this.dirty.has(info.path) && this.sheets.includes(info)) this.pkg.setText(info.path, partText(part.doc, this.pkg.text(info.path)));
    }
    if (this.stylesPath && this.dirty.has(this.stylesPath)) this.pkg.setText(this.stylesPath, partText(this.stylesDoc, this.pkg.text(this.stylesPath)));
    if (this.bookDirty) this.pkg.setText(this.pkg.mainPath, partText(this.doc, this.pkg.text(this.pkg.mainPath)));
    this.bookDirty = false;
    const bytes = this.pkg.toBytes();
    return this.encryption ? encryptWorkbook(bytes, this.encryption.password, this.encryption.streams) : bytes;
  }

  _fullCalcOnLoad() {
    const doc = this.doc;
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
    this.bookDirty = true;
  }
}
