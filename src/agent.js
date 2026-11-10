// A workbook for an AI agent (or any program) to read and edit without a browser: its sheets, names and tables, any
// range as values, shown text and formulas, a Markdown view, search, notes, and edits in plain terms (cells entered as
// a person types them, blocks of values, rows and columns inserted or deleted, sheets added, renamed or deleted).
// Ranges are what the Name Box takes ("B2:D9", "Data!A:A", a defined name, a table or table reference). Saving keeps
// everything else in the file as it was. In Node, set up a DOMParser/XMLSerializer first (setXmlImpl, as for the
// Workbook). Excel has no tracked changes for cells, so edits are made directly; in an editor (`api.agent()`) each is
// one step of its undo history.
import { Workbook } from './ooxml/workbook.js';
import { goToTarget, goToNames } from './ooxml/goto.js';
import { entryChanges } from './ooxml/entry.js';
import { editText } from './ooxml/input.js';
import { cellText } from './ooxml/numfmt.js';
import { formatCodeOf } from './ooxml/styles.js';
import { findAll } from './ooxml/findops.js';
import { readNotes } from './ooxml/notes.js';
import { readTables } from './ooxml/tables.js';
import { ruleAt, valid, VALIDATION_MESSAGE } from './ooxml/validation.js';
import { refName, rangeName, colIndex, MAX_ROWS, MAX_COLS } from './ooxml/refs.js';
import { sheetPrefix } from './calc/tokens.js';

/** Opens an .xlsx or .xlsm (bytes) for reading and editing. */
export function openWorkbook(bytes) {
  return new AgentWorkbook(Workbook.fromBytes(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)));
}

/** A cell's value as JavaScript has it: a number, string or boolean, an error as its text ("#DIV/0!"), or null. */
function valueOf(cell) {
  if (!cell || cell.v == null) return null;
  if (cell.t === 'n') return cell.v;
  if (cell.t === 'b') return !!cell.v;
  return String(cell.v);
}

export class AgentWorkbook {
  /** `edits`: how changes are made, { cells(si, contents, styles), structure(fn, kind: 'lines' | 'sheets') } (an
   *  editor's, for its history). */
  constructor(wb, edits = null) {
    this._wb = typeof wb === 'function' ? wb : () => wb;
    this.edits = edits || {
      cells: (si, contents, styles) => {
        const { circular } = this.wb.apply(si, contents);
        if (styles.length) this.wb.setStyles(si, styles);
        return { circular };
      },
      structure: (fn) => fn(this.wb),
    };
  }

  get wb() { return this._wb(); }

  // --- Reading --------------------------------------------------------------------------------------------------

  /** Every sheet, in tab order: [{ name, kind ('worksheet' or 'chartsheet'), hidden, used (its used range, or
   *  null), tables (names) }]. */
  sheets() {
    const tables = readTables(this.wb);
    return this.wb.sheets.map((info, si) => {
      const worksheet = info.kind === 'worksheet';
      const sh = worksheet ? this.wb.sheet(si) : null;
      return {
        name: info.name,
        kind: info.kind,
        hidden: !this.wb.canShow(si) && worksheet,
        used: sh ? usedRange(sh) : null,
        tables: tables.filter((t) => t.si === si).map((t) => t.name),
      };
    });
  }

  /** The defined names and tables there are to read or go to: [{ name, refersTo, table }]. */
  names() {
    const tables = readTables(this.wb);
    return goToNames(this.wb, this.defaultSheet()).map((name) => {
      const t = tables.find((x) => x.name === name);
      if (t) return { name, refersTo: this.qualified(t.si, t), table: true, columns: t.columns };
      return { name, refersTo: this.wb.definedName(name, this.defaultSheet())?.text ?? null, table: false };
    });
  }

  /** The sheet a range without a sheet name is on: `sheet` (a name or index), else the active sheet. */
  defaultSheet(sheet = null) {
    if (sheet == null) return this.wb.activeSheet;
    const si = typeof sheet === 'number' ? sheet : this.wb.sheetIndexOf(sheet);
    if (si < 0 || si >= this.wb.sheets.length || this.wb.sheets[si].kind !== 'worksheet') throw new Error(`no such worksheet: ${sheet}`);
    return si;
  }

  /** Where `range` is: { si, r0, c0, r1, c1 }, whole columns and rows cut to the used range. Throws on a range the
   *  Name Box wouldn't take. */
  where(range, { sheet = null } = {}) {
    const si = this.defaultSheet(sheet);
    const text = range ?? usedRange(this.wb.sheet(si)) ?? 'A1';
    const t = goToTarget(this.wb, si, text);
    if (t.error) throw new Error(`not a range, name or table: ${text}`);
    const sh = this.wb.sheet(t.si);
    const { r0, c0 } = t.sel;
    let { r1, c1 } = t.sel;
    if (r1 === MAX_ROWS - 1) r1 = Math.max(r0, sh.maxRow ?? 0);
    if (c1 === MAX_COLS - 1) c1 = Math.max(c0, sh.maxCol ?? 0);
    return { si: t.si, r0, c0, r1, c1 };
  }

  /** "Data!B2:C3", with the sheet's name quoted when it needs it. */
  qualified(si, area) { return sheetPrefix(this.wb.sheets[si].name) + rangeName(area); }

  /**
   * A range's cells, row by row: { range ("Data!A1:E9"), rows: [[{ ref, value, text, formula }]] }. `value` is the
   * value (or a formula's result) as a number, string, boolean, error text or null; `text` is what the cell shows, its
   * number format applied; `formula` starts with "=" (null for a constant).
   */
  read(range = null, options = {}) {
    const a = this.where(range, options);
    const sh = this.wb.sheet(a.si);
    const rows = [];
    for (let r = a.r0; r <= a.r1; r++) {
      const row = [];
      for (let c = a.c0; c <= a.c1; c++) row.push(this.cell(a.si, sh, r, c));
      rows.push(row);
    }
    return { range: this.qualified(a.si, a), rows };
  }

  /** One cell as read() gives it. */
  cell(si, sh, r, c) {
    const cell = sh.cells.get(r)?.get(c);
    const code = formatCodeOf(this.wb.styles, cell?.s || 0);
    return {
      ref: refName(r, c),
      value: valueOf(cell),
      text: cell ? cellText(cell, Infinity, code, this.wb.date1904) : '',
      formula: cell && cell.f != null && cell.f !== '' ? `=${cell.f}` : (cell && cell.shared != null ? editText(cell, code, this.wb.date1904) : null),
    };
  }

  /**
   * A range as a Markdown table: the column letters across the top, each row's number first, each cell as it shows
   * (`formulas`: a formula cell as its formula). Rows and columns past the used range are left out.
   */
  markdown(range = null, { formulas = false, ...options } = {}) {
    const { range: name, rows } = this.read(range, options);
    if (!rows.length) return '';
    const esc = (s) => String(s).replace(/\\/g, '\\\\').replace(/\|/g, '\\|').replace(/\r?\n/g, '<br>');
    const a = this.where(range, options);
    const head = ['', ...rows[0].map((x) => x.ref.replace(/\d+$/, ''))];
    const lines = [`${name}`, '', `| ${head.join(' | ')} |`, `|${head.map(() => '---').join('|')}|`];
    rows.forEach((row, k) => lines.push(`| ${a.r0 + k + 1} | ${row.map((x) => esc(formulas && x.formula ? x.formula : x.text)).join(' | ')} |`));
    return lines.join('\n');
  }

  /**
   * Cells whose formula or text (`lookIn` 'formulas', the default, or 'values': as shown) matches `query`, with
   * Excel's wildcards (* and ?, ~ to escape): [{ ref ("Data!B4"), text }]. `within` 'workbook' (the default) or
   * 'sheet' (of `sheet`).
   */
  find(query, { matchCase = false, whole = false, lookIn = 'formulas', within = 'workbook', sheet = null } = {}) {
    const si = this.defaultSheet(sheet);
    return findAll(this.wb, si, query, { matchCase, whole, lookIn, within }).map((h) => {
      const x = this.cell(h.si, this.wb.sheet(h.si), h.r, h.c);
      return { ref: this.qualified(h.si, { r0: h.r, c0: h.c, r1: h.r, c1: h.c }), text: x.formula ?? x.text };
    });
  }

  /** A sheet's notes and threaded comments: [{ ref, kind ('note' or 'thread'), author, text, replies }]. */
  notes({ sheet = null } = {}) {
    const si = this.defaultSheet(sheet);
    return [...readNotes(this.wb, si)].map(([key, n]) => {
      const [r, c] = key.split(',').map(Number);
      const first = n.kind === 'thread' ? n.items[0] : n;
      return {
        ref: refName(r, c), kind: n.kind, author: first?.author ?? null,
        text: n.kind === 'thread' ? first.text : n.text,
        replies: n.kind === 'thread' ? n.items.slice(1).map((x) => ({ author: x.author, text: x.text })) : [],
      };
    }).sort((x, y) => (x.ref.length - y.ref.length) || x.ref.localeCompare(y.ref));
  }

  // --- Editing --------------------------------------------------------------------------------------------------

  /**
   * Enter text into cells as a person types it ({ "B2": "12", "C2": "=B2*2", "D2": "50%", "Data!A9": "Kiwi" }):
   * numbers, dates, percentages and TRUE read as Excel reads them, "=" a formula, "'" text. A cell's data validation
   * refuses a value it wouldn't take unless `force`. Returns { written: [refs], refused: [{ ref, error }] }; the
   * written cells' formulas, and every formula reading them, are worked out again.
   */
  write(entries, { sheet = null, force = false } = {}) {
    const bySheet = new Map();
    const refused = [];
    const written = [];
    for (const [ref, text] of Object.entries(entries)) {
      let a;
      try { a = this.where(ref, { sheet }); } catch (e) { refused.push({ ref, error: e.message }); continue; }
      if (a.r0 !== a.r1 || a.c0 !== a.c1) { refused.push({ ref, error: 'one cell at a time: use setValues for a block' }); continue; }
      const sh = this.wb.sheet(a.si);
      if (sh.arrays.some((x) => a.r0 >= x.r0 && a.r0 <= x.r1 && a.c0 >= x.c0 && a.c0 <= x.c1)) { refused.push({ ref, error: "You can't change part of an array" }); continue; }
      const ch = entryChanges(this.wb, a.si, a.r0, a.c0, String(text ?? ''));
      if (ch.error) { refused.push({ ref, error: ch.error }); continue; }
      const rule = !force && ruleAt(sh, a.r0, a.c0);
      if (rule && rule.showError) {
        const v = ch.content.formula != null ? this.wb.calc.formulaAt(a.si, a.r0, a.c0, ch.content.formula) : ch.content.value;
        if (!valid(this.wb, a.si, rule, a.r0, a.c0, v)) { refused.push({ ref, error: rule.error || VALIDATION_MESSAGE }); continue; }
      }
      if (!bySheet.has(a.si)) bySheet.set(a.si, { contents: [], styles: [] });
      bySheet.get(a.si).contents.push(ch.content);
      if (ch.style) bySheet.get(a.si).styles.push(ch.style);
      written.push(this.qualified(a.si, a));
    }
    for (const [si, { contents, styles }] of bySheet) this.edits.cells(si, contents, styles);
    return { written, refused };
  }

  /**
   * A block of values from `topLeft` ("A1" or "Data!B2"), row by row: numbers, strings (as they are, never read as a
   * formula or a date), booleans and null (an empty cell). Returns the range written.
   */
  setValues(topLeft, rows, { sheet = null } = {}) {
    const a = this.where(topLeft, { sheet });
    const contents = [];
    rows.forEach((row, dr) => row.forEach((value, dc) => {
      if (value !== null && !['number', 'string', 'boolean'].includes(typeof value)) throw new Error(`not a cell value: ${value}`);
      contents.push({ r: a.r0 + dr, c: a.c0 + dc, value });
    }));
    if (!contents.length) return null;
    const width = Math.max(...rows.map((row) => row.length));
    this.edits.cells(a.si, contents, []);
    return this.qualified(a.si, { r0: a.r0, c0: a.c0, r1: a.r0 + rows.length - 1, c1: a.c0 + width - 1 });
  }

  /** Insert `count` rows before row `at` (its number, 1-based), as Excel's Insert does: formulas, names, tables, merges
   *  and the rest follow. Excel's refusals throw with its words. */
  insertRows(at, count = 1, { sheet = null } = {}) { this.lines(sheet, 'row', lineIndex('row', at), count); }
  deleteRows(at, count = 1, { sheet = null } = {}) { this.lines(sheet, 'row', lineIndex('row', at), -count); }
  /** The same for columns, `at` a letter ("C") or 1-based number. */
  insertColumns(at, count = 1, { sheet = null } = {}) { this.lines(sheet, 'col', lineIndex('col', at), count); }
  deleteColumns(at, count = 1, { sheet = null } = {}) { this.lines(sheet, 'col', lineIndex('col', at), -count); }

  lines(sheet, axis, at, n) {
    const si = this.defaultSheet(sheet);
    if (!Number.isInteger(n) || n === 0) throw new Error('count must be a whole number other than 0');
    this.edits.structure((wb) => wb.changeLines(si, axis, at, n), 'lines');
  }

  /** A new worksheet (at the end, or before sheet `before`); returns its name. Excel's refusals throw. */
  addSheet(name = null, { before = null } = {}) {
    const at = before == null ? this.wb.sheets.length : this.defaultSheet(before);
    let k = -1;
    this.edits.structure((wb) => { k = wb.addSheet(at, name); }, 'sheets');
    return this.wb.sheets[k].name;
  }

  renameSheet(sheet, name) { const si = this.defaultSheet(sheet); this.edits.structure((wb) => wb.renameSheet(si, name), 'sheets'); }
  deleteSheet(sheet) { const si = this.defaultSheet(sheet); this.edits.structure((wb) => wb.deleteSheet(si), 'sheets'); }

  /** Work out every formula again (Excel's Ctrl+Alt+F9). */
  calculate() { this.wb.calculate(); }

  /** The workbook's bytes, every part not changed as it was. */
  save() { return this.wb.toBytes(); }
}

/** Row "5" / 5, or column "C" / 3, as a 0-based index. */
function lineIndex(axis, at) {
  if (axis === 'col' && typeof at === 'string' && /^[A-Za-z]{1,3}$/.test(at)) return colIndex(at);
  const n = Number(at);
  if (!Number.isInteger(n) || n < 1) throw new Error(`not a ${axis === 'row' ? 'row' : 'column'}: ${at}`);
  return n - 1;
}

/** The range from the first to the last cell with something in it ("A1:E9"), or null for an empty sheet. */
function usedRange(sh) {
  let r0 = Infinity, c0 = Infinity, r1 = -1, c1 = -1;
  for (const [r, line] of sh.cells) {
    for (const [c, cell] of line) {
      if (cell.v == null && cell.f == null && cell.shared == null) continue;
      if (r < r0) r0 = r;
      if (r > r1) r1 = r;
      if (c < c0) c0 = c;
      if (c > c1) c1 = c;
    }
  }
  return r1 < 0 ? null : rangeName({ r0, c0, r1, c1 });
}
