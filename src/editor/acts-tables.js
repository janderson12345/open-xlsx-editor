// The ribbon's table actions (acts.js): Insert > Table, Ctrl+T and Ctrl+L, Home > Format as Table and the Table
// Design tab, each through the editor's sheetChange command as one step of the history, the table changed by
// ooxml/tableedit.js as Excel changes it (docs/TABLES.md).
import { createTableBox, resizeTableBox } from './tableboxes.js';
import { messageBox } from './dialogs.js';
import { pickBox } from './datadialogs.js';
import { parseRange, rangeName } from '../ooxml/refs.js';
import { currentRegion } from '../ooxml/special.js';
import { readTables } from '../ooxml/tables.js';
import { createTable, tableAt, setTableStyle, setTableOption, setTotalsRow, setTotalFunction, setHeaderRow, renameTable, resizeTable, TABLE_EDIT_MESSAGES } from '../ooxml/tableedit.js';
import { convertToRange } from '../ooxml/tableconvert.js';
import { sheetPrefix } from '../calc/tokens.js';

/** The total row's drop-down, as Excel lists it: [function, label] ('none' empties the cell). */
export const TOTAL_MENU = [['none', 'None'], ['average', 'Average'], ['count', 'Count'], ['countNums', 'Count Numbers'], ['max', 'Max'], ['min', 'Min'], ['sum', 'Sum'], ['stdDev', 'StdDev'], ['var', 'Var']];
/** A range as Excel's boxes show it: =$B$2:$E$7. */
const absolute = (a) => `=${rangeName(a).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}`;

export function createTableActs(env) {
  const { root, state, editing } = env;
  const wb = () => state.workbook;
  const si = () => state.sheetIndex;
  const active = () => { const s = editing().selection; return [s.ar, s.ac]; };
  const here = () => { const [r, c] = active(); return tableAt(wb(), si(), r, c); };
  const home = (x) => { state.grid.el.focus({ preventScroll: true }); return x; };
  /** A change to the active cell's table, one step of the history (refused on a protected sheet, as Excel refuses
   *  table changes there). */
  const change = (fn) => {
    let out = null;
    const done = editing().run('sheetChange', (book, k) => { out = fn(book, k); return out?.error ? false : out !== false; }, 'always');
    if (out?.error) messageBox(root, out.error);
    return home(done && !out?.error);
  };
  /** A typed range of the sheet shown ($B$2:$E$7, B2:E7, =Sheet1!B2:E7), or null. */
  const rangeOf = (text) => {
    let t = String(text || '').trim().replace(/^=/, '');
    const bang = t.lastIndexOf('!');
    if (bang >= 0) {
      const sheet = t.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'");
      if (sheet.toLowerCase() !== wb().sheets[si()].name.toLowerCase()) return null;
      t = t.slice(bang + 1);
    }
    return parseRange(t.replace(/\$/g, ''));
  };
  /** Excel's guess at headers: a first row of text over a row that isn't all text, or a single row. */
  const guessHeaders = (a) => {
    const line = wb().sheet(si()).cells.get(a.r0);
    const next = wb().sheet(si()).cells.get(a.r0 + 1);
    let text = true;
    let below = false;
    for (let c = a.c0; c <= a.c1; c++) {
      const v = line?.get(c)?.v;
      if (v != null && typeof v !== 'string') text = false;
      const w = next?.get(c)?.v;
      if (w != null && typeof w !== 'string') below = true;
    }
    return text && (below || a.r0 === a.r1);
  };
  /** Create Table for the selection (one cell: the data round it), then `make(area, headers)`. */
  const askRange = (make, title) => {
    const s = editing().selection;
    const one = s.r0 === s.r1 && s.c0 === s.c1;
    const area = one ? currentRegion(wb().sheet(si()), s.ar, s.ac) : { r0: s.r0, c0: s.c0, r1: s.r1, c1: s.c1 };
    return createTableBox(root, {
      title, range: absolute(area), headers: guessHeaders(area),
      check: (a) => (rangeOf(a.range) ? null : 'The reference isn\'t valid.'),
    }).then((a) => (a ? change((book, k) => make(book, k, rangeOf(a.range), a.headers)) : home(false)));
  };

  const acts = {
    /** The active cell's table, for the Table Design tab. */
    tableAt: () => here(),
    /** Insert > Table, Ctrl+T, Ctrl+L: Create Table, the table in the workbook's default style. */
    tableInsert() {
      if (here()) return false;
      return askRange((book, k, area, headers) => createTable(book, k, area, { headers }), 'Create Table');
    },
    /** Home > Format as Table's style `name`: the active cell's table given it, or (outside one) a table made in it. */
    formatAsTable(name) {
      const t = here();
      if (t) return change((book) => setTableStyle(book, t, name));
      return askRange((book, k, area, headers) => createTable(book, k, area, { headers, style: name }), 'Create Table');
    },
    /** The style gallery's style for the active cell's table (null: None). */
    tableStyle(name) { const t = here(); return !!t && change((book) => setTableStyle(book, t, name)); },
    /** Table Style Options: `key` first, last, rows, cols, filterButton; headers and totals (Header Row, Total Row). */
    tableOption(key, on) {
      const t = here();
      if (!t) return false;
      if (key === 'headers') return change((book) => setHeaderRow(book, t, on));
      if (key === 'totals') return change((book) => setTotalsRow(book, t, on));
      return change((book) => setTableOption(book, t, key, on));
    },
    /** Table Name: the name typed (Enter), refused in Excel's words. */
    tableRename(name) {
      const t = here();
      if (!t || name === t.name) return false;
      return change((book) => renameTable(book, t, name));
    },
    /** Resize Table: the box, the table's range to start. */
    tableResize() {
      const t = here();
      if (!t) return false;
      return resizeTableBox(root, { range: absolute(t), check: (text) => {
        const a = rangeOf(text);
        if (!a) return 'The reference isn\'t valid.';
        return a.r0 !== t.r0 || a.r1 <= a.r0 || a.r1 < t.r0 || a.r0 > t.r1 || a.c1 < t.c0 || a.c0 > t.c1 ? TABLE_EDIT_MESSAGES.resize : null;
      } }).then((text) => (text ? change((book) => resizeTable(book, readTables(book).find((x) => x.name === t.name), rangeOf(text))) : home(false)));
    },
    /** Convert to Range, after Excel's question. */
    async tableConvert() {
      const t = here();
      if (!t) return false;
      const b = await messageBox(root, 'Do you want to convert the table to a normal range?', ['Yes', 'No']);
      if (b !== 'Yes') return home(false);
      return change((book) => convertToRange(book, t));
    },
    /** The total row's function for column c of the active cell's table (TOTAL_MENU's). */
    tableTotal(c, fn) { const t = here(); return !!t && t.totals && change((book) => setTotalFunction(book, t, c - t.c0, fn)); },
    /** The total row's drop-down at the active cell (its button, or Alt+Down), under `rect`. */
    tableTotalMenu(rect) {
      const t = here();
      const [r, c] = active();
      if (!t || !t.totals || r !== t.r1 || state.readOnly) return false;
      const at = rect || state.grid.el.getBoundingClientRect();
      pickBox(root, { title: 'Total', at: { x: at.left, y: at.bottom + 2 }, items: TOTAL_MENU.map((x) => x[1]) })
        .then((i) => (i != null ? acts.tableTotal(c, TOTAL_MENU[i][0]) : home(false)));
      return true;
    },
    /** The table's name as Excel's Name Box would give a reference to it, for the tab. */
    tableRef: () => { const t = here(); return t ? `${sheetPrefix(wb().sheets[t.si].name)}${rangeName(t)}` : ''; },
  };
  return (name, args) => (acts[name] ? acts[name](...args) : undefined);
}
