// Commands for parts of a workbook beyond its cells (commands.js shares its state with them): going to a found
// cell, changes of several sheets at once, PivotTables refreshed, notes and rules changed, another editor's contents
// and the agent's changes.
import { readPivots, refreshPivot, pivotAt } from '../ooxml/pivot.js';
import { readQueries, refreshQueryTable, setQueryFormula } from '../ooxml/queries.js';
import { addHyperlink } from '../ooxml/links.js';
import { SheetError } from '../ooxml/sheetops.js';
import { setCalcMode, calculateNow, calculateSheet, calculateFull, setIteration } from '../ooxml/calcmode.js';

/** These commands: `env` is createCommands's shared state and helpers (commands.js). */
export function partCommands(env) {
  const { allowed, commands, ctx, state, idOf, runSteps, sheetOp } = env;
  /** After a calculation: the grid and the status bar drawn again, circular references said. */
  const calculated = (out) => {
    ctx.grid.refresh();
    ctx.changed?.();
    if (out?.circular?.length) ctx.announce(`Circular reference: ${out.circular.length} formulas read themselves and weren't worked out`);
    return true;
  };
  return {
    /** Go to a found cell ({ si, r, c }): its sheet shown, the cell selected. */
    goTo(hit) {
      if (hit.si !== ctx.sheetIndex() && !commands.showSheet(hit.si)) return false;
      commands.select(hit.r, hit.c);
      return true;
    },
    /** Changes of several sheets ([{ si, contents, styles, merges }], findops.js replaceChanges) as one step. */
    applyChanges(steps) { return runSteps(steps); },
    /** Refresh (Alt+F5: the PivotTable at the active cell) or Refresh All (Ctrl+Alt+F5: every one), from their
     *  sources as Excel's Refresh does (pivot.js), as one step of the history; a pivot that can't be refreshed here
     *  leaves everything as it was, and says so. */
    refreshPivots(all = false) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      // (Excel refused a refresh on a protected sheet, Use PivotTable allowed or not: its cells are locked.)
      if (!allowed('always')) return false;
      const wb = ctx.workbook();
      const here = pivotAt(wb, ctx.sheetIndex(), state.sel.ar, state.sel.ac);
      const list = all ? readPivots(wb).filter((p, k, a) => a.findIndex((q) => q.cachePath === p.cachePath) === k) : here ? [here] : [];
      if (!list.length) { ctx.announce(all ? 'There are no PivotTables to refresh' : 'The active cell is not in a PivotTable'); return false; }
      const before = wb.snapshot();
      for (const p of list) {
        const out = refreshPivot(wb, p);
        if (out.error) {
          wb.restore(before);
          ctx.reloaded?.();
          ctx.alert(out.why ? `${out.error}: ${out.why}.` : out.error);
          return false;
        }
        for (const ch of out.changes) if (ch.contents.length) wb.apply(ch.si, ch.contents);
      }
      state.undo.push({ sheet: idOf(ctx.sheetIndex()), snap: true, before, after: wb.snapshot() });
      state.redo = [];
      ctx.edited?.({ book: true });
      ctx.reloaded?.();
      ctx.changed?.();
      ctx.announce(list.length === 1 ? 'PivotTable refreshed' : `${list.length} PivotTables refreshed`);
      return true;
    },
    /** Rows a host gave for queries and connections ([{ name, kind: 'query' | 'connection', data }]: queriespane.js
     *  refreshThrough) written into their tables as Excel's Refresh writes them (queries.js refreshQueryTable), all as
     *  one step of the history; one that can't be written leaves everything as it was, and says why. */
    queryRows(results) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      // (A protected sheet: its cells are locked, as for a PivotTable's refresh.)
      if (!allowed('always')) return false;
      const wb = ctx.workbook();
      const before = wb.snapshot();
      let rows = 0;
      try {
        for (const { name, kind, data } of results) {
          const { queries, connections } = readQueries(wb);
          const q = kind === 'connection' ? connections.find((c) => c.name === name) : queries.find((x) => x.name === name);
          const out = q ? refreshQueryTable(wb, q, data) : { error: `There's no ${kind === 'connection' ? 'connection' : 'query'} named ${name}` };
          if (out.error) throw new Error(out.error);
          rows += out.rows;
        }
      } catch (e) {
        wb.restore(before);
        ctx.reloaded?.();
        ctx.alert(e.message);
        return false;
      }
      state.undo.push({ sheet: idOf(ctx.sheetIndex()), snap: true, before, after: wb.snapshot() });
      state.redo = [];
      ctx.edited?.({ book: true });
      ctx.reloaded?.();
      ctx.changed?.();
      ctx.announce(`${rows.toLocaleString('en-US')} row${rows === 1 ? '' : 's'} loaded`);
      return true;
    },
    /** The Advanced Editor's Done: query `name`'s M changed to `m` in the DataMashup (queries.js setQueryFormula), one
     *  step of the history; Excel's words when it can't be written. */
    queryFormula(name, m) {
      let refused = null;
      const done = commands.sheetChange((wb) => { const out = setQueryFormula(wb, name, m); if (out?.error) { refused = out.error; return false; } return true; });
      if (refused) ctx.alert(refused);
      return done;
    },
    /** Formulas > Calculation Options: 'auto', 'autoNoTable' or 'manual' (calcmode.js). Not a step of the history,
     *  as in Excel; to Automatic, the formulas waiting are worked out. */
    calcMode(mode) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const out = setCalcMode(ctx.workbook(), mode);
      ctx.edited?.({ book: true });
      return calculated(out);
    },
    /** Iterative calculation on or off, with its Maximum Iterations and Maximum Change ({ on, count, delta }:
     *  calcmode.js setIteration); circular formulas worked out again. Not a step of the history, as in Excel. */
    iteration(settings) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const out = setIteration(ctx.workbook(), settings);
      ctx.edited?.({ book: true });
      return calculated(out);
    },
    /** Calculate Now (F9: what changed since reaches), Calculate Sheet (Shift+F9: on the sheet shown) or every formula
     *  of the workbook (Ctrl+Alt+F9: 'all'). */
    calculate(how = 'now') {
      const wb = ctx.workbook();
      return calculated(how === 'sheet' ? calculateSheet(wb, ctx.sheetIndex()) : how === 'all' ? calculateFull(wb) : calculateNow(wb));
    },
    /** Insert > Link's answer for cell (r, c) of the sheet shown (links.js addHyperlink's { address, location, tooltip,
     *  text }), as one step of the history; `text` given replaces a text cell's value, as Excel's Text to display does.
     *  A protected sheet takes it only with Insert hyperlinks allowed and the cell unlocked (measured). */
    insertHyperlink(r, c, link) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      if (!allowed('insertHyperlinks') || !allowed('contents', [{ r, c }])) return false;
      const done = commands.sheetChange((wb, si) => {
        addHyperlink(wb, si, r, c, link);
        const v = wb.sheet(si).cells.get(r)?.get(c);
        if (link.text != null && v?.f == null && (v?.v == null || typeof v.v === 'string') && v?.v !== link.text) wb.setCell(si, { r, c }, link.text);
      });
      // (The Hyperlink style may be a new cell format: its CSS made.)
      if (done) ctx.restyled?.();
      return done;
    },
    /** A change to the sheet's notes or rules, `run(wb, si)` (noteedit.js setNote, deleteNote; Workbook
     *  setValidation), as one step of the history. `kind`, what a protected sheet must allow for it (commands.js
     *  allowed: 'objects', 'formatCells', 'always' for never; none for a change it doesn't stop). */
    sheetChange(run, kind = null) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      if (kind && !allowed(kind)) return false;
      const wb = ctx.workbook();
      const before = wb.snapshot();
      const formats = wb.styles.cellXfs.length;
      if (run(wb, ctx.sheetIndex()) === false) return false;
      state.undo.push({ sheet: idOf(ctx.sheetIndex()), snap: true, before, after: wb.snapshot() });
      state.redo = [];
      ctx.edited?.({ book: true });
      // (Cell formats added, as Convert to Range adds them: their classes made before the sheet is drawn again.)
      if (wb.styles.cellXfs.length !== formats) ctx.restyled?.();
      ctx.reloaded?.();
      ctx.changed?.();
      return true;
    },
    /** Cell contents [{ r, c, value } or { r, c, formula }] of sheet si that another editor of a shared workbook
     *  changed (collab.js): made here outside the history, and drawn. */
    remoteContents(si, contents) {
      ctx.workbook().apply(si, contents);
      if (si === ctx.sheetIndex()) ctx.grid.refresh();
      ctx.changed?.();
      return true;
    },
    /** A change another editor of a shared workbook made (collab.js), `run(wb)`: made here outside the history, and
     *  drawn, with the cell formats it added, and the tabs again when it changed the sheets (`sheets`). */
    remoteChange(run, { sheets = false } = {}) {
      const wb = ctx.workbook();
      const was = ctx.sheetIndex();
      const shownId = idOf(was);
      const formats = wb.styles.cellXfs.length;
      run(wb);
      if (sheets) env.sheetsMoved(wb, shownId, was);
      if (wb.styles.cellXfs.length !== formats) ctx.restyled?.();
      ctx.reloaded?.();
      ctx.changed?.();
      return true;
    },
    /** A change by the agent (agent.js), `run(wb)`: of rows or columns ('lines': one step of the history) or of the
     *  sheets ('sheets', as the tab menu's). A refusal throws, in Excel's words, for the agent to read. */
    agentChange(run, kind) {
      if (ctx.readOnly()) throw new SheetError('The workbook is read-only');
      if (kind === 'sheets') return sheetOp(run, null, true);
      const wb = ctx.workbook();
      const before = wb.snapshot();
      run(wb);
      state.undo.push({ sheet: idOf(ctx.sheetIndex()), snap: true, before, after: wb.snapshot() });
    ctx.edited?.({ book: true });
      state.redo = [];
      commands.dropClip();
      ctx.reloaded?.();
      ctx.changed?.();
      return true;
    },
  };
}
