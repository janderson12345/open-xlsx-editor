// Commands that move data (commands.js shares its state with them): copy, cut and paste (Paste Special too), fill
// and AutoFill, sort, AutoSum, the Name Box and the filter.
import { readTables } from '../ooxml/tables.js';
import { applyTableFilter } from '../ooxml/tablefilter.js';
import { refName, colName, parseRange } from '../ooxml/refs.js';
import * as clip from '../ooxml/clipops.js';
import { pasteHtml } from '../ooxml/cliphtml.js';
import { linkCopy } from '../ooxml/extlinks.js';
import { fillCopy, autoFill, doubleClickFill } from '../ooxml/fillops.js';
import { fillAcrossSteps, fillAcrossParts } from '../ooxml/fillacross.js';
import { quickSort, sortChanges, sortedNotes } from '../ooxml/sortops.js';
import { setSortState } from '../ooxml/sortstate.js';
import { autoSumAreas } from '../ooxml/autosumops.js';
import { goToTarget } from '../ooxml/goto.js';
import { validName, defineName, namesOnSheet } from '../ooxml/names.js';
import { sheetPrefix } from '../calc/tokens.js';
import { specialCells, firstShown } from '../ooxml/special.js';
import { areasOf, areasSelection } from './selection.js';
import { filterArea } from '../ooxml/filterops.js';

/** These commands: `env` is createCommands's shared state and helpers (commands.js). */
/** Excel's refusal of Copy and Cut over more than one range. (Excel copies ranges that line up, the same rows or the
 *  same columns; here, only one.) */
export const MULTIPLE_MESSAGE = "This action won't work on multiple selections.";

export function dataCommands(env) {
  const { allowed, commands, ctx, state, sheet, show, runSteps, sizeStep, moved, idOf } = env;
  /** Paste's steps run as one step of the history, the copy's notes, validation and conditional formats
   *  (`parts.notes`, `parts.validation`, `parts.cf`: clipops.js pasteNotes, pasteValidation, pasteConditionals; `opts.merge`
   *  for All merging conditional formats) put over the pasted cells too when `copy` is given and has or meets any: then
   *  the step is the workbook before and after. */
  function withNotes(steps, copy, opts = {}, parts = { notes: true, validation: true, cf: true }) {
    const wb = ctx.workbook();
    const si = ctx.sheetIndex();
    const area = steps.find((x) => x.si === si)?.area;
    if (!copy || !area) return runSteps(steps);
    const before = wb.snapshot();
    if (!runSteps(steps)) return false;
    // (A cut: its notes and validation move, clipops.js moveCutParts.)
    const notes = copy.cut ? clip.moveCutParts(wb, si, area, copy) : parts.notes && clip.pasteNotes(wb, si, area, copy, opts);
    const rules = !copy.cut && parts.validation && clip.pasteValidation(wb, si, area, copy, { transpose: !!opts.transpose });
    const cf = !copy.cut && parts.cf && clip.pasteConditionals(wb, si, area, copy, { merge: !!opts.merge, transpose: !!opts.transpose });
    if (!notes && !rules && !cf) return true;
    // (The steps' own entry gives way to one for the whole paste, notes and all.)
    state.undo.pop();
    state.undo.push({ sheet: idOf(si), snap: true, before, after: wb.snapshot() });
    ctx.reloaded?.();
    ctx.changed?.();
    return true;
  }
  /** The cells copied in an editor on this page (this one, or another: pagebooks.js keeps the last copy) that `text`,
   *  the system clipboard's, is the text of; with no text to go by, the last copy made. */
  function copyFor(text) {
    const last = ctx.sharedClip?.() || null;
    if (text == null || !text) return last || state.clip;
    return [last, state.clip].find((c) => c && c.text === text) || null;
  }
  return {
    /** Copy (Ctrl+C) or Cut (Ctrl+X) the selection: kept to be pasted, outlined until then (Esc lets it go). Returns
     *  the snapshot (clipops.js), whose `text` the caller sets to what went on the system clipboard. */
    copy(cut = false) {
      if (cut && ctx.readOnly()) { ctx.announce('The workbook is read-only'); return null; }
      if (state.sel.more?.length) { ctx.alert(MULTIPLE_MESSAGE); return null; }
      state.clip = clip.snapshot(ctx.workbook(), ctx.sheetIndex(), state.sel, cut);
      ctx.clipShown?.(state.clip);
      return state.clip;
    },
    cut() { return commands.copy(true); },
    /** Let the copy or cut go (Esc): its outline goes. */
    dropClip() {
      if (!state.clip) return false;
      ctx.dropShared?.(state.clip);
      state.clip = null;
      ctx.clipShown?.(null);
      return true;
    },
    /** Paste (Ctrl+V) at the selection: the cells copied or cut here, or, when `text` from the system clipboard isn't
     *  what was copied here, that text read cell by cell as if typed. A paste of a cut moves the cells and lets the cut
     *  go. The pasted cells are selected. */
    paste(text = null, html = null, { widths = false } = {}) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const wb = ctx.workbook();
      const si = ctx.sheetIndex();
      const copy = copyFor(text);
      const mine = copy && (!copy.cut || copy.wb === wb);
      // (From another program: its HTML with the formats, as Excel pastes it (cliphtml.js), else its text as typed.)
      const steps = mine ? (copy.cut ? clip.pasteCut : clip.pasteCopy)(wb, si, state.sel, copy)
        : html ? pasteHtml(wb, si, state.sel, html, { widths }) || (text ? clip.pasteText(wb, si, state.sel, text) : null)
          : text ? clip.pasteText(wb, si, state.sel, text) : null;
      if (!steps) return false;
      const sizes = widths && (mine ? clip.pasteWidths(state.sel, copy) : steps[0].widths || []);
      if (sizes?.length) {
        // (Keep Source Column Widths: the cells and the widths as one step of the history.)
        const before = wb.snapshot();
        if (!withNotes(steps, mine ? copy : null)) return false;
        for (const { c, width } of sizes) wb.setLines(si, 'col', c, c, { width });
        state.undo.pop();
        state.undo.push({ sheet: idOf(si), snap: true, before, after: wb.snapshot() });
        ctx.reloaded?.();
        ctx.changed?.();
      } else if (!withNotes(steps, mine ? copy : null)) return false;
      if (mine && copy.cut) commands.dropClip();
      const area = steps.find((x) => x.si === si)?.area;
      if (area) show({ ar: area.r0, ac: area.c0, ...area }, null, 'Pasted, ');
      return true;
    },
    /** Paste Special (Ctrl+Alt+V) of the cells copied here at the selection (clipops.js pasteSpecial: `paste`, `op`,
     *  `skip`, `transpose`; Column widths, `paste` 'widths'), or Paste Link (`link`), as one step of the history. The
     *  copy stays to be pasted again, as in Excel. Not for cells cut, nor for text from another program. */
    pasteSpecial({ paste = 'all', op = null, skip = false, transpose = false, link = false, text = undefined } = {}) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      // (The system clipboard's text, when it could be read: another program's cells pasted as values are that text
      // typed in, as Excel's Paste Values pastes them (measured: no formats, the HTML left aside).)
      const copy = text === undefined ? (ctx.sharedClip?.() || state.clip) : copyFor(text);
      if (!copy && text && paste === 'values' && !link && !op) {
        const steps = clip.pasteText(ctx.workbook(), ctx.sheetIndex(), state.sel, text);
        if (!runSteps(steps)) return false;
        const area = steps[0].area;
        show({ ar: area.r0, ac: area.c0, ...area }, null, 'Pasted values, ');
        return true;
      }
      if (!copy) { ctx.announce('Copy cells first: Paste Special pastes cells copied in this workbook or another one open here'); return false; }
      if (copy.cut) { ctx.announce("Paste Special isn't available for cells that were cut: paste them, or copy them instead"); return false; }
      const wb = ctx.workbook();
      const si = ctx.sheetIndex();
      if (paste === 'widths' && !link) {
        const widths = clip.pasteWidths(state.sel, copy);
        return sizeStep((book, k) => { for (const { c, width } of widths) book.setLines(k, 'col', c, c, { width }); }, 'Pasted column widths');
      }
      if (link && copy.wb !== wb) {
        // (Another workbook's cells: a link to it, its copied values kept (extlinks.js), and formulas naming it, as
        // one step of the history with the link's part.)
        const before = wb.snapshot();
        const book = linkCopy(wb, copy, copy.wb.fileName || 'Book1.xlsx');
        const steps = clip.pasteLink(wb, si, state.sel, copy, { book });
        if (!runSteps(steps)) { wb.restore(before); return false; }
        state.undo.pop();
        state.undo.push({ sheet: idOf(si), snap: true, before, after: wb.snapshot() });
        ctx.reloaded?.();
        ctx.changed?.();
        const area = steps[0].area;
        show({ ar: area.r0, ac: area.c0, ...area }, null, 'Pasted links, ');
        return true;
      }
      // (All using Source theme and All merging conditional formats paste the cells as All does: within one workbook
      // the theme is the same; the second leaves the rules already over the pasted cells.)
      const cells = paste === 'theme' || paste === 'merging' ? 'all' : paste;
      const steps = link ? clip.pasteLink(wb, si, state.sel, copy) : clip.pasteSpecial(wb, si, state.sel, copy, { paste: cells, op, skip, transpose });
      // (All and All except borders bring the copy's notes and validation, Comments and Notes and Validation each its
      // own, as Excel's do; the conditional formats come with All, Formats and All except borders, with Skip blanks
      // or an operation too, as Excel's did: clipcf.js.)
      const plain = !op && !skip;
      const parts = { notes: plain && ['all', 'noBorders', 'comments'].includes(cells), validation: plain && ['all', 'noBorders', 'validation'].includes(cells), cf: ['all', 'noBorders', 'formats'].includes(cells) };
      if (!withNotes(steps, !link && (parts.notes || parts.validation || parts.cf) ? copy : null, { transpose, merge: paste === 'merging' }, parts)) return false;
      const area = steps[0].area;
      show({ ar: area.r0, ac: area.c0, ...area }, null, link ? 'Pasted links, ' : 'Pasted, ');
      return true;
    },
    /** A data tool's steps (datatools.js, flashfill.js) as one step of the history, the cells it changed selected,
     *  `words` said. */
    runTool(steps, words = '') {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      if (!runSteps(steps)) return false;
      const area = steps[0].area;
      if (area) show({ ar: area.r0, ac: area.c0, ...area }, null, words);
      else if (words) ctx.announce(words);
      return true;
    },
    /** Fill Down ('down', Ctrl+D), Right ('right', Ctrl+R), Up or Left over the selection (fillops.js fillCopy), as one
     *  step of the history. */
    fill(dir) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const steps = fillCopy(ctx.workbook(), ctx.sheetIndex(), state.sel, dir);
      if (!runSteps(steps)) return false;
      const a = steps[0].area;
      show({ ...state.sel, r0: a.r0, c0: a.c0, r1: a.r1, c1: a.c1 }, state.focus, `Filled ${dir}, `);
      return true;
    },
    /** Home > Fill > Across Worksheets: the selection copied to the same cells of each other grouped sheet
     *  (fillacross.js), 'all', 'contents' or 'formats' as Excel's box offers, as one step of the history. */
    fillAcross(type = 'all') {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const wb = ctx.workbook();
      const si = ctx.sheetIndex();
      const targets = commands.groupedSheets();
      if (targets.length < 2) { ctx.announce('Group the sheets to fill across first'); return false; }
      const before = wb.snapshot();
      if (!runSteps(fillAcrossSteps(wb, si, state.sel, targets, type))) return false;
      if (fillAcrossParts(wb, si, state.sel, targets, type)) {
        // (Hyperlinks, validation and conditional formats are parts of the sheets: the whole workbook goes back.)
        state.undo.pop();
        state.undo.push({ sheet: idOf(si), snap: true, before, after: wb.snapshot() });
        ctx.edited?.({ book: true });
        ctx.reloaded?.();
        ctx.changed?.();
      }
      ctx.announce(`Filled across ${targets.length - 1} ${targets.length === 2 ? 'sheet' : 'sheets'}`);
      return true;
    },
    /** The fill handle dragged from the selection out to `to` (fillops.js autoFill): Excel's guess, or with Ctrl held
     *  (`ctrl`) the other way, as Excel's Ctrl-drag: a single cell counts on, more cells are copied. The filled range
     *  is selected. */
    autoFill(to, { ctrl = false, type = null } = {}) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const s = state.sel;
      const one = s.r0 === s.r1 && s.c0 === s.c1;
      const how = type || (ctrl ? (one ? 'series' : 'copy') : 'default');
      const steps = autoFill(ctx.workbook(), ctx.sheetIndex(), s, to, how);
      if (!steps || !runSteps(steps)) return false;
      show({ ...s, r0: to.r0, c0: to.c0, r1: to.r1, c1: to.c1 }, null, `Filled ${refName(to.r0, to.c0)} to ${refName(to.r1, to.c1)}, `);
      // (Auto Fill Options: the fill to be done again another way, index.js.)
      ctx.filled?.({ from: { r0: s.r0, c0: s.c0, r1: s.r1, c1: s.c1 }, to: { r0: to.r0, c0: to.c0, r1: to.r1, c1: to.c1 } });
      return true;
    },
    /** A double click on the fill handle: the selection filled down beside the next column's data (doubleClickFill). */
    fillDoubleClick() {
      const to = doubleClickFill(ctx.workbook().sheet(ctx.sheetIndex()), state.sel);
      return to ? this.autoFill(to) : false;
    },
    /** Sort (sortops.js), as one step of the history; the selection stays. With `how` ({ area, keys, header,
     *  matchCase, leftToRight }, as the Sort box gives it) that sort; else A to Z ('asc') or Z to A ('desc') as
     *  Excel's buttons sort, the data beside a selection taken in too when `expand`. Excel's refusals in its words. */
    sort(how = 'asc', { expand = false } = {}) {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      if (!allowed('sort')) return false;
      const wb = ctx.workbook();
      const si = ctx.sheetIndex();
      const plan = typeof how === 'string' ? quickSort(wb, si, state.sel, how === 'desc', { expand }) : how;
      const step = plan.error ? plan : sortChanges(wb, si, plan.area, plan);
      if (step.error) { ctx.alert(step.error); return false; }
      const before = wb.snapshot();
      const ran = step.contents.length > 0;
      if (ran && !runSteps([step])) return false;
      // (The sort saved as the sheet's sortState, as Excel saves it after every sort, sortstate.js, and the notes moved
      // with their cells, sortops.js sortedNotes: then the step is the workbook before and after, the cells' own entry
      // giving way to it.)
      const noted = ran && sortedNotes(wb, si, step);
      if (setSortState(wb, si, plan) || noted) {
        if (ran) state.undo.pop();
        state.undo.push({ sheet: idOf(si), snap: true, before, after: wb.snapshot() });
        state.redo = [];
        ctx.reloaded?.();
        ctx.changed?.();
      }
      const order = typeof how === 'string' ? (how === 'desc' ? ', Z to A' : ', A to Z') : '';
      ctx.announce(`Sorted ${refName(plan.area.r0, plan.area.c0)} to ${refName(plan.area.r1, plan.area.c1)}${order}`);
      return true;
    },
    /** AutoSum (Alt+=) and its Average, Count Numbers, Max and Min (`fn`), as one step of the history, the totals
     *  selected with what they sum (autosumops.js). Where Excel leaves a formula being edited (one cell, or nothing
     *  found to sum) that cell is selected and { edit: text } returned, for the cell editor to open with it. */
    autoSum(fn = 'sum') {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      // (Several areas: each summed as alone, all in one step, each area grown by its totals.)
      const areas = areasOf(state.sel);
      const step = autoSumAreas(ctx.workbook(), ctx.sheetIndex(), areas.length > 1 ? areas : [state.sel], fn);
      if (step.edit) { commands.select(step.edit.r, step.edit.c); return { edit: step.edit.text }; }
      if (!runSteps([step])) return false;
      if (step.areas) show(areasSelection(step.areas, { r: state.sel.ar, c: state.sel.ac }), state.focus, 'AutoSum, ');
      else show({ ...state.sel, ...step.area }, state.focus, 'AutoSum, ');
      return true;
    },
    /** The Name Box: go to the reference or name typed (goto.js), on its sheet; Excel's words when it can't. */
    nameBox(text) {
      const to = goToTarget(ctx.workbook(), ctx.sheetIndex(), text);
      // (A name that isn't there yet names the selection, workbook-wide, as Excel's Name Box does.)
      if (to.error && to.name && validName(text.trim())) {
        const s = state.sel;
        const abs = (r, c) => `$${colName(c)}$${r + 1}`;
        const refersTo = `=${sheetPrefix(ctx.workbook().sheets[ctx.sheetIndex()].name)}${abs(s.r0, s.c0)}${s.r0 === s.r1 && s.c0 === s.c1 ? '' : `:${abs(s.r1, s.c1)}`}`;
        const done = commands.sheetChange((wb) => { defineName(wb, { name: text.trim(), refersTo }); });
        if (done) ctx.announce(`${text.trim()} names the selection`);
        return done;
      }
      if (to.error) { if (to.name) ctx.announce(to.error); else ctx.alert(to.error); return false; }
      if (to.si !== ctx.sheetIndex() && !commands.showSheet(to.si)) return false;
      show(to.sel, null, 'Went to ');
      return true;
    },
    /** Paste Names' Paste List, as Excel's Range.ListNames writes it (measured over COM): from the active cell down,
     *  each name the sheet sees (names.js namesOnSheet) and beside it, as text, what it refers to ("=Sheet1!$A$1"). */
    pasteNameList() {
      const si = ctx.sheetIndex();
      const { ar, ac } = state.sel;
      const list = namesOnSheet(ctx.workbook(), si);
      if (!list.length) return false;
      const contents = list.flatMap((n, k) => [{ r: ar + k, c: ac, value: n.name }, { r: ar + k, c: ac + 1, value: `=${n.text}` }]);
      return runSteps([{ si, contents, styles: [], area: { r0: ar, c0: ac, r1: ar + list.length - 1, c1: ac + 1 } }]);
    },
    /** Go To Special: the cells of a kind (special.js; `choice` as it takes it) selected, the first of them in sight
     *  active (special.js firstShown); Excel's "No cells were found." when there are none. */
    goToSpecial(choice) {
      const s = state.sel;
      const got = specialCells(ctx.workbook(), ctx.sheetIndex(), areasOf(s), { r: s.ar, c: s.ac }, choice);
      if (got.error) { ctx.alert(got.error); return false; }
      const n = got.areas.length;
      show(areasSelection(got.areas, firstShown(sheet(), got.areas)), null, n > 1 ? `${n} ranges selected, ` : 'Selected ');
      return true;
    },
    /** Filter (Ctrl+Shift+L): the sheet's AutoFilter turned on over what Sort would take, or off (filterops.js). */
    filterToggle() {
      if (ctx.readOnly()) { ctx.announce('The workbook is read-only'); return false; }
      const sh = sheet();
      if (sh.autoFilter) return moved((wb, si) => wb.setAutoFilter(si, null)) === true && (ctx.announce('Filter off'), true);
      const { area, error } = filterArea(sh, state.sel);
      if (error) { ctx.alert(error); return false; }
      return moved((wb, si) => wb.setAutoFilter(si, { ...area, columns: new Map() })) === true && (ctx.announce(`Filter on ${refName(area.r0, area.c0)} to ${refName(area.r1, area.c1)}`), true);
    },
    /** Column c of the filter shows rows with `show` ({ values: shown texts, blank }), or all of them (null); with
     *  `table`, the filter of that table's buttons (tablefilter.js). */
    filterBy(c, show, table = null) {
      const t = table && readTables(ctx.workbook()).find((x) => x.name === table);
      const f = t ? t.filter : sheet().autoFilter;
      if (!f || ctx.readOnly()) return false;
      const columns = new Map([...f.columns].filter(([, col]) => !col.other || !t));
      // (A list of values, or filtermore.js's other filters as they come.)
      if (show) columns.set(c, show.values ? { values: show.values, blank: !!show.blank, ...(show.dates?.length ? { dates: show.dates } : {}), other: false } : { ...show, other: false });
      else columns.delete(c);
      const area = { r0: f.r0, c0: f.c0, r1: f.r1, c1: f.c1, columns };
      return moved((wb, si) => (t ? applyTableFilter(wb, t, area) : wb.setAutoFilter(si, area)), 'autoFilter') === true;
    },
    /** Home > Sort & Filter > Clear (every column's filter taken off, the buttons kept), and Reapply (the rows
     *  filtered again, after edits). */
    filterClear() {
      const f = sheet().autoFilter;
      if (f) return f.columns.size > 0 && commands.filterAll(new Map());
      // (An Advanced Filter in place: its list's rows shown again, as Excel's Clear does.)
      const wb = ctx.workbook();
      const si = ctx.sheetIndex();
      const db = wb.definedNames.find((d) => d.name === '_xlnm._FilterDatabase' && d.sheet === si);
      const area = db && parseRange(db.text.replace(/^.*!/, '').replace(/\$/g, ''));
      if (!area || ctx.readOnly()) return false;
      return commands.sheetChange((book, k) => { book.setLines(k, 'row', area.r0 + 1, area.r1, { hidden: false }); });
    },
    filterReapply() { const f = sheet().autoFilter; return !!f && commands.filterAll(f.columns); },
    filterAll(columns) {
      const f = sheet().autoFilter;
      if (!f || ctx.readOnly()) return false;
      return moved((wb, si) => wb.setAutoFilter(si, { r0: f.r0, c0: f.c0, r1: f.r1, c1: f.c1, columns: new Map(columns) }), 'autoFilter') === true;
    },
    /** A filter button's Sort A to Z (or Z to A, `desc`): the filter's range by column c, its heading row kept. Not
     *  yet while rows are filtered out (Excel then sorts only the rows shown). */
    filterSort(c, desc, table = null) {
      const sh = sheet();
      const f = table ? readTables(ctx.workbook()).find((x) => x.name === table)?.filter : sh.autoFilter;
      if (!f) return false;
      for (let r = f.r0 + 1; r <= f.r1; r++) if (sh.rows.get(r)?.hidden) { ctx.announce("Sorting while rows are filtered out isn't available yet"); return false; }
      return commands.sort({ area: { r0: f.r0, c0: f.c0, r1: f.r1, c1: f.c1 }, keys: [{ at: c, desc }], header: true });
    },
  };
}
