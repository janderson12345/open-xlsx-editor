// More of the ribbon's actions (acts.js): Insert > Charts and a chart's edits, comment threads, the Format Cells box and
// Home > Conditional Formatting (its menu's rules, Clear Rules, the Rules Manager). Each runs through the editor's
// commands as one step of the history.
import { messageBox } from './dialogs.js';
import { cfRuleBox } from './datadialogs.js';
import { readNotes } from '../ooxml/notes.js';
import { deleteNote } from '../ooxml/noteedit.js';
import { colName } from '../ooxml/refs.js';
import { areasOf } from './selection.js';
import { CF_PRESETS } from '../ooxml/cfedit.js';
import { formatCellsBox } from './formatbox.js';
import { threadBox } from './threadbox.js';
import { addThread, replyThread, editThread } from '../ooxml/threadedit.js';
import { insertChart } from '../ooxml/chartnew.js';
import { BAR_COLORS } from '../ooxml/cfbarsedit.js';
import { setChartTitle, setChartLegend, setChartLabels, setChartKind, setChartAnchor, deleteChart } from '../ooxml/chartedit.js';
import { chartBox } from './chartbox.js';
import { currentRegion } from '../ooxml/special.js';
import { xfOf, fontFamily } from '../ooxml/styles.js';
import { newRuleBox, rulesManager } from './cfmanager.js';
import { advancedFilterBox } from './advfilterbox.js';
import { seriesBox, removeDuplicatesBox, textToColumnsBox, moveCopyBox } from './datatoolboxes.js';
import { fillSeries, removeDuplicates, textToColumns } from '../ooxml/datatools.js';
import { flashFill } from '../ooxml/flashfill.js';
import { deleteRule, moveRule, setStopIfTrue, setRuleRanges } from '../ooxml/cfmanage.js';
import { insertFunctionBox } from './functionbox.js';
import { openErrorMenu } from './errormenu.js';
import { errorCheckingBox } from './errorbox.js';
import { evaluateFormulaBox } from './evalbox.js';
import { protectBox, passwordAsk, workbookProtectBox } from './protectbox.js';
import { sheetProtection, workbookProtection } from '../ooxml/protection.js';
import { DEFAULT_RECENT } from '../calc/catalog.js';

/** The Font tab's list: the theme's fonts first, then these. */
const FONT_NAMES = ['Aptos', 'Aptos Narrow', 'Arial', 'Calibri', 'Cambria', 'Consolas', 'Courier New', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Verdana'];

/** The actions: `env` is acts.js's (mountXlsxEditor's state and helpers) with its `act`. Returns run(name, args, i):
 *  what the action returns, or undefined for one that isn't here. */
export function createMoreActs(env) {
  const { root, state, options, editing, act } = env;
  return (name, args, i) => {
    // (Excel's error button's menu, at a point: errormenu.js.)
    if (name === 'errorMenu') return openErrorMenu(env, args[0]);
    // (Formulas > Error Checking: errorbox.js.)
    if (name === 'errorChecking') return errorCheckingBox(env);
    // (Formulas > Evaluate Formula: evalbox.js.)
    if (name === 'evaluateFormula') return evaluateFormulaBox(env);
    // (Formulas > Watch Window, shown or hidden: watchwindow.js.)
    if (name === 'watchWindow') { state.watch?.toggle(); return true; }
    if (name === 'insertFunction') {
      // (Formulas > Insert Function, Shift+F3, or a function from the Function Library's lists (args[0]): typed at the
      // caret of the cell being edited, or the active cell opened with "=NAME(", and put first in Most Recently Used.)
      const ed = state.editing;
      if (state.readOnly || !ed) return false;
      const editingNow = ed.editorOpen;
      ed.holdEditor(true);
      const recent = state.recentFunctions || DEFAULT_RECENT;
      const chosen = args[0] ? Promise.resolve(args[0]) : insertFunctionBox(root, { recent });
      return chosen.then((fn) => {
        if (!fn) { ed.holdEditor(false); if (!editingNow) state.grid.el.focus({ preventScroll: true }); return false; }
        state.recentFunctions = [fn, ...recent.filter((n) => n !== fn)].slice(0, 10);
        ed.typeFunction(fn);
        ed.holdEditor(false);
        return true;
      });
    }
    if (name === 'protectWorkbookBox') {
      // (Protect Workbook: Excel's box, a password asked for twice; on a protected workbook, unprotected.)
      const p = workbookProtection(state.workbook);
      if (p) {
        if (!p.password) return editing().run('unprotectWorkbook', '');
        return passwordAsk(root, { title: 'Unprotect Workbook', label: 'Password:' }).then((pw) => pw != null && editing().run('unprotectWorkbook', pw));
      }
      return workbookProtectBox(root).then(async (pw) => {
        if (pw == null) return false;
        if (pw) {
          const again = await passwordAsk(root, { title: 'Confirm Password', label: 'Reenter password to proceed.' });
          if (again == null) return false;
          if (again !== pw) { await messageBox(root, 'Confirmation password is not identical.'); return false; }
        }
        return editing().run('protectWorkbook', pw);
      });
    }
    if (name === 'protectBox') {
      // (Protect Sheet, or Unprotect Sheet once it is: Excel's boxes, a password asked for twice.)
      const p = sheetProtection(state.workbook, i);
      if (p) {
        if (!p.password) return editing().run('unprotectSheet', '');
        return passwordAsk(root, { title: 'Unprotect Sheet', label: 'Password:' }).then((pw) => pw != null && editing().run('unprotectSheet', pw));
      }
      return protectBox(root).then(async (answer) => {
        if (!answer) return false;
        if (answer.password) {
          const again = await passwordAsk(root, { title: 'Confirm Password', label: 'Reenter password to proceed.' });
          if (again == null) return false;
          if (again !== answer.password) { await messageBox(root, 'Confirmation password is not identical.'); return false; }
        }
        return editing().run('protectSheet', { ...answer.options, password: answer.password });
      });
    }
    if (name === 'commentBox') {
      // (Excel's comment thread on the active cell: started, replied to, a comment edited or a reply deleted, the whole
      // thread deleted; each one step of Undo, the box opened again on the thread after each.)
      const s = editing().selection;
      const ref = `${colName(s.ac)}${s.ar + 1}`;
      const n = readNotes(state.workbook, i).get(`${s.ar},${s.ac}`);
      if (n?.kind === 'note') return messageBox(root, `${ref} has a note: a cell has a note or a comment, not both.`);
      const box = (root.querySelector('.oxe-sel-active') || state.grid.el).getBoundingClientRect();
      return threadBox(root, { ref, items: n?.items || [], at: { x: box.right + 8, y: box.top }, readOnly: state.readOnly }).then((answer) => {
        state.grid.el.focus({ preventScroll: true });
        if (!answer) return false;
        const who = options.author || 'Author';
        let refused = null;
        editing().run('sheetChange', (wb, si) => {
          try {
            if (answer.removeAll) return deleteNote(wb, si, s.ar, s.ac);
            if (answer.post != null) { if (n) replyThread(wb, si, s.ar, s.ac, answer.post, who); else addThread(wb, si, s.ar, s.ac, answer.post, who); }
            else if (answer.remove != null) editThread(wb, si, s.ar, s.ac, answer.remove, null);
            else if (answer.edit != null) editThread(wb, si, s.ar, s.ac, answer.edit, answer.text);
          } catch (e) { refused = e.message; return false; }
          return true;
        });
        if (refused) return messageBox(root, refused).then(() => false);
        return answer.removeAll ? true : act('commentBox');
      });
    }
    if (name === 'cfAdd') {
      // (Excel's Conditional Formatting menu: the rule's box, then the rule for the selection.)
      const s = editing().selection;
      const sel = [{ r0: s.r0, c0: s.c0, r1: s.r1, c1: s.c1 }];
      const add = (rule) => editing().run('sheetChange', (wb, si) => wb.setConditional(si, sel, rule));
      const presets = Object.entries(CF_PRESETS).map(([value, p]) => ({ value, label: p.label }));
      const kind = args[0];
      if (kind === 'scale3' || kind === 'scale2') return add({ kind: 'colorScale', points: kind === 'scale2' ? 2 : 3 });
      // (Data Bars and Icon Sets: added at once, as Excel's galleries add them.)
      if (kind === 'dataBar') return add({ kind: 'dataBar', color: BAR_COLORS[args[1]], solid: args[2] === 'solid' });
      if (kind === 'iconSet') return add({ kind: 'iconSet', set: args[1] });
      const WORDS = { greaterThan: 'GREATER THAN', lessThan: 'LESS THAN', between: 'BETWEEN', equal: 'EQUAL TO', containsText: 'CONTAIN THE TEXT' };
      const TITLES = { greaterThan: 'Greater Than', lessThan: 'Less Than', between: 'Between', equal: 'Equal To', containsText: 'Text That Contains', duplicate: 'Duplicate Values', top: 'Top 10 Items', topPercent: 'Top 10%', bottom: 'Bottom 10 Items', bottomPercent: 'Bottom 10%', above: 'Above Average', below: 'Below Average' };
      const ranked = ['top', 'topPercent', 'bottom', 'bottomPercent'].includes(kind);
      const text = WORDS[kind] ? `Format cells that are ${WORDS[kind]}:` : ranked ? `Format cells that rank in the ${kind.startsWith('top') ? 'TOP' : 'BOTTOM'}:` : kind === 'duplicate' ? '' : `Format cells that are ${kind === 'above' ? 'ABOVE' : 'BELOW'} AVERAGE for the selected range`;
      const inputs = kind === 'between' ? [{ label: 'From', value: '' }, { label: 'and', value: '' }] : WORDS[kind] ? [{ label: 'Value', value: '' }] : ranked ? [{ label: kind.endsWith('Percent') ? 'Percent' : 'Items', value: '10' }] : [];
      return cfRuleBox(root, { title: TITLES[kind], text, inputs, choice: kind === 'duplicate' ? [{ value: 'duplicate', label: 'Duplicate' }, { value: 'unique', label: 'Unique' }] : null, presets, preset: kind === 'below' ? 'yellow' : 'lightRed' }).then((a) => {
        if (a) {
          if (WORDS[kind] && kind !== 'containsText') add({ kind: 'cellIs', operator: kind, f1: a.values[0], f2: a.values[1], preset: a.preset });
          else if (kind === 'containsText') add({ kind: 'containsText', text: a.values[0], preset: a.preset });
          else if (kind === 'duplicate') add({ kind: 'duplicate', unique: a.choice === 'unique', preset: a.preset });
          else if (ranked) add({ kind: 'top10', rank: Math.max(1, Math.round(Number(a.values[0]) || 10)), percent: kind.endsWith('Percent'), bottom: kind.startsWith('bottom'), preset: a.preset });
          else add({ kind: 'average', below: kind === 'below', preset: a.preset });
        }
        state.grid.el.focus({ preventScroll: true });
      });
    }
    if (name === 'cfClear') {
      const s = editing().selection;
      return editing().run('sheetChange', (wb, si) => wb.setConditional(si, args[0] === 'sheet' ? null : [{ r0: s.r0, c0: s.c0, r1: s.r1, c1: s.c1 }], null));
    }
    if (name === 'insertChart') {
      // (Insert > Charts: the selection charted; one cell selected, the table round it, as Excel takes it.)
      const s = editing().selection;
      const sel = s.r0 === s.r1 && s.c0 === s.c1 ? currentRegion(state.workbook.sheet(i), s.r0, s.c0) : s;
      let made = null;
      const done = editing().run('sheetChange', (wb, si) => { made = insertChart(wb, si, sel, args[0]); return !!made; });
      if (!done && !made && !state.readOnly) return messageBox(root, 'Select the cells to chart first: a table with numbers in it.').then(() => false);
      return done;
    }
    if (name === 'chartAnchor') return editing().run('sheetChange', (wb, si) => { setChartAnchor(wb, si, args[0], args[1], args[2]); });
    if (name === 'chartDelete') return editing().run('sheetChange', (wb, si) => { deleteChart(wb, si, args[0]); });
    if (name === 'chartEdit') {
      // (Edit Chart's answer: each change as Excel writes it, chartedit.js; refused for a chart written otherwise.)
      const [k, change] = args;
      let refused = null;
      const done = editing().run('sheetChange', (wb, si) => {
        try {
          if (change.kind) setChartKind(wb, si, k, change.kind);
          if ('title' in change) setChartTitle(wb, si, k, change.title);
          if ('legend' in change) setChartLegend(wb, si, k, change.legend);
          if ('labels' in change) setChartLabels(wb, si, k, change.labels);
        } catch (e) { refused = e.message; return false; }
        return true;
      });
      if (refused) return messageBox(root, refused).then(() => false);
      state.picker?.pick(k, false);
      return done;
    }
    if (name === 'chartBox') {
      const chart = (state.charts || []).find((c) => c.index === args[0]);
      if (!chart) return false;
      return chartBox(root, chart).then((change) => {
        state.grid.el.focus({ preventScroll: true });
        return change ? act('chartEdit', args[0], change) : false;
      });
    }
    if (name === 'formatCellsBox') {
      // (Ctrl+1: the active cell's format in Excel's box; what was changed there given to the selection.)
      const wb = state.workbook;
      const s = editing().selection;
      const cell = wb.sheet(i).cells.get(s.ar)?.get(s.ac) || null;
      const xf = xfOf(wb.styles, cell?.s || 0);
      const font = wb.styles.fonts[xf.fontId];
      const look = { xf, font, fontName: fontFamily(font, wb.theme), fontNames: [...new Set([wb.theme.fonts.minor, wb.theme.fonts.major, ...FONT_NAMES])],
        fill: wb.styles.fills[xf.fillId], border: wb.styles.borders[xf.borderId], code: editing().look().code, cell, date1904: wb.date1904,
        many: areasOf(s).some((a) => a.r1 > a.r0 || a.c1 > a.c0), color: (c) => state.looks?.color(c) };
      return formatCellsBox(root, look, args[0]).then((change) => {
        if (change) editing().run('formatCells', change);
        state.grid.el.focus({ preventScroll: true });
        return !!change;
      });
    }
    if (name === 'advancedFilterBox') {
      // (Data > Advanced: the table round the active cell offered as the list; Excel's words when it refuses.)
      const s = editing().selection;
      const list = currentRegion(state.workbook.sheet(i), s.ar, s.ac);
      return advancedFilterBox(root, { list }).then((spec) => {
        state.grid.el.focus({ preventScroll: true });
        if (!spec) return false;
        let refused = null;
        editing().run('sheetChange', (wb, si) => { refused = wb.advancedFilter(si, spec); return !refused; });
        return refused ? messageBox(root, refused).then(() => false) : true;
      });
    }
    if (name === 'seriesBox') {
      // (Home > Fill > Series...: the selection's first cells carried on.)
      const s = editing().selection;
      return seriesBox(root, { across: s.c1 - s.c0 > s.r1 - s.r0 }).then((o) => {
        state.grid.el.focus({ preventScroll: true });
        return o ? editing().run('runTool', fillSeries(state.workbook, i, s, o), 'Filled ') : false;
      });
    }
    if (name === 'flashFill') {
      // (Ctrl+E: the active cell's column, from its examples, over the rows of the table round it.)
      const s = editing().selection;
      const region = currentRegion(state.workbook.sheet(i), s.ar, s.ac);
      const sources = [];
      for (let c = region.c0; c <= region.c1; c++) if (c !== s.ac) sources.push(c);
      const got = flashFill(state.workbook, i, { col: s.ac, r0: region.r0 + 1, r1: region.r1, sourceCols: sources.filter((c) => c < s.ac).length ? sources.filter((c) => c < s.ac) : sources });
      if (got.error) return messageBox(root, got.error).then(() => false);
      return editing().run('runTool', got.steps, `Flash Fill changed ${got.filled} cells. `);
    }
    if (name === 'removeDuplicatesBox') {
      // (Data > Remove Duplicates: the table round the active cell, its columns by heading; Excel's message after.)
      const s = editing().selection;
      const wb = state.workbook;
      const sheet = wb.sheet(i);
      const range = s.r0 === s.r1 && s.c0 === s.c1 ? currentRegion(sheet, s.r0, s.c0) : s;
      const heads = [];
      for (let c = range.c0; c <= range.c1; c++) { const v = sheet.cells.get(range.r0)?.get(c)?.v; heads.push({ c, text: v == null ? '' : String(v) }); }
      const headers = heads.every((h) => h.text !== '' && Number.isNaN(Number(h.text)));
      return removeDuplicatesBox(root, { columns: heads.map((h) => ({ c: h.c, label: headers ? h.text : `Column ${colName(h.c)}` })), headers }).then((o) => {
        state.grid.el.focus({ preventScroll: true });
        if (!o) return false;
        const got = removeDuplicates(state.workbook, i, range, o.cols, o.header);
        editing().run('runTool', got.steps);
        const words = got.removed ? `${got.removed} duplicate values found and removed; ${got.kept} unique values remain.` : `No duplicate values found.`;
        return messageBox(root, words).then(() => true);
      });
    }
    if (name === 'textToColumnsBox') {
      // (Data > Text to Columns: one column of the selection, split across from it.)
      const s = editing().selection;
      if (s.c0 !== s.c1) return messageBox(root, 'Microsoft Excel can convert only one column at a time. The range can be many rows tall but no more than one column wide. Try again by selecting cells in one column only.');
      return textToColumnsBox(root, { dest: `$${colName(s.c0)}$${s.r0 + 1}` }).then((o) => {
        state.grid.el.focus({ preventScroll: true });
        return o ? editing().run('runTool', textToColumns(state.workbook, i, s, o), 'Converted ') : false;
      });
    }
    if (name === 'moveCopyBox') {
      // (A sheet tab's Move or Copy...: moved before the sheet chosen, or copied there.)
      const k = args[0] ?? i;
      const wb = state.workbook;
      return moveCopyBox(root, { name: wb.sheets[k].name, sheets: wb.sheets.map((x) => x.name), at: k }).then((o) => {
        state.grid.el.focus({ preventScroll: true });
        if (!o) return false;
        if (o.copy) return editing().run('copySheet', k, o.before);
        const to = o.before > k ? o.before - 1 : o.before;
        return to === k ? true : editing().run('moveSheet', k, Math.min(to, wb.sheets.length - 1));
      });
    }
    if (name === 'cfNewRule') {
      // (New Formatting Rule: the rule for the selection.)
      const s = editing().selection;
      return newRuleBox(root).then((rule) => {
        state.grid.el.focus({ preventScroll: true });
        return rule ? editing().run('sheetChange', (wb, si) => wb.setConditional(si, areasOf(s).map(({ r0, c0, r1, c1 }) => ({ r0, c0, r1, c1 })), rule)) : false;
      });
    }
    if (name === 'cfManager') {
      // (The Rules Manager: each change one step of the history, the list read again after it.)
      const s = editing().selection;
      const run = (change) => {
        let refused = null;
        editing().run('sheetChange', (wb, si) => {
          try {
            wb.changeConditional(si, () => {
              if (change.remove != null) deleteRule(wb, si, change.remove);
              if (change.move) moveRule(wb, si, change.move.p, change.move.to);
              if (change.stop) setStopIfTrue(wb, si, change.stop.p, change.stop.on);
              if (change.ranges) setRuleRanges(wb, si, change.ranges.p, change.ranges.ranges);
            });
          } catch (e) { refused = e.message; return false; }
          return true;
        });
        return refused;
      };
      const rules = () => state.workbook.sheet(state.sheetIndex).condFormats.flatMap((cf) => cf.rules.map((rule) => ({ rule, ranges: cf.ranges })));
      return rulesManager(root, { rules, selection: areasOf(s), run, newRule: () => act('cfNewRule') }).then(() => state.grid.el.focus({ preventScroll: true }));
    }
    return undefined;
  };
}
