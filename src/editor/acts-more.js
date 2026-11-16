// More of the ribbon's actions (acts.js): Insert > Charts and a chart's edits, comment threads, the Format Cells box and
// Home > Conditional Formatting (its menu's rules, Clear Rules, the Rules Manager). Each runs through the editor's
// commands as one step of the history.
import { messageBox, numberBox, optionsBox } from './dialogs.js';
import { iterationBox } from './iterbox.js';
import { pageSetupBox } from './pagesetupbox.js';
import { printAreaOf, printTitlesOf } from '../ooxml/pagesetupedit.js';
import { iteration } from '../ooxml/calcmode.js';
import { cfRuleBox } from './datadialogs.js';
import { readNotes } from '../ooxml/notes.js';
import { deleteNote } from '../ooxml/noteedit.js';
import { colName, rangeName } from '../ooxml/refs.js';
import { areasOf } from './selection.js';
import { CF_PRESETS, CF_PERIODS, CF_GALLERY } from '../ooxml/cfedit.js';
import { formatCellsBox } from './formatbox.js';
import { createStyleActs, STYLE_ACTS } from './acts-styles.js';
import { threadBox } from './threadbox.js';
import { mentionPeople } from './mentions.js';
import { addThread, replyThread, editThread, resolveThread, workbookPeople } from '../ooxml/threadedit.js';
import { insertChart } from '../ooxml/chartnew.js';
import { BAR_COLORS } from '../ooxml/cfbarsedit.js';
import { setChartTitle, setChartLegend, setChartLabels, setChartKind, setChartAnchor, deleteChart } from '../ooxml/chartedit.js';
import { setObjectAnchor, deleteObjects } from '../ooxml/drawobjects.js';
import { chartBox } from './chartbox.js';
import { createChartDataActs } from './acts-chartdata.js';
import { createChartActs } from './chartacts.js';
import { currentRegion } from '../ooxml/special.js';
import { xfOf, fontFamily } from '../ooxml/styles.js';
import { newRuleBox, rulesManager } from './cfmanager.js';
import { advancedFilterBox } from './advfilterbox.js';
import { seriesBox, removeDuplicatesBox, textToColumnsBox, moveCopyBox } from './datatoolboxes.js';
import { fillSeries, removeDuplicates, textToColumns } from '../ooxml/datatools.js';
import { flashFill } from '../ooxml/flashfill.js';
import { justifySteps, JUSTIFY_EXTENDS } from '../ooxml/justifyops.js';
import { invalidCells } from '../ooxml/validation.js';
import { excelTextWidth, JUSTIFY_MARGIN } from './textwidth.js';
import { sheetGeometry } from './geometry.js';
import { deleteRule, moveRule, setStopIfTrue, setRuleRanges, editRule } from '../ooxml/cfmanage.js';
import { insertFunctionBox } from './functionbox.js';
import { functionArgumentsBox } from './fnargsbox.js';
import { callForCaret } from '../calc/fnargs.js';
import { pasteNameBox, applyNamesBox } from './namesbox.js';
import { applicableNames, applyNamesSteps } from '../ooxml/applynames.js';
import { openErrorMenu } from './errormenu.js';
import { errorCheckingBox, errorOptionsBox } from './errorbox.js';
import { evaluateFormulaBox } from './evalbox.js';
import { protectBox, passwordAsk, workbookProtectBox } from './protectbox.js';
import { editRangesBox, refersOf } from './rangesbox.js';
import { linkBox } from './linkbox.js';
import { hyperlinkAt } from '../ooxml/links.js';
import { listNames, namesOnSheet } from '../ooxml/names.js';
import { sheetProtection, workbookProtection } from '../ooxml/protection.js';
import { DEFAULT_RECENT } from '../calc/catalog.js';
import { readPivots, pivotAt } from '../ooxml/pivot.js';

/** The Font tab's list: the theme's fonts first, then these. */
const FONT_NAMES = ['Aptos', 'Aptos Narrow', 'Arial', 'Calibri', 'Cambria', 'Consolas', 'Courier New', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Verdana'];

/** The actions: `env` is acts.js's (mountXlsxEditor's state and helpers) with its `act`. Returns run(name, args, i):
 *  what the action returns, or undefined for one that isn't here. */
export function createMoreActs(env) {
  const { root, state, options, editing, act } = env;
  const chartData = createChartDataActs(env);
  const chartActs = createChartActs(env, FONT_NAMES);
  const styleActs = createStyleActs(env, FONT_NAMES);
  return (name, args, i) => {
    // (A chart's data, Select Data: acts-chartdata.js; its formatting: chartacts.js.)
    const data = chartData(name, args);
    if (data !== undefined) return data;
    const chart = name.startsWith('chart') ? chartActs(name, args) : undefined;
    if (chart !== undefined) return chart;
    // (Excel's error button's menu, at a point: errormenu.js.)
    if (name === 'errorMenu') return openErrorMenu(env, args[0]);
    // (Formulas > Error Checking: errorbox.js.)
    if (name === 'errorChecking') return errorCheckingBox(env);
    if (name === 'errorOptions') {
      // (Excel's Error Checking Options: kept in the browser, as Excel keeps them as its own settings.)
      return errorOptionsBox(root, state.errorOptions).then((got) => {
        if (got) { state.errorOptions = got; try { localStorage.setItem('oxe-error-options', JSON.stringify(got)); } catch { /* storage blocked */ } state.refreshErrors?.(); }
        state.grid.el.focus({ preventScroll: true });
        return !!got;
      });
    }
    // (Formulas > Evaluate Formula: evalbox.js.)
    if (name === 'evaluateFormula') return evaluateFormulaBox(env);
    if (name === 'scaleBox') {
      // (Page Layout > Scale to Fit > Scale: a percentage from 10 to 400, as Excel's; it takes fit to pages off.)
      if (state.readOnly) return false;
      const p = state.workbook.sheet(i).page;
      return numberBox(root, { title: 'Scale', label: 'Scale (%):', value: String(p.fitToPage ? 100 : p.scale), min: 10, max: 400, refusal: 'Enter a scale from 10 to 400.' }).then((n) => {
        state.grid.el.focus({ preventScroll: true });
        return n == null ? false : editing().run('pageLayout', { scale: Math.round(n) });
      });
    }
    if (name === 'printTitlesBox' || name === 'customMarginsBox') {
      // (Page Layout > Print Titles, and Margins > Custom Margins...: Excel's Page Setup box at its Sheet or Margins tab,
      // pagesetupbox.js.)
      if (state.readOnly) return false;
      const abs = (a) => rangeName(a).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2');
      const t = printTitlesOf(state.workbook, i);
      const page = state.workbook.sheet(i).page;
      const now = {
        margins: page.margins, center: { h: page.centerH, v: page.centerV }, pageOrder: page.pageOrder,
        area: (printAreaOf(state.workbook, i) || []).map(abs).join(','),
        rows: t.rows ? `$${t.rows[0] + 1}:$${t.rows[1] + 1}` : '',
        cols: t.cols ? `$${colName(t.cols[0])}:$${colName(t.cols[1])}` : '',
      };
      return pageSetupBox(root, now, { tab: name === 'customMarginsBox' ? 'margins' : 'sheet' }).then((got) => {
        state.grid.el.focus({ preventScroll: true });
        return got ? editing().run('pageSetupSheet', got) : false;
      });
    }
    if (name === 'iterationBox') {
      // (Excel's File > Options > Formulas, its iterative calculation part: iterbox.js.)
      if (state.readOnly) return false;
      return iterationBox(root, iteration(state.workbook)).then((got) => {
        state.grid.el.focus({ preventScroll: true });
        return got ? editing().run('iteration', got) : false;
      });
    }
    // (Formulas > Watch Window, shown or hidden: watchwindow.js.)
    if (name === 'watchWindow') { state.watch?.toggle(); return true; }
    // (Data > Queries & Connections, shown or hidden: queriespane.js.)
    if (name === 'queriesPane') { state.queries?.toggle(); return true; }
    if (name === 'refreshData') {
      // (Refresh All (Ctrl+Alt+F5, args[0] true) and Refresh (Alt+F5): the queries and connections loaded to tables,
      // through the host, then the PivotTables; Refresh, the one at the active cell. Neither there: the PivotTables'
      // command says so.)
      const all = !!args[0];
      const s = editing().selection;
      const wb = state.workbook;
      const pivots = all ? readPivots(wb).length > 0 : !!pivotAt(wb, state.sheetIndex, s.ar, s.ac);
      const tables = state.queries?.refreshTables(all, state.sheetIndex, s.ar, s.ac);
      if (!tables) return editing().run('refreshPivots', all);
      return Promise.resolve(tables).then((done) => (pivots ? editing().run('refreshPivots', all) : done));
    }
    if (name === 'insertFunction' || name === 'functionArguments') {
      // (Formulas > Insert Function, Shift+F3, AutoSum's More Functions, or a function from the Function Library's
      // lists (args[0]): put first in Most Recently Used, then Function Arguments for it, as Excel goes on. Function
      // Arguments alone (Ctrl+A after a function's name in a cell being edited): for that call, its arguments in the
      // boxes. args[1], the boxes' values for the API.)
      const ed = state.editing;
      if (state.readOnly || !ed) return false;
      const editingNow = ed.editorOpen;
      const text = editingNow ? ed.editor.el.value : '';
      const caret = editingNow ? [ed.editor.el.selectionStart, ed.editor.el.selectionEnd] : [0, 0];
      const call = name === 'functionArguments' && !args[0] ? callForCaret(text, caret[0], { inside: false }) : null;
      if (name === 'functionArguments' && !args[0] && !call) return false;
      ed.holdEditor(true);
      const recent = state.recentFunctions || DEFAULT_RECENT;
      const back = () => { ed.holdEditor(false); if (!editingNow) state.grid.el.focus({ preventScroll: true }); return false; };
      const chosen = call ? Promise.resolve(call.name) : args[0] ? Promise.resolve(args[0]) : insertFunctionBox(root, { recent });
      return chosen.then((fn) => {
        if (!fn) return back();
        if (name === 'insertFunction') state.recentFunctions = [fn, ...recent.filter((n) => n !== fn)].slice(0, 10);
        const at = { si: state.sheetIndex, r: editing().selection.ar, c: editing().selection.ac };
        state.workbook.calc.build();
        const values = call ? call.args : (args[1] || []);
        // (The sheet pointed at from its boxes, as the formula bar's pointing; a name defined isn't put in quotes.)
        const defined = new Set([...namesOnSheet(state.workbook, at.si).map((n) => n.name), ...(state.workbook.calc.tables || []).map((t) => t.name)].map((n) => n.toLowerCase()));
        const pointer = (input) => ed.pointWith(input, at);
        const isName = (n) => defined.has(String(n).toLowerCase());
        return functionArgumentsBox(root, { name: fn, values, calc: state.workbook.calc, at, pointer, isName }).then((got) => {
          // (Back on the cell's own sheet if another was shown to point at, then pointing back with the cell editor.)
          if (state.sheetIndex !== at.si) editing().run('showSheet', at.si);
          ed.pointWith(null);
          if (!got) return back();
          // (The call in place of the name and arguments it opened on, or at the caret; "=" first in an empty cell.)
          const [from, to] = call ? [call.start, call.end] : caret;
          const formula = editingNow ? `${text.slice(0, from)}${text === '' ? '=' : ''}${got.text}${text.slice(to)}` : `=${got.text}`;
          ed.holdEditor(false);
          // (Entered as one step; one Excel would refuse, =SUM() with nothing in it, left being edited.)
          ed.putFormula(formula, got.refused);
          return !got.refused;
        });
      });
    }
    if (name === 'useName' || name === 'pasteNames') {
      // (Use in Formula's names (args[0]) and Paste Names (F3): the name typed at the caret of the cell being edited,
      // or the active cell opened with "=Name"; Paste List writes them all (pasteNameList). No names, no box.)
      const ed = state.editing;
      const names = namesOnSheet(state.workbook, i);
      if (state.readOnly || !ed || !names.length) return false;
      const editingNow = ed.editorOpen;
      ed.holdEditor(true);
      const chosen = name === 'useName' ? Promise.resolve({ name: args[0] }) : pasteNameBox(root, names.map((n) => n.name), { list: !editingNow });
      return chosen.then((got) => {
        // (Typed before the editor is let go: it keeps the caret only while held.)
        if (got?.name) ed.typeFunction(got.name, '');
        ed.holdEditor(false);
        if (!got || got.list) state.grid.el.focus({ preventScroll: true });
        if (!got) return false;
        return got.list ? editing().run('pasteNameList') : true;
      });
    }
    if (name === 'linkBox') {
      // (Insert > Link, Ctrl+K: Excel's Insert Hyperlink box, Edit Hyperlink on a cell with one.)
      const s = editing().selection;
      const cell = state.workbook.sheet(i).cells.get(s.ar)?.get(s.ac);
      const text = cell == null || cell.v == null ? '' : cell.f == null && typeof cell.v === 'string' ? cell.v : null;
      const link = hyperlinkAt(state.workbook, i, s.ar, s.ac);
      const sheets = state.workbook.sheets.filter((x) => x.kind === 'worksheet').map((x) => x.name);
      const names = listNames(state.workbook).filter((n) => n.sheet == null).map((n) => n.name);
      return linkBox(root, { link, text, sheets, names, sheet: state.workbook.sheets[i].name }).then((answer) => {
        state.grid.el.focus({ preventScroll: true });
        if (!answer) return false;
        if (answer.remove) return editing().run('clearAs', 'removeHyperlinks');
        const done = editing().run('insertHyperlink', s.ar, s.ac, answer);
        // (Create New Document: the link written as Excel writes it; the document itself the host's to make, as a page
        // can't make files.)
        if (done && answer.create) options.onCreateDocument?.(answer.create.name, { editNow: answer.create.editNow });
        return done;
      });
    }
    if (name === 'editRangesBox') {
      // (Allow Edit Ranges: Excel's box, a new range offered the selection's cells.)
      const refers = refersOf(areasOf(editing().selection));
      return editRangesBox(root, { workbook: () => state.workbook, si: i, run: (cmd, ...a) => editing().run(cmd, ...a), refers, locked: !!sheetProtection(state.workbook, i) });
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
      // (@ offers the host's people and the workbook's own.)
      const people = mentionPeople(typeof options.people === 'function' ? options.people() : options.people, workbookPeople(state.workbook));
      return threadBox(root, { ref, items: n?.items || [], at: { x: box.right + 8, y: box.top }, readOnly: state.readOnly, resolved: !!n?.resolved, people }).then((answer) => {
        state.grid.el.focus({ preventScroll: true });
        if (!answer) return false;
        const who = options.author || 'Author';
        let refused = null;
        editing().run('sheetChange', (wb, si) => {
          try {
            if (answer.removeAll) return deleteNote(wb, si, s.ar, s.ac);
            if (answer.resolve != null) return resolveThread(wb, si, s.ar, s.ac, answer.resolve);
            const mentions = answer.mentions || [];
            if (answer.post != null) { if (n) replyThread(wb, si, s.ar, s.ac, answer.post, who, { mentions }); else addThread(wb, si, s.ar, s.ac, answer.post, who, { mentions }); }
            else if (answer.remove != null) editThread(wb, si, s.ar, s.ac, answer.remove, null);
            else if (answer.edit != null) editThread(wb, si, s.ar, s.ac, answer.edit, answer.text, { mentions });
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
      const add = (rule) => editing().run('sheetChange', (wb, si) => wb.setConditional(si, sel, rule), 'formatCells');
      const presets = Object.entries(CF_PRESETS).map(([value, p]) => ({ value, label: p.label }));
      const kind = args[0];
      if (kind === 'scale3' || kind === 'scale2') return add({ kind: 'colorScale', points: kind === 'scale2' ? 2 : 3 });
      // (Excel's Color Scales gallery: cfedit.js CF_GALLERY.)
      if (kind === 'gallery') { const colors = CF_GALLERY[args[1]]?.[1]; return colors ? add({ kind: 'colorScale', points: colors.length, colors }) : false; }
      // (Data Bars and Icon Sets: added at once, as Excel's galleries add them.)
      if (kind === 'dataBar') return add({ kind: 'dataBar', color: BAR_COLORS[args[1]], solid: args[2] === 'solid' });
      if (kind === 'iconSet') return add({ kind: 'iconSet', set: args[1] });
      const WORDS = { greaterThan: 'GREATER THAN', lessThan: 'LESS THAN', between: 'BETWEEN', equal: 'EQUAL TO', containsText: 'CONTAIN THE TEXT' };
      const TITLES = { greaterThan: 'Greater Than', lessThan: 'Less Than', between: 'Between', equal: 'Equal To', containsText: 'Text That Contains', timePeriod: 'A Date Occurring', duplicate: 'Duplicate Values', top: 'Top 10 Items', topPercent: 'Top 10%', bottom: 'Bottom 10 Items', bottomPercent: 'Bottom 10%', above: 'Above Average', below: 'Below Average' };
      const ranked = ['top', 'topPercent', 'bottom', 'bottomPercent'].includes(kind);
      const text = WORDS[kind] ? `Format cells that are ${WORDS[kind]}:` : ranked ? `Format cells that rank in the ${kind.startsWith('top') ? 'TOP' : 'BOTTOM'}:` : kind === 'duplicate' || kind === 'timePeriod' ? '' : `Format cells that are ${kind === 'above' ? 'ABOVE' : 'BELOW'} AVERAGE for the selected range`;
      const inputs = kind === 'between' ? [{ label: 'From', value: '' }, { label: 'and', value: '' }] : WORDS[kind] ? [{ label: 'Value', value: '' }] : ranked ? [{ label: kind.endsWith('Percent') ? 'Percent' : 'Items', value: '10' }] : [];
      return cfRuleBox(root, { title: TITLES[kind], text, inputs, choice: kind === 'duplicate' ? [{ value: 'duplicate', label: 'Duplicate' }, { value: 'unique', label: 'Unique' }] : kind === 'timePeriod' ? Object.entries(CF_PERIODS).map(([value, [label]]) => ({ value, label })) : null, choiceLabel: kind === 'timePeriod' ? 'Format cells that contain a date occurring:' : undefined, presets, preset: kind === 'below' ? 'yellow' : 'lightRed' }).then((a) => {
        if (a) {
          if (WORDS[kind] && kind !== 'containsText') add({ kind: 'cellIs', operator: kind, f1: a.values[0], f2: a.values[1], preset: a.preset });
          else if (kind === 'containsText') add({ kind: 'containsText', text: a.values[0], preset: a.preset });
          else if (kind === 'timePeriod') add({ kind: 'timePeriod', period: a.choice, preset: a.preset });
          else if (kind === 'duplicate') add({ kind: 'duplicate', unique: a.choice === 'unique', preset: a.preset });
          else if (ranked) add({ kind: 'top10', rank: Math.max(1, Math.round(Number(a.values[0]) || 10)), percent: kind.endsWith('Percent'), bottom: kind.startsWith('bottom'), preset: a.preset });
          else add({ kind: 'average', below: kind === 'below', preset: a.preset });
        }
        state.grid.el.focus({ preventScroll: true });
      });
    }
    if (name === 'cfClear') {
      const s = editing().selection;
      return editing().run('sheetChange', (wb, si) => wb.setConditional(si, args[0] === 'sheet' ? null : [{ r0: s.r0, c0: s.c0, r1: s.r1, c1: s.c1 }], null), 'formatCells');
    }
    if (name === 'insertChart') {
      // (Insert > Charts: the selection charted; one cell selected, the table round it, as Excel takes it.)
      const s = editing().selection;
      const sel = s.r0 === s.r1 && s.c0 === s.c1 ? currentRegion(state.workbook.sheet(i), s.r0, s.c0) : s;
      let made = null;
      // (Excel refused a new chart on a protected sheet, Edit objects allowed or not.)
      const done = editing().run('sheetChange', (wb, si) => { made = insertChart(wb, si, sel, args[0]); return !!made; }, 'always');
      if (!done && !made && !state.readOnly) return messageBox(root, 'Select the cells to chart first: a table with numbers in it.').then(() => false);
      return done;
    }
    if (name === 'chartAnchor') return editing().run('sheetChange', (wb, si) => { setChartAnchor(wb, si, args[0], args[1], args[2]); }, 'objects');
    if (name === 'chartDelete') return editing().run('sheetChange', (wb, si) => { deleteChart(wb, si, args[0]); }, 'objects');
    // (Pictures, shapes and charts moved or sized together, and deleted together: each one step, drawobjects.js.)
    if (name === 'drawingsPlace') {
      return editing().run('sheetChange', (wb, si) => {
        for (const p of args[0]) if (p.chart != null) setChartAnchor(wb, si, p.chart, p.from, p.to); else setObjectAnchor(wb, si, p.k, p.from, p.to, p.rect);
      }, 'objects');
    }
    if (name === 'drawingsDelete') return editing().run('sheetChange', (wb, si) => { deleteObjects(wb, si, args[0]); }, 'objects');
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
      }, 'objects');
      if (refused) return messageBox(root, refused).then(() => false);
      state.picker?.pick(k, false);
      return done;
    }
    if (name === 'chartBox') {
      const chart = (state.charts || []).find((c) => c.index === args[0]);
      if (!chart) return false;
      return chartBox(root, chart).then((change) => {
        state.grid.el.focus({ preventScroll: true });
        if (change?.selectData) return act('chartDataBox', args[0]);
        return change ? act('chartEdit', args[0], change) : false;
      });
    }
    // (New Cell Style... and a style's Modify...: acts-styles.js.)
    if (STYLE_ACTS.has(name)) return styleActs(name, args);
    if (name === 'formatCellsBox' && state.editing?.textSelected) {
      // (Ctrl+1 with characters selected in the cell editor or the formula bar: the box with its Font tab alone, on
      // the first selected character's font, and what was changed there given to those characters only, as Excel's.)
      const wb = state.workbook;
      const ed = state.editing;
      const s = editing().selection;
      const cell = wb.sheet(i).cells.get(s.ar)?.get(s.ac) || null;
      const xf = xfOf(wb.styles, cell?.s || 0);
      const own = ed.selectedFont || {};
      const font = { ...wb.styles.fonts[xf.fontId] };
      for (const [k, v] of Object.entries(own)) if (v != null) font[k] = k === 'color' && typeof v === 'string' ? { rgb: v } : v;
      const look = { xf, font, fontName: own.name || fontFamily(wb.styles.fonts[xf.fontId], wb.theme), fontNames: [...new Set([wb.theme.fonts.minor, wb.theme.fonts.major, ...FONT_NAMES])],
        part: true, color: (c) => state.looks?.color(c) };
      ed.holdText(true);
      return formatCellsBox(root, look, 'Font').then((change) => {
        // (Automatic is theme colour 1, as the Font group's Automatic on part of the text.)
        const f = change?.font ? { ...change.font } : null;
        if (f && 'color' in f && !f.color) f.color = { theme: 1 };
        const done = !!f && ed.markEditorFont(f);
        ed.holdText(false);
        return done;
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
    if (name === 'applyNames') {
      // (Formulas > Define Name > Apply Names...: the names chosen put in place of the references to them in the
      // selection's formulas, the whole sheet's for one cell (applynames.js). The two messages are the editor's own:
      // Excel's weren't measured.)
      if (state.readOnly) return false;
      const wb = state.workbook;
      const usable = applicableNames(wb, i);
      if (!usable.length) return messageBox(root, 'There are no names here that refer to cells of this sheet.').then(() => false);
      return applyNamesBox(root, usable.map((n) => n.name)).then((o) => {
        state.grid.el.focus({ preventScroll: true });
        if (!o) return false;
        const got = applyNamesSteps(wb, i, editing().selection, usable.filter((n) => o.names.includes(n.name)), o);
        if (got.none) return messageBox(root, 'No references were found to replace with these names.').then(() => false);
        return editing().run('runTool', got.steps, `Names applied in ${got.count} cell${got.count === 1 ? '' : 's'}, `);
      });
    }
    if (name === 'circleInvalid' || name === 'clearCircles') {
      // (Data Validation's menu: a circle round each cell of the sheet shown that breaks its rule, kept with the
      // sheet, not saved; or none. index.js drops one when its cell is put right, and all of them once lines or cells
      // are inserted or deleted on the sheet, as Excel's go: `lines` is the sheet's count of those then.)
      const info = state.workbook.sheets[i];
      const cells = name === 'circleInvalid' ? invalidCells(state.workbook, i) : [];
      cells.lines = state.workbook.lineChanges.get(info.sheetId) || 0;
      state.circles.set(info, cells);
      state.grid.setCircles(cells);
      return true;
    }
    if (name === 'fillJustify') {
      // (Home > Fill > Justify: the first column's text rewrapped to the selection's width, measured in its first
      // cell's font as Excel measures it; Excel's refusal, or its question when the text goes on below.)
      if (state.readOnly) return false;
      const s = editing().selection;
      const wb = state.workbook;
      const cols = sheetGeometry(wb.sheet(i)).cols;
      const room = cols.start(s.c1 + 1) - cols.start(s.c0) - JUSTIFY_MARGIN;
      const font = state.looks.deco(wb.sheet(i).cells.get(s.r0)?.get(s.c0)?.s || 0).font;
      const got = justifySteps(wb, i, s, (line) => excelTextWidth(line, font.css, font.sizePx) <= room);
      if (got.error) return messageBox(root, got.error).then(() => false);
      const go = got.extends ? messageBox(root, JUSTIFY_EXTENDS, ['OK', 'Cancel']).then((b) => b === 'OK') : Promise.resolve(true);
      return go.then((ok) => {
        state.grid.el.focus({ preventScroll: true });
        return ok ? editing().run('runTool', got.steps, 'Justified ') : false;
      });
    }
    if (name === 'fillAcrossBox') {
      // (Home > Fill > Across Worksheets, sheets grouped: Excel's box, All, Contents or Formats.)
      if (state.readOnly) return false;
      return optionsBox(root, { title: 'Fill Across Worksheets', groups: [{ key: 'fill', legend: 'Fill', options: ['All', 'Contents', 'Formats'], value: 'All' }] }).then((got) => {
        state.grid.el.focus({ preventScroll: true });
        return got ? editing().run('fillAcross', got.choices.fill.toLowerCase()) : false;
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
      const lines = [];
      for (let r = s.r0; r <= s.r1 && lines.length < 5; r++) { const v = state.workbook.sheet(i).cells.get(r)?.get(s.c0)?.v; if (v != null) lines.push(String(v)); }
      return textToColumnsBox(root, { dest: `$${colName(s.c0)}$${s.r0 + 1}`, lines }).then((o) => {
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
        return rule ? editing().run('sheetChange', (wb, si) => wb.setConditional(si, areasOf(s).map(({ r0, c0, r1, c1 }) => ({ r0, c0, r1, c1 })), rule), 'formatCells') : false;
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
              if (change.edit) editRule(wb, si, change.edit.p, change.edit.rule);
            });
          } catch (e) { refused = e.message; return false; }
          return true;
        }, 'formatCells');
        return refused;
      };
      const rules = () => state.workbook.sheet(state.sheetIndex).condFormats.flatMap((cf) => cf.rules.map((rule) => ({ rule, ranges: cf.ranges })));
      return rulesManager(root, { rules, selection: areasOf(s), run, newRule: () => act('cfNewRule') }).then(() => state.grid.el.focus({ preventScroll: true }));
    }
    return undefined;
  };
}
