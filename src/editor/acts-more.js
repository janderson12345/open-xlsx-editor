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
import { deleteRule, moveRule, setStopIfTrue, setRuleRanges } from '../ooxml/cfmanage.js';

/** The Font tab's list: the theme's fonts first, then these. */
const FONT_NAMES = ['Aptos', 'Aptos Narrow', 'Arial', 'Calibri', 'Cambria', 'Consolas', 'Courier New', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Verdana'];

/** The actions: `env` is acts.js's (mountXlsxEditor's state and helpers) with its `act`. Returns run(name, args, i):
 *  what the action returns, or undefined for one that isn't here. */
export function createMoreActs(env) {
  const { root, state, options, editing, act } = env;
  return (name, args, i) => {
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
