// The status bar's figures for the selection (stats.js) and Customize Status Bar, as Excel's: index.js shares its state.
import { selectionStats } from '../ooxml/stats.js';
import { needsCalculation } from '../ooxml/calcmode.js';

/** The status bar for mountXlsxEditor (`state`, `options`, `menu`): show() its figures now, menu(point) Customize Status Bar. */
export function createStatusBar({ state, options, menu }) {
  /** The status bar's figures in Excel's order: [key, menu label, the bar's label, the value from stats.js]. */
  const STATUS_ITEMS = [['average', 'Average', 'Average', (s) => s.text?.average], ['count', 'Count', 'Count', (s) => String(s.count)],
    ['numericalCount', 'Numerical Count', 'Numerical Count', (s) => (s.text ? String(s.numbers) : null)], ['minimum', 'Minimum', 'Min', (s) => s.text?.min],
    ['maximum', 'Maximum', 'Max', (s) => s.text?.max], ['sum', 'Sum', 'Sum', (s) => s.text?.sum]];
  /** Which of them are on: `statusBar` ({ minimum: true, ... }) over the user's last choice in this browser over
   *  Excel's default three (Average, Count, Sum). */
  const statusOn = (() => {
    let saved = {};
    try { saved = JSON.parse(localStorage.getItem('oxe-status') || '{}') || {}; } catch { /* storage blocked */ }
    return { average: true, count: true, numericalCount: false, minimum: false, maximum: false, sum: true, ...saved, ...(typeof options.statusBar === 'object' ? options.statusBar : {}) };
  })();
  /** The status bar: the figures on (STATUS_ITEMS) for the selection (stats.js); Count alone for text or an error. */
  const showStats = () => {
    if (!state.tabs || !state.editing) return;
    // (Auto Fill Options go once the selection is no longer the cells filled.)
    const sel = state.editing.commands.selection;
    if (state.lastFill && ['r0', 'c0', 'r1', 'c1'].some((k) => sel[k] !== state.lastFill.to[k])) { state.lastFill = null; state.grid?.showFillOptions(null); }
    if (options.statusBar === false) { state.tabs.status('', true); return; }
    const s = selectionStats(state.workbook, state.sheetIndex, state.editing.commands.selection);
    const parts = !s ? [] : STATUS_ITEMS.filter(([k]) => statusOn[k]).map(([, , label, value]) => [label, value(s)]).filter(([, v]) => v != null).map(([l, v]) => `${l}: ${v}`);
    // (Manual calculation with formulas waiting: Excel's Calculate.)
    if (needsCalculation(state.workbook)) parts.unshift('Calculate');
    state.tabs.status(parts.join('    '));
  };
  /** Customize Status Bar (the status bar right-clicked): each figure ticked or not, its value now beside it. */
  const statusMenu = (point) => {
    const s = state.editing && selectionStats(state.workbook, state.sheetIndex, state.editing.commands.selection);
    menu.open(STATUS_ITEMS.map(([k, label, , value]) => ({ label, checked: statusOn[k], value: s ? value(s) ?? '' : '', run: () => {
      statusOn[k] = !statusOn[k];
      try { localStorage.setItem('oxe-status', JSON.stringify(statusOn)); } catch { /* storage blocked */ }
      showStats();
    } })), point, { label: 'Customize Status Bar' });
  };
  return { show: showStats, menu: statusMenu };
}
