// A chart's data as Excel's Select Data changes it (chartdata.js): each change a named command, one step of the
// history, refused with Excel's words when the reference isn't one; and the Select Data Source box (chartdatabox.js)
// that runs them. Commands take the chart's index (readCharts's) first:
//   chartData(k, range, byRows)       the Chart data range; byRows null keeps the chart's own way, true or false turns
//                                     it (Switch Row/Column)
//   chartSeriesAdd(k, { name, val, cat })        chartSeriesEdit(k, n, { name, val, cat })
//   chartSeriesRemove(k, n)           chartSeriesMove(k, n, -1 or 1)
//   chartBlanks(k, { blanks, hidden })           chartDataBox(k)
import { messageBox } from './dialogs.js';
import { selectDataBox } from './chartdatabox.js';
import { pickedText } from './refpick.js';
import { chartData, setChartData, addChartSeries, editChartSeries, removeChartSeries, moveChartSeries, setChartBlanks, refData } from '../ooxml/chartdata.js';

const NAMES = new Set(['chartData', 'chartSeriesAdd', 'chartSeriesEdit', 'chartSeriesRemove', 'chartSeriesMove', 'chartBlanks', 'chartDataBox']);

/** The commands, for acts-more.js: run(name, args) gives what the command returns, or undefined for another's. */
export function createChartDataActs({ root, state, editing, act }) {
  /** One change of the shown sheet's chart part, as one step of the history (Excel's message when refused). */
  const change = (fn) => {
    let refused = null;
    const done = editing().run('sheetChange', (wb, si) => {
      try { fn(wb, si); } catch (e) { refused = e.message; return false; }
      return true;
    }, 'objects');
    if (refused) return messageBox(root, refused).then(() => false);
    return done;
  };
  /** A series as the box lists it: its name's text, else Excel's Series1, Series2... by its place. */
  const labelOf = (wb, si, s, n) => {
    if (s.name && 'text' in s.name) return s.name.text || `Series${n + 1}`;
    if (s.name?.f) { try { return refData(wb, s.name.f, false, wb.sheets[si].name).values.filter((v) => v != null).join(' ') || `Series${n + 1}`; } catch { /* another book's */ } }
    return `Series${n + 1}`;
  };
  const box = (k) => {
    const chart = (state.charts || []).find((c) => c.index === k);
    const si = state.sheetIndex;
    const read = () => {
      const wb = state.workbook;
      const d = chartData(wb, si, k);
      return { ...d, scatter: /^scatter/.test(d.kind || ''), lines: /^(line|scatter)/.test(d.kind || ''), series: d.series.map((s, n) => ({ ...s, label: labelOf(wb, si, s, n) })) };
    };
    try { read(); } catch (e) { return messageBox(root, e.message).then(() => false); }
    const pick = {
      root, grid: state.grid,
      selected: () => { const s = editing().selection; return pickedText(state.workbook.sheets[state.sheetIndex].name, s); },
      start: () => state.picker?.pick(null, false),
      back: () => { if (state.sheetIndex !== si) editing().run('showSheet', si); },
    };
    return selectDataBox(root, { read, run: (name, ...a) => act(name, k, ...a), undo: () => editing().run('undo'), pick, name: chart?.name || 'Chart' }).then((n) => {
      state.grid.el.focus({ preventScroll: true });
      state.picker?.pick(k, false);
      return n > 0;
    });
  };
  return (name, args) => {
    if (!NAMES.has(name)) return undefined;
    const [k, a, b] = args;
    if (name === 'chartData') return change((wb, si) => setChartData(wb, si, k, a, b ?? null));
    if (name === 'chartSeriesAdd') return change((wb, si) => addChartSeries(wb, si, k, a));
    if (name === 'chartSeriesEdit') return change((wb, si) => editChartSeries(wb, si, k, a, b));
    if (name === 'chartSeriesRemove') return change((wb, si) => removeChartSeries(wb, si, k, a));
    if (name === 'chartSeriesMove') return change((wb, si) => moveChartSeries(wb, si, k, a, b));
    if (name === 'chartBlanks') return change((wb, si) => setChartBlanks(wb, si, k, a));
    return box(k);
  };
}
