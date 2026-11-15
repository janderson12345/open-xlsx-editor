// A chart's formatting as named commands (acts-more.js hands them here), each one step of Undo: chartFormat (the
// Format Chart box's answer: { target, look, font, title, type, secondary }), and the single changes it is made of,
// chartAxisTitle, chartFill, chartFont, chartSeriesType and chartSeriesAxis; chartFormatBox opens the box. Each
// change is written as Excel writes it (ooxml/chartformat.js, chartcombo.js), refused with a message for a chart that
// can't take it here.
import { messageBox } from './dialogs.js';
import { chartFormatBox } from './chartformatbox.js';
import { setAxisTitle, setChartFill, setChartFont } from '../ooxml/chartformat.js';
import { setSeriesType, setSeriesAxis } from '../ooxml/chartcombo.js';

/** The change's edits on chart k of sheet si, in the order Excel's pane would make them. */
function apply(wb, si, k, change) {
  const t = change.target;
  if ('title' in change) setAxisTitle(wb, si, k, t, change.title);
  if (change.look) setChartFill(wb, si, k, t, change.look);
  if (change.font) setChartFont(wb, si, k, t, change.font);
  if (change.type) setSeriesType(wb, si, k, t.series, change.type);
  if ('secondary' in change) setSeriesAxis(wb, si, k, t.series, change.secondary);
}

/** The commands: `env` is acts.js's. Returns run(name, args): what the command returns, or undefined for one not here. */
export function createChartActs(env, fontNames) {
  const { root, state, editing, act } = env;
  return (name, args) => {
    if (name === 'chartFormat') {
      const [k, change] = args;
      let refused = null;
      const done = editing().run('sheetChange', (wb, si) => {
        try { apply(wb, si, k, change); } catch (e) { refused = e.message; return false; }
        return true;
      }, 'objects');
      if (refused) return messageBox(root, refused).then(() => false);
      state.picker?.pick(k, false);
      return done;
    }
    if (name === 'chartAxisTitle') return act('chartFormat', args[0], { target: args[1], title: args[2] });
    if (name === 'chartFill') return act('chartFormat', args[0], { target: args[1], look: args[2] });
    if (name === 'chartFont') return act('chartFormat', args[0], { target: args[1], font: args[2] });
    if (name === 'chartSeriesType') return act('chartFormat', args[0], { target: { series: args[1] }, type: args[2] });
    if (name === 'chartSeriesAxis') return act('chartFormat', args[0], { target: { series: args[1] }, secondary: !!args[2] });
    if (name === 'chartFormatBox') {
      const chart = (state.charts || []).find((c) => c.index === args[0]);
      if (!chart) return false;
      const theme = state.workbook.theme;
      return chartFormatBox(root, chart, { fontNames: [...new Set([theme.fonts.minor, theme.fonts.major, ...fontNames])] }).then((change) => {
        state.grid.el.focus({ preventScroll: true });
        return change ? act('chartFormat', args[0], change) : false;
      });
    }
    return undefined;
  };
}
