// The Edit Chart box: a picked chart's kind (Excel's Change Chart Type, by the Insert tab's names), its title (none,
// Excel's own, or text), its legend's place (or none) and its data labels. Resolves to the change (only what was
// changed) for acts.js chartEdit, { selectData: true } for its Select Data... button (chartdatabox.js), or null.
import { modal, button } from './dialogs.js';
import { make, check, select, labelled } from './formatbox-kit.js';
import { CHART_MENUS } from './inserttab.js';

/** The Insert tab's name for a chart as charts.js reads it, or null for a kind it doesn't offer. */
export function kindOf(chart) {
  const g = chart.grouping;
  if (chart.kind === 'bar') return `${chart.dir === 'bar' ? 'bar' : 'column'}${g === 'stacked' ? 'Stacked' : g === 'percentStacked' ? 'Stacked100' : 'Clustered'}`;
  if (chart.kind === 'line') return chart.series.some((s) => s.marker && s.marker !== 'none') ? 'lineMarkers' : 'line';
  if (chart.kind === 'area') return g === 'stacked' ? 'areaStacked' : 'area';
  if (chart.kind === 'scatter') return chart.series.some((s) => s.smooth) ? 'scatterSmooth' : chart.series.some((s) => s.line) ? 'scatterLines' : 'scatter';
  return chart.kind === 'pie' || chart.kind === 'doughnut' ? chart.kind : null;
}

/** The box for `chart` (charts.js readCharts's). */
export function chartBox(root, chart) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-chart-box');
    const now = kindOf(chart);
    const kinds = CHART_MENUS.flatMap(([, , , list]) => list);
    const kind = select([...(now ? [] : [['', 'As it is']]), ...kinds], now || '');
    kind.dataset.first = '';
    const titleMode = select([['none', 'None'], ['auto', 'Excel\'s own (the series\' name, or Chart Title)'], ['text', 'This text:']], chart.title == null ? 'none' : chart.ownTitle ? 'text' : 'auto');
    const titleText = make('input', { type: 'text', className: 'oxe-dialog-field', value: chart.ownTitle ? chart.title : '', spellcheck: true });
    titleText.setAttribute('aria-label', 'Chart title');
    const legend = select([['none', 'None'], ['b', 'Bottom'], ['t', 'Top'], ['l', 'Left'], ['r', 'Right']], chart.legend || 'none');
    const labels = check('Data labels (values)', chart.labels);
    const sync = () => { titleText.disabled = titleMode.value !== 'text'; };
    titleMode.addEventListener('change', sync);
    sync();
    const ok = () => {
      const out = {};
      if (kind.value && kind.value !== now) out.kind = kind.value;
      const title = titleMode.value === 'none' ? null : titleMode.value === 'auto' ? 'auto' : titleText.value;
      const was = chart.title == null ? null : chart.ownTitle ? chart.title : 'auto';
      if (title !== was) out.title = title;
      const pos = legend.value === 'none' ? null : legend.value;
      if (pos !== (chart.legend || null)) out.legend = pos;
      if (labels.input.checked !== !!chart.labels) out.labels = labels.input.checked;
      done(Object.keys(out).length ? out : null);
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName === 'INPUT' && e.target.type === 'text') { e.preventDefault(); ok(); } });
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: `Edit Chart: ${chart.name}` }),
      labelled('Chart type:', kind), labelled('Chart title:', titleMode), titleText, labelled('Legend:', legend), labels.el,
      make('div', { className: 'oxe-dialog-buttons' }, button('Select Data...', () => done({ selectData: true })), button('OK', ok), button('Cancel', () => done(null))));
  });
}
