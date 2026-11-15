// The Format Chart box (a picked chart's menu, Format Chart..., or Ctrl+1): one part of the chart at a time, picked
// from Excel's names for them, and what can be set on it here: the chart and plot areas' and a series' (or one of
// its points') fill and line; the font of the title, the axes, their titles, the legend and the data labels; an axis's
// title; a series' chart type and axis (Change Series Chart Type), on charts of columns, lines and areas. Resolves to
// the change for chartacts.js chartFormat ({ target, look, font, title, type, secondary }, only what was changed), or null.
import { modal, button } from './dialogs.js';
import { make, check, select, labelled } from './formatbox-kit.js';
import { STANDARD_COLORS } from './sheetmenu.js';
import { SERIES_TYPES } from '../ooxml/chartcombo.js';

const TYPE_NAMES = { column: 'Clustered Column', line: 'Line', lineMarkers: 'Line with Markers', area: 'Area' };

/** The parts of `chart` this box formats: [[key, name]]. */
export function chartElements(chart) {
  const pie = chart.kind === 'pie' || chart.kind === 'doughnut';
  const horizontal = chart.kind === 'bar' && chart.dir === 'bar';
  const out = [['chart', 'Chart Area'], ['plot', 'Plot Area']];
  if (chart.title) out.push(['title', 'Chart Title']);
  if (!pie) {
    out.push(['h', horizontal ? 'Horizontal (Value) Axis' : 'Horizontal (Category) Axis'], ['v', horizontal ? 'Vertical (Category) Axis' : 'Vertical (Value) Axis']);
    if (chart.axis2) out.push(['v2', 'Secondary Vertical (Value) Axis']);
    if (horizontal ? chart.axis.title : chart.catTitle) out.push(['hTitle', 'Horizontal Axis Title']);
    if (horizontal ? chart.catTitle : chart.axis.title) out.push(['vTitle', 'Vertical Axis Title']);
  }
  if (chart.legend) out.push(['legend', 'Legend']);
  if (chart.labels) out.push(['labels', 'Data Labels']);
  for (const s of chart.series) out.push([`s${s.index}`, `Series "${s.name}"`]);
  return out;
}

/** A colour field: As it is, Automatic, None, Excel's standard colours, or one picked. value(): undefined (as it
 *  is), 'auto', 'none' or 'RRGGBB'. */
function colourField(noneLabel, auto = true) {
  const sel = select([['', 'As it is'], ...(auto ? [['auto', 'Automatic']] : []), ...(noneLabel ? [['none', noneLabel]] : []), ...STANDARD_COLORS.map(([n, hex]) => [hex, n]), ['more', 'More Colors...']], '');
  const custom = make('input', { type: 'color', value: '#000000', hidden: true, title: 'Color' });
  sel.addEventListener('change', () => { custom.hidden = sel.value !== 'more'; });
  const value = () => (sel.value === '' ? undefined : sel.value === 'more' ? custom.value.slice(1).toUpperCase() : sel.value);
  return { el: make('span', { className: 'oxe-fc-color' }, sel, custom), sel, custom, value };
}

/** The box for `chart` (charts.js readCharts's); `fontNames` the names offered. */
export function chartFormatBox(root, chart, { fontNames = [] } = {}) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-chart-format-box');
    const parts = chartElements(chart);
    const part = select(parts, parts[0][0]);
    // Fill and line.
    const fill = colourField('No fill');
    const line = colourField('No line');
    const point = select([['', 'All points'], ...chart.cats.map((c, i) => [String(i), `Point ${i + 1}${c != null ? ` (${c})` : ''}`])], '');
    const looks = make('fieldset', { className: 'oxe-chart-format-look' }, make('legend', { textContent: 'Fill and line' }), labelled('Fill:', fill.sel), fill.custom, labelled('Line:', line.sel), line.custom, labelled('Points:', point));
    // Font.
    const name = select([['', 'As it is'], ...fontNames.map((n) => [n, n])], '');
    const size = make('input', { type: 'number', min: 1, max: 409, step: 0.5, className: 'oxe-dialog-field', placeholder: 'As it is' });
    const bold = select([['', 'As it is'], ['1', 'Bold'], ['0', 'Regular']], '');
    const colour = colourField(null, false);
    const fonts = make('fieldset', { className: 'oxe-chart-format-font' }, make('legend', { textContent: 'Font' }), labelled('Font:', name), labelled('Size:', size), labelled('Style:', bold), labelled('Color:', colour.sel), colour.custom);
    // An axis's title.
    const titleMode = select([['', 'As it is'], ['none', 'None'], ['auto', 'Axis Title'], ['text', 'This text:']], '');
    const titleText = make('input', { type: 'text', className: 'oxe-dialog-field', spellcheck: true });
    titleText.setAttribute('aria-label', 'Axis title');
    const titles = make('fieldset', { className: 'oxe-chart-format-title' }, make('legend', { textContent: 'Axis title' }), labelled('Title:', titleMode), titleText);
    // A series' chart type and axis.
    const type = select([['', 'As it is'], ...Object.keys(SERIES_TYPES).map((t) => [t, TYPE_NAMES[t]])], '');
    const secondary = check('Secondary axis', false);
    const combo = (chart.groups || []).every((g) => ['bar', 'line', 'area'].includes(g.kind) && !(g.kind === 'bar' && g.dir === 'bar'));
    const series = make('fieldset', { className: 'oxe-chart-format-series' }, make('legend', { textContent: 'Series' }), labelled('Chart type:', type), secondary.el);
    const sync = () => {
      const k = part.value;
      const s = k[0] === 's' && chart.series.find((x) => `s${x.index}` === k);
      looks.hidden = !(k === 'chart' || k === 'plot' || s);
      point.parentElement.hidden = !s;
      fonts.hidden = !['title', 'h', 'v', 'v2', 'hTitle', 'vTitle', 'legend', 'labels'].includes(k);
      titles.hidden = !['h', 'v', 'v2'].includes(k);
      series.hidden = !s || !combo;
      if (s) secondary.input.checked = !!s.secondary;
      titleText.disabled = titleMode.value !== 'text';
    };
    part.addEventListener('change', sync);
    titleMode.addEventListener('change', sync);
    sync();
    const ok = () => {
      const k = part.value;
      const s = k[0] === 's' && chart.series.find((x) => `s${x.index}` === k);
      const out = { target: s ? { series: s.index, ...(point.value !== '' ? { point: Number(point.value) } : {}) } : k };
      if (!looks.hidden) {
        const look = {};
        if (fill.value() !== undefined) look.fill = fill.value();
        if (line.value() !== undefined) look.line = line.value();
        if (Object.keys(look).length) out.look = look;
      }
      if (!fonts.hidden) {
        const font = {};
        if (name.value) font.name = name.value;
        if (size.value !== '' && Number(size.value) > 0) font.size = Number(size.value);
        if (bold.value) font.bold = bold.value === '1';
        if (colour.value()) font.color = colour.value();
        if (Object.keys(font).length) out.font = font;
      }
      if (!titles.hidden && titleMode.value) out.title = titleMode.value === 'none' ? null : titleMode.value === 'auto' ? 'auto' : titleText.value;
      if (!series.hidden) {
        if (type.value) out.type = type.value;
        if (secondary.input.checked !== !!s.secondary) out.secondary = secondary.input.checked;
      }
      done(Object.keys(out).length > 1 ? out : null);
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName === 'INPUT' && e.target.type !== 'checkbox') { e.preventDefault(); ok(); } });
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: `Format Chart: ${chart.name}` }),
      labelled('Format:', part), looks, fonts, titles, series,
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))));
  });
}
