// Excel's Select Data Source box for a chart (its menu's Select Data..., or Edit Chart's Select Data... button): the
// Chart data range, Switch Row/Column, the series (Legend Entries) with Add, Edit, Remove, Move Up and Move Down, and
// Hidden and Empty Cells. As in Excel each change is made at once, the chart redrawn behind the box, and Cancel takes
// back every change made since it opened. Each change is a named command (acts-chartdata.js), one step of Undo.
import { modal, button } from './dialogs.js';
import { make, check, select, labelled } from './formatbox-kit.js';
import { refField } from './refpick.js';

/** Excel's Edit Series box: the series' name (text, or a reference), values and categories (X values for a scatter
 *  chart), each a reference field. Resolves to { name, val, cat } or null. */
export function seriesBox(root, { title, series = {}, scatter, pick }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-series-box');
    const name = refField('Series name:', series.name == null ? '' : 'text' in series.name ? series.name.text : `=${series.name.f}`, pick);
    const val = refField(scatter ? 'Series Y values:' : 'Series values:', series.val ? `=${series.val}` : '', pick);
    const cat = refField(scatter ? 'Series X values:' : 'Category (axis) labels:', series.cat ? `=${series.cat}` : '', pick);
    name.input.dataset.first = '';
    const ok = () => done({ name: name.input.value, val: val.input.value, cat: cat.input.value });
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName === 'INPUT') { e.preventDefault(); ok(); } });
    const field = (text, f, key) => make('div', { className: 'oxe-fc-field' }, make('label', { htmlFor: (f.input.id = `${id}-${key}`), textContent: text }), f.el);
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: title }),
      field('Series name:', name, 'name'), field(scatter ? 'Series Y values:' : 'Series values:', val, 'val'),
      field(scatter ? 'Series X values:' : 'Category (axis) labels:', cat, 'cat'),
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))));
  });
}

/** Excel's Hidden and Empty Cell Settings: empty cells as gaps, zero, or (lines and scatter charts) the line carried
 *  over them; and data in hidden rows and columns drawn or not. Resolves to { blanks, hidden } or null. */
export function blanksBox(root, { blanks, hidden, lines }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    const how = select([['gap', 'Gaps'], ['zero', 'Zero'], ...(lines ? [['span', 'Connect data points with line']] : [])], blanks === 'span' && !lines ? 'gap' : blanks);
    how.dataset.first = '';
    const shown = check('Show data in hidden rows and columns', hidden);
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Hidden and Empty Cell Settings' }),
      labelled('Show empty cells as:', how), shown.el,
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', () => done({ blanks: how.value, hidden: shown.input.checked })), button('Cancel', () => done(null))));
  });
}

/**
 * The Select Data Source box. `env`: { read() (the chart's data: chartdata.js chartData, each series with its
 * `label`), run(name, ...args) (a command: true when made, false or a promise of false when refused), undo() (one
 * step back), pick (refpick.js's), name (the chart's) }. Resolves to the number of changes kept.
 */
export function selectDataBox(root, env) {
  let made = 0;
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-select-data');
    const range = refField('Chart data range:', '', env.pick);
    range.input.dataset.first = '';
    const note = make('p', { className: 'oxe-dialog-text oxe-select-data-note' });
    const list = make('select', { className: 'oxe-dialog-select oxe-series-list', size: 6 });
    list.setAttribute('aria-label', 'Legend Entries (Series)');
    const btns = {};
    let now = null;
    const fill = (at = list.selectedIndex) => {
      now = env.read();
      range.input.value = now.range || '';
      note.textContent = now.range ? '' : 'The data range is too complex to be displayed. Changing it replaces all of the series.';
      list.replaceChildren(...now.series.map((s, k) => make('option', { value: String(k), textContent: s.label })));
      list.selectedIndex = Math.min(Math.max(at, 0), now.series.length - 1);
      sync();
    };
    const sync = () => {
      const k = list.selectedIndex;
      btns.edit.disabled = k < 0;
      btns.remove.disabled = k < 0 || now.series.length < 2;
      btns.up.disabled = k <= 0;
      btns.down.disabled = k < 0 || k >= now.series.length - 1;
      btns.switch.disabled = !now.range;
    };
    /** A command run; the box shows the chart's data again after it. */
    const step = async (at, name, ...args) => {
      const ok = await env.run(name, ...args);
      if (ok) made++;
      fill(at);
      return ok;
    };
    range.input.addEventListener('change', () => {
      const text = range.input.value.trim();
      if (text && text !== (now.range || '')) step(list.selectedIndex, 'chartData', text, null);
    });
    btns.switch = button('Switch Row/Column', () => step(0, 'chartData', now.range, !now.byRows));
    btns.add = button('Add', async () => {
      const got = await seriesBox(root, { title: 'Edit Series', scatter: now.scatter, pick: env.pick });
      if (got) await step(now.series.length, 'chartSeriesAdd', { name: got.name.trim() === '' ? null : got.name, val: got.val, cat: got.cat });
      list.focus();
    });
    btns.edit = button('Edit', async () => {
      const k = list.selectedIndex;
      const got = await seriesBox(root, { title: 'Edit Series', series: now.series[k], scatter: now.scatter, pick: env.pick });
      if (got) await step(k, 'chartSeriesEdit', k, got);
      list.focus();
    });
    btns.remove = button('Remove', () => step(list.selectedIndex, 'chartSeriesRemove', list.selectedIndex));
    btns.up = button('Move Up', () => step(list.selectedIndex - 1, 'chartSeriesMove', list.selectedIndex, -1));
    btns.down = button('Move Down', () => step(list.selectedIndex + 1, 'chartSeriesMove', list.selectedIndex, 1));
    btns.blanks = button('Hidden and Empty Cells', async () => {
      const got = await blanksBox(root, { blanks: now.blanks, hidden: now.hidden, lines: now.lines });
      if (got && (got.blanks !== now.blanks || got.hidden !== now.hidden)) await step(list.selectedIndex, 'chartBlanks', got);
    });
    btns.up.setAttribute('aria-label', 'Move Up');
    btns.down.setAttribute('aria-label', 'Move Down');
    list.addEventListener('change', sync);
    list.addEventListener('dblclick', () => btns.edit.click());
    const cancel = () => { for (; made > 0; made--) env.undo(); done(0); };
    dlg.addEventListener('cancel', (e) => { e.preventDefault(); e.stopImmediatePropagation(); cancel(); }, true);
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: `Select Data Source: ${env.name}` }),
      make('div', { className: 'oxe-fc-field' }, make('label', { htmlFor: range.input.id = `${id}-range`, textContent: 'Chart data range:' }), range.el), note,
      make('div', { className: 'oxe-dialog-buttons oxe-select-data-switch' }, btns.switch),
      make('h3', { className: 'oxe-dialog-subtitle', textContent: 'Legend Entries (Series)' }),
      make('div', { className: 'oxe-dialog-buttons' }, btns.add, btns.edit, btns.remove, btns.up, btns.down), list,
      make('div', { className: 'oxe-dialog-buttons' }, btns.blanks, button('OK', () => done(made)), button('Cancel', cancel)));
    fill(0);
  });
}
