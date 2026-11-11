// The data tools' boxes, as Excel lays them out: Series (Home > Fill > Series...), Remove Duplicates and Text to
// Columns (Data > Data Tools), and Move or Copy (a sheet tab's menu). Each resolves to its choices or null; the tools
// themselves are datatools.js's and sheetcopy.js's.
import { modal, button } from './dialogs.js';
import { make, check, select, labelled } from './formatbox-kit.js';
import { parseRef } from '../ooxml/refs.js';

const radios = (id, name, pairs, value) => {
  const inputs = {};
  const el = make('div', { className: 'oxe-fc-radios' }, ...pairs.map(([v, label]) => {
    const input = make('input', { type: 'radio', name: `${id}-${name}`, id: `${id}-${name}-${v}`, checked: v === value });
    inputs[v] = input;
    return make('div', { className: 'oxe-fc-check' }, input, make('label', { htmlFor: input.id, textContent: label }));
  }));
  return { el, value: () => Object.entries(inputs).find(([, x]) => x.checked)?.[0] ?? value, inputs };
};
const shell = (root, title, cls, build) => modal(root, (dlg, done, id) => {
  dlg.setAttribute('aria-labelledby', `${id}-title`);
  dlg.classList.add(cls);
  const why = make('p', { className: 'oxe-dialog-text', role: 'alert' });
  const { body, ok } = build(id, done, (t) => { why.textContent = t; });
  dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName === 'INPUT' && e.target.type !== 'checkbox') { e.preventDefault(); ok(); } });
  dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: title }), ...body, why,
    make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))));
  dlg.querySelector('input, select')?.setAttribute('data-first', '');
});

/** Fill > Series: `rows` the selection's rows across (more columns than rows) or not. Resolves to fillSeries's options. */
export function seriesBox(root, { across }) {
  return shell(root, 'Series', 'oxe-series-box', (id, done, say) => {
    const by = radios(id, 'by', [['rows', 'Rows'], ['columns', 'Columns']], across ? 'rows' : 'columns');
    const type = radios(id, 'type', [['linear', 'Linear'], ['growth', 'Growth'], ['date', 'Date']], 'linear');
    const unit = radios(id, 'unit', [['day', 'Day'], ['weekday', 'Weekday'], ['month', 'Month'], ['year', 'Year']], 'day');
    const trend = check('Trend', false);
    const step = make('input', { type: 'text', className: 'oxe-dialog-field', value: '1' });
    const stop = make('input', { type: 'text', className: 'oxe-dialog-field' });
    const sync = () => { for (const x of Object.values(unit.inputs)) x.disabled = type.value() !== 'date'; };
    for (const x of Object.values(type.inputs)) x.addEventListener('change', sync);
    sync();
    const ok = () => {
      const s = Number(step.value);
      const t = stop.value.trim() === '' ? null : Number(stop.value);
      if (!Number.isFinite(s) || (t != null && !Number.isFinite(t))) { say('Type numbers for the step and stop values.'); return; }
      done({ by: by.value(), type: type.value(), unit: unit.value(), step: s, stop: t, trend: trend.input.checked });
    };
    return { body: [make('div', { className: 'oxe-fc-grid' }, make('fieldset', {}, make('legend', { textContent: 'Series in' }), by.el),
      make('fieldset', {}, make('legend', { textContent: 'Type' }), type.el), make('fieldset', {}, make('legend', { textContent: 'Date unit' }), unit.el)), trend.el,
    labelled('Step value:', step), labelled('Stop value:', stop)], ok };
  });
}

/** Remove Duplicates: the range's columns (their headings or letters) to tick, and My data has headers. */
export function removeDuplicatesBox(root, { columns, headers }) {
  return shell(root, 'Remove Duplicates', 'oxe-dups-box', (id, done, say) => {
    const has = check('My data has headers', headers);
    const boxes = columns.map((c) => check(c.label, true));
    const list = make('div', { className: 'oxe-filter-list', role: 'group', 'aria-label': 'Columns' }, ...boxes.map((b) => b.el));
    const ok = () => {
      const cols = columns.filter((_, k) => boxes[k].input.checked).map((c) => c.c);
      if (!cols.length) { say('Tick at least one column.'); return; }
      done({ cols, header: has.input.checked });
    };
    return { body: [make('p', { className: 'oxe-dialog-text', textContent: 'To delete duplicate values, select one or more columns that contain duplicates.' }),
      button('Select All', () => boxes.forEach((b) => { b.input.checked = true; })), button('Unselect All', () => boxes.forEach((b) => { b.input.checked = false; })), has.el, list], ok };
  });
}

/** Text to Columns (Delimited): the delimiters, Treat consecutive delimiters as one, the text qualifier, Destination. */
export function textToColumnsBox(root, { dest }) {
  return shell(root, 'Convert Text to Columns', 'oxe-ttc-box', (id, done, say) => {
    const tab = check('Tab', true);
    const semicolon = check('Semicolon', false);
    const comma = check('Comma', false);
    const space = check('Space', false);
    const other = make('input', { type: 'text', className: 'oxe-fc-num', maxLength: 1 });
    other.setAttribute('aria-label', 'Other delimiter');
    const consecutive = check('Treat consecutive delimiters as one', false);
    const quote = select([['"', '"'], ["'", "'"], ['', '{none}']], '"');
    const to = make('input', { type: 'text', className: 'oxe-dialog-field', value: dest, spellcheck: false });
    const ok = () => {
      const at = parseRef(to.value.replace(/^=/, '').replace(/^.*!/, '').replace(/\$/g, '').trim());
      if (!at) { say('Type the destination cell, like $A$1.'); return; }
      done({ delimiters: { tab: tab.input.checked, semicolon: semicolon.input.checked, comma: comma.input.checked, space: space.input.checked, other: other.value }, consecutive: consecutive.input.checked, quote: quote.value, dest: at });
    };
    return { body: [make('fieldset', {}, make('legend', { textContent: 'Delimiters' }), tab.el, semicolon.el, comma.el, space.el, labelled('Other:', other)), consecutive.el,
      labelled('Text qualifier:', quote), labelled('Destination:', to)], ok };
  });
}

/** Move or Copy: the sheet to put it before (or the end), and Create a copy. Resolves to { before (an index, or the
 *  count for the end), copy }. */
export function moveCopyBox(root, { name, sheets, at }) {
  return shell(root, 'Move or Copy', 'oxe-movecopy-box', (id, done) => {
    const list = make('select', { className: 'oxe-dialog-list', size: Math.min(8, sheets.length + 1) },
      ...sheets.map((s, k) => make('option', { value: String(k), textContent: s, selected: k === at })), make('option', { value: String(sheets.length), textContent: '(move to end)' }));
    const copy = check('Create a copy', false);
    const ok = () => done({ before: Number(list.value), copy: copy.input.checked });
    list.addEventListener('dblclick', ok);
    return { body: [make('p', { className: 'oxe-dialog-text', textContent: `Move selected sheets: ${name}` }), labelled('Before sheet:', list), copy.el], ok };
  });
}
