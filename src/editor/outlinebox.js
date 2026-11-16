// The outline's boxes, as Excel lays them out: Group and Ungroup (Rows or Columns, for a selection that isn't whole
// lines), Settings (Data > Outline's launcher) and Subtotal. Each resolves to its choices or null; the work is
// outline.js's and subtotal.js's.
import { modal, button, optionsBox } from './dialogs.js';
import { make, check, select, labelled } from './formatbox-kit.js';

/** Group or Ungroup: Rows (first, as Excel's) or Columns; resolves to 'row', 'col' or null. */
export function groupAxisBox(root, title) {
  return optionsBox(root, { title, groups: [{ key: 'axis', options: ['Rows', 'Columns'], value: 'Rows' }] })
    .then((got) => (got ? (got.choices.axis === 'Columns' ? 'col' : 'row') : null));
}

/** Settings: Summary rows below detail, Summary columns to right of detail, Automatic styles, and Create and Apply
 *  Styles beside OK, as Excel's box has them; resolves to { below, right, styles, then } (then null, 'create' or
 *  'apply') or null. */
export function outlineSettingsBox(root, { below, right, styles }) {
  const checks = [{ key: 'below', label: 'Summary rows below detail', value: below }, { key: 'right', label: 'Summary columns to right of detail', value: right }, { key: 'styles', label: 'Automatic styles', value: styles }];
  return optionsBox(root, { title: 'Settings', text: 'Direction', checks, buttons: ['Create', 'Apply Styles'] })
    .then((got) => (got ? { below: got.checks.below, right: got.checks.right, styles: got.checks.styles, then: got.button === 'Create' ? 'create' : got.button === 'Apply Styles' ? 'apply' : null } : null));
}

/** Excel's Use function list, in its order. */
const FUNCTIONS = [['sum', 'Sum'], ['count', 'Count'], ['average', 'Average'], ['max', 'Max'], ['min', 'Min'], ['product', 'Product'],
  ['countNums', 'Count Numbers'], ['stdDev', 'StdDev'], ['stdDevp', 'StdDevp'], ['var', 'Var'], ['varp', 'Varp']];

/** Subtotal: At each change in (the list's headings), Use function, Add subtotal to (the last column ticked, as Excel
 *  ticks it the first time), Replace current subtotals, Page break between groups, Summary below data, and Remove All.
 *  `columns` [{ c, label }]. Resolves to { by, fn, cols, replace, pageBreaks, below }, { remove: true } or null. */
export function subtotalBox(root, { columns, below = true }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-subtotal-box');
    const by = select(columns.map((x) => [String(x.c), x.label]), String(columns[0].c));
    const fn = select(FUNCTIONS, 'sum');
    by.dataset.first = '';
    const boxes = columns.map((x, k) => check(x.label, k === columns.length - 1));
    const list = make('div', { className: 'oxe-filter-list', role: 'group', 'aria-label': 'Add subtotal to' }, ...boxes.map((b) => b.el));
    const replace = check('Replace current subtotals', true);
    const breaks = check('Page break between groups', false);
    const summary = check('Summary below data', below);
    const why = make('p', { className: 'oxe-dialog-text', role: 'alert' });
    const ok = () => {
      const cols = columns.filter((_, k) => boxes[k].input.checked).map((x) => x.c);
      if (!cols.length) { why.textContent = 'Choose a column to add subtotals to.'; return; }
      done({ by: Number(by.value), fn: fn.value, cols, replace: replace.input.checked, pageBreaks: breaks.input.checked, below: summary.input.checked });
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON') { e.preventDefault(); ok(); } });
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Subtotal' }),
      labelled('At each change in:', by), labelled('Use function:', fn), make('p', { className: 'oxe-dialog-text', textContent: 'Add subtotal to:' }), list,
      replace.el, breaks.el, summary.el, why,
      make('div', { className: 'oxe-dialog-buttons' }, button('Remove All', () => done({ remove: true })), button('OK', ok), button('Cancel', () => done(null))));
  });
}
