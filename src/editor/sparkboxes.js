// The sparkline boxes, after Excel's: Create Sparklines (Insert > Sparklines) and Edit Sparklines (the data and
// location ranges), Edit Sparkline Data and Sparkline Date Range (one range), and Hidden and Empty Cell Settings.
// Each resolves to the answer, or null for Cancel and Esc. References are checked by the caller (sparkedit.js).
import { modal, button } from './dialogs.js';
import { make, labelled, check } from './formatbox-kit.js';

/** A box of range fields: `fields` [{ key, label, value }] under `text`; resolves to { key: text }. `check(answer)`
 *  gives a refusal to show (the box stays open) or null. */
export function rangesBox(root, { title, text = '', fields, check: refuse = () => null }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    const inputs = fields.map((f) => make('input', { type: 'text', className: 'oxe-dialog-field', value: f.value || '', spellcheck: false }));
    inputs[fields.findIndex((f) => !f.value) >= 0 ? fields.findIndex((f) => !f.value) : 0].dataset.first = '';
    const why = make('p', { className: 'oxe-dialog-note', role: 'alert' });
    const ok = () => {
      const answer = Object.fromEntries(fields.map((f, k) => [f.key, inputs[k].value.trim()]));
      const no = refuse(answer);
      if (no) { why.textContent = no; return; }
      done(answer);
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName === 'INPUT') { e.preventDefault(); ok(); } });
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: title }),
      ...(text ? [make('p', { className: 'oxe-dialog-text', textContent: text })] : []),
      ...fields.map((f, k) => labelled(f.label, inputs[k])), why,
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))));
  });
}

/** Hidden and Empty Cell Settings: show empty cells as gaps, zero or (a line only) connected; data in hidden rows and
 *  columns shown or not. Resolves to { empty, hidden }. */
export function hiddenEmptyBox(root, { empty, hidden, line }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    const choices = [['gap', 'Gaps'], ['zero', 'Zero'], ['span', 'Connect data points with line']].map(([value, label]) => {
      const input = make('input', { type: 'radio', name: `${id}-empty`, id: `${id}-${value}`, value, checked: value === empty, disabled: value === 'span' && !line });
      return make('div', { className: 'oxe-fc-check' }, input, make('label', { htmlFor: input.id, textContent: label }));
    });
    (choices.find((c) => c.firstChild.checked) || choices[0]).firstChild.dataset.first = '';
    const shown = check('Show data in hidden rows and columns', hidden);
    const ok = () => done({ empty: choices.map((c) => c.firstChild).find((i) => i.checked)?.value || empty, hidden: shown.input.checked });
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName === 'INPUT') { e.preventDefault(); ok(); } });
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Hidden and Empty Cell Settings' }),
      make('fieldset', {}, make('legend', { textContent: 'Show empty cells as:' }), ...choices), shown.el,
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))));
  });
}
