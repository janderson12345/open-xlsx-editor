// The table boxes, after Excel's: Create Table (Insert > Table, Ctrl+T, Ctrl+L and Format as Table: where the data is,
// and My table has headers) and Resize Table (the new range). Each resolves to the answer, or null for Cancel and Esc;
// the range typed is checked by `check` (acts-tables.js), the box staying open with its refusal.
import { modal, button } from './dialogs.js';
import { make, labelled, check } from './formatbox-kit.js';

/** Create Table: resolves to { range, headers }. */
export function createTableBox(root, { range, headers, check: refuse = () => null, title = 'Create Table' }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    const input = make('input', { type: 'text', className: 'oxe-dialog-field', value: range, spellcheck: false });
    input.dataset.first = '';
    const has = check('My table has headers', headers);
    const why = make('p', { className: 'oxe-dialog-note', role: 'alert' });
    const ok = () => {
      const answer = { range: input.value.trim(), headers: has.input.checked };
      const no = refuse(answer);
      if (no) { why.textContent = no; return; }
      done(answer);
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName === 'INPUT' && e.target.type === 'text') { e.preventDefault(); ok(); } });
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: title }),
      labelled('Where is the data for your table?', input), has.el, why,
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))));
  });
}

/** Resize Table: resolves to the range typed. */
export function resizeTableBox(root, { range, check: refuse = () => null }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    const input = make('input', { type: 'text', className: 'oxe-dialog-field', value: range, spellcheck: false });
    input.dataset.first = '';
    const why = make('p', { className: 'oxe-dialog-note', role: 'alert' });
    const ok = () => {
      const no = refuse(input.value.trim());
      if (no) { why.textContent = no; return; }
      done(input.value.trim());
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName === 'INPUT') { e.preventDefault(); ok(); } });
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Resize Table' }),
      labelled('Select the new data range for your table:', input),
      make('p', { className: 'oxe-dialog-text', textContent: 'Note: The headers must remain in the same row, and the resulting table range must overlap the original table range.' }), why,
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))));
  });
}
