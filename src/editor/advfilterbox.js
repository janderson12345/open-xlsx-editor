// Data > Sort & Filter > Advanced...: Excel's Advanced Filter box (filter the list in place, or copy to another
// location; List range, Criteria range, Copy to; Unique records only). Resolves to { list, criteria, copyTo, unique }
// (ranges of the sheet shown) or null.
import { modal, button } from './dialogs.js';
import { make, check, labelled } from './formatbox-kit.js';
import { parseRange, parseRef, rangeName } from '../ooxml/refs.js';

const read = (t) => parseRange(String(t).replace(/^=/, '').replace(/^.*!/, '').replace(/\$/g, '').trim());

/** The box, its List range filled with `list` (the table round the active cell). */
export function advancedFilterBox(root, { list }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-adv-filter');
    const inPlace = make('input', { type: 'radio', name: `${id}-act`, id: `${id}-in`, checked: true });
    const copy = make('input', { type: 'radio', name: `${id}-act`, id: `${id}-copy` });
    const listField = make('input', { type: 'text', className: 'oxe-dialog-field', value: list ? `$${rangeName(list).replace(/([A-Z]+)(\d+)/g, '$1$$$2').replace(':', ':$')}` : '', spellcheck: false });
    const critField = make('input', { type: 'text', className: 'oxe-dialog-field', spellcheck: false });
    const toField = make('input', { type: 'text', className: 'oxe-dialog-field', spellcheck: false, disabled: true });
    const unique = check('Unique records only', false);
    const sync = () => { toField.disabled = !copy.checked; };
    inPlace.addEventListener('change', sync);
    copy.addEventListener('change', sync);
    const why = make('p', { className: 'oxe-dialog-text', role: 'alert' });
    const ok = () => {
      const l = read(listField.value);
      if (!l || l.r1 <= l.r0) { why.textContent = 'The list range must be a range of cells with headings, like $A$1:$C$9.'; return; }
      const crit = critField.value.trim() ? read(critField.value) : null;
      if (critField.value.trim() && (!crit || crit.r1 <= crit.r0)) { why.textContent = 'The criteria range must be headings with at least one row under them.'; return; }
      let copyTo = null;
      if (copy.checked) {
        const at = parseRef(toField.value.replace(/^=/, '').replace(/^.*!/, '').replace(/\$/g, '').split(':')[0].trim());
        if (!at) { why.textContent = 'Type the cell to copy to, like $I$1.'; return; }
        copyTo = at;
      }
      done({ list: l, criteria: crit, copyTo, unique: unique.input.checked });
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName === 'INPUT') { e.preventDefault(); ok(); } });
    inPlace.dataset.first = '';
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Advanced Filter' }),
      make('fieldset', {}, make('legend', { textContent: 'Action' }),
        make('div', { className: 'oxe-fc-check' }, inPlace, make('label', { htmlFor: inPlace.id, textContent: 'Filter the list, in-place' })),
        make('div', { className: 'oxe-fc-check' }, copy, make('label', { htmlFor: copy.id, textContent: 'Copy to another location' }))),
      labelled('List range:', listField), labelled('Criteria range:', critField), labelled('Copy to:', toField), unique.el, why,
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))));
  });
}
