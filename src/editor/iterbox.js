// Excel's iterative calculation settings (its File > Options > Formulas: Enable iterative calculation, Maximum
// Iterations, Maximum Change), reached here from Formulas > Calculation Options. calcmode.js setIteration writes them
// as Excel does (measured); Maximum Iterations takes 1 to 32767, Maximum Change any number from 0, as Excel's.
import { modal, button } from './dialogs.js';

const make = (tag, props = {}, ...kids) => { const e = Object.assign(document.createElement(tag), props); e.append(...kids); return e; };

/** The box, started at `current` ({ on, count, delta }, calcmode.js iteration): resolves to the same shape, or null
 *  for Cancel and Esc. A number out of range keeps the box open with a note. */
export function iterationBox(root, current) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-iteration-box');
    const on = make('input', { type: 'checkbox', id: `${id}-on`, checked: !!current.on });
    on.dataset.first = '';
    const field = (key, label, value) => {
      const input = make('input', { id: `${id}-${key}`, className: 'oxe-dialog-field', value: String(value), inputMode: 'decimal' });
      return { input, line: make('div', { className: 'oxe-dialog-row' }, make('label', { className: 'oxe-dialog-label', htmlFor: input.id, textContent: label }), input) };
    };
    const count = field('count', 'Maximum Iterations:', current.count);
    const delta = field('delta', 'Maximum Change:', current.delta);
    // (Excel dims the two numbers while iterative calculation is off.)
    const dim = () => { count.input.disabled = !on.checked; delta.input.disabled = !on.checked; };
    on.addEventListener('change', dim);
    const note = make('p', { className: 'oxe-dialog-note' });
    note.setAttribute('role', 'alert');
    const ok = () => {
      const n = Number(count.input.value.trim());
      const d = Number(delta.input.value.trim());
      if (!Number.isInteger(n) || n < 1 || n > 32767) { note.textContent = 'Maximum Iterations takes a whole number from 1 to 32767.'; count.input.focus(); return; }
      if (delta.input.value.trim() === '' || !(d >= 0)) { note.textContent = 'Maximum Change takes a number from 0 up.'; delta.input.focus(); return; }
      done({ on: on.checked, count: n, delta: d });
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON') { e.preventDefault(); ok(); } });
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Iterative Calculation' }),
      make('div', { className: 'oxe-dialog-option' }, on, make('label', { htmlFor: on.id, textContent: 'Enable iterative calculation' })),
      count.line, delta.line, note,
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))));
    dim();
  });
}
