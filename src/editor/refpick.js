// A reference field of a box, with Excel's collapse button: pressed, the box (and any box under it) steps aside for a
// small bar holding the field, and the cells selected on the sheet, by mouse or keyboard, go into it as an absolute
// reference with the sheet's name (=Sheet1!$A$1:$C$5), as Excel's reference fields fill. Enter, Esc or the bar's
// button brings the boxes back with the field filled in. Another sheet's tab may be clicked while picking; the box's
// own sheet is shown again afterwards.
import { make } from './formatbox-kit.js';
import { announce } from './live.js';
import { colName } from '../ooxml/refs.js';
import { sheetPrefix } from '../calc/tokens.js';

/** A selection as a series reference: =Sheet1!$B$2:$B$5 ($B$2 for one cell). */
export const pickedText = (sheet, { r0, c0, r1, c1 }) => `=${sheetPrefix(sheet)}$${colName(c0)}$${r0 + 1}${r0 === r1 && c0 === c1 ? '' : `:$${colName(c1)}$${r1 + 1}`}`;

/**
 * A text field with a collapse button: { el, input }. `pick` is { root, grid (grid.js's), selected() (the shown
 * sheet's selection as pickedText gives it), start() (a chart picked let go, for the keys to reach the cells), back()
 * (the box's own sheet shown again) }.
 */
export function refField(label, value, pick) {
  const input = make('input', { type: 'text', className: 'oxe-dialog-field oxe-ref-input', value: value ?? '', spellcheck: false });
  input.setAttribute('aria-label', label);
  const btn = make('button', { type: 'button', className: 'oxe-dialog-button oxe-ref-collapse', textContent: '⤒' });
  btn.setAttribute('aria-label', `${label} Pick the cells on the sheet`);
  btn.title = 'Pick the cells on the sheet';
  btn.addEventListener('click', () => collapse(input, label, pick));
  return { input, el: make('div', { className: 'oxe-ref-field' }, input, btn) };
}

/** The boxes stepped aside while the cells are picked for `input`; resolves when they are back. */
export function collapse(input, label, { root, grid, selected, start, back }) {
  return new Promise((resolve) => {
    start?.();
    const open = [...root.querySelectorAll('dialog.oxe-dialog[open]')];
    for (const d of open) d.close();
    const field = make('input', { type: 'text', className: 'oxe-dialog-field oxe-ref-input', value: input.value, spellcheck: false });
    field.setAttribute('aria-label', label);
    const done = make('button', { type: 'button', className: 'oxe-dialog-button', textContent: '⤓' });
    done.setAttribute('aria-label', 'Back to the box');
    const bar = make('div', { className: 'oxe-refbar' }, make('span', { className: 'oxe-refbar-label', textContent: label }), field, done);
    bar.setAttribute('role', 'group');
    bar.setAttribute('aria-label', label);
    // (After the grid has taken the mouse or key: its selection is the new one then.)
    const fill = (e) => { if (!bar.contains(e?.target)) setTimeout(() => { field.value = selected(); }); };
    const finish = () => {
      window.removeEventListener('mouseup', fill, true);
      window.removeEventListener('keyup', onKeyUp, true);
      window.removeEventListener('keydown', onKey, true);
      bar.remove();
      back?.();
      input.value = field.value;
      for (const d of open) d.showModal();
      input.focus();
      input.dispatchEvent(new Event('change', { bubbles: true }));
      resolve(input.value);
    };
    // (On the sheet: the selection keys move it as ever, and each change fills the field; typed in the bar's field, the
    // reference is the one typed. Enter or Esc, wherever the focus is, ends.)
    const onKey = (e) => { if (e.key === 'Enter' || e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); finish(); } };
    const onKeyUp = (e) => { if (!['Enter', 'Escape', 'Tab'].includes(e.key)) fill(e); };
    done.addEventListener('click', finish);
    window.addEventListener('mouseup', fill, true);
    window.addEventListener('keyup', onKeyUp, true);
    window.addEventListener('keydown', onKey, true);
    root.append(bar);
    grid.el.focus({ preventScroll: true });
    announce(`${label} Select the cells on the sheet, then press Enter.`);
  });
}
