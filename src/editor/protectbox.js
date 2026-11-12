// Excel's Protect Sheet box (Review > Protect Sheet), its Confirm Password box, and Unprotect Sheet's password box,
// in Excel's words and order (ooxml/protection.js writes what they say).
import { modal, button } from './dialogs.js';

/** The box's list, in Excel's order: [key, label, ticked at first]. */
const LIST = [
  ['selectLocked', 'Select locked cells', true], ['selectUnlocked', 'Select unlocked cells', true], ['formatCells', 'Format cells', false],
  ['formatColumns', 'Format columns', false], ['formatRows', 'Format rows', false], ['insertColumns', 'Insert columns', false],
  ['insertRows', 'Insert rows', false], ['insertHyperlinks', 'Insert hyperlinks', false], ['deleteColumns', 'Delete columns', false],
  ['deleteRows', 'Delete rows', false], ['sort', 'Sort', false], ['autoFilter', 'Use AutoFilter', false],
  ['pivotTables', 'Use PivotTable & PivotChart', false], ['objects', 'Edit objects', false], ['scenarios', 'Edit scenarios', false],
];

const el = (tag, props = {}) => Object.assign(document.createElement(tag), props);
const passwordField = (id) => el('input', { id, className: 'oxe-dialog-field', type: 'password', autocomplete: 'new-password', spellcheck: false });

/** Protect Sheet: resolves to { password, options } (protection.js protectSheet's), or null. */
export function protectBox(root) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    const h = el('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Protect Sheet' });
    const whole = el('input', { type: 'checkbox', id: `${id}-on`, checked: true });
    const wholeLine = el('div', { className: 'oxe-dialog-option' });
    wholeLine.append(whole, el('label', { htmlFor: whole.id, textContent: 'Protect worksheet and contents of locked cells' }));
    const pwLabel = el('label', { className: 'oxe-dialog-text oxe-dialog-label', htmlFor: `${id}-pw`, textContent: 'Password to unprotect sheet:' });
    const pw = passwordField(`${id}-pw`);
    pw.dataset.first = '';
    const set = el('fieldset', { className: 'oxe-dialog-group oxe-protect-list' });
    set.append(el('legend', { textContent: 'Allow all users of this worksheet to:' }));
    const boxes = new Map(LIST.map(([key, label, on], k) => {
      const input = el('input', { type: 'checkbox', id: `${id}-o${k}`, checked: on });
      const line = el('div', { className: 'oxe-dialog-option' });
      line.append(input, el('label', { htmlFor: input.id, textContent: label }));
      set.append(line);
      return [key, input];
    }));
    // (As Excel's box: no unlocked cells to select, none locked either; a locked one selectable, unlocked ones too.)
    boxes.get('selectUnlocked').addEventListener('change', (e) => { if (!e.target.checked) boxes.get('selectLocked').checked = false; });
    boxes.get('selectLocked').addEventListener('change', (e) => { if (e.target.checked) boxes.get('selectUnlocked').checked = true; });
    const okButton = button('OK', () => {
      const on = (k) => boxes.get(k).checked;
      done({
        password: pw.value,
        options: {
          allow: Object.fromEntries(LIST.slice(2, 13).map(([k]) => [k, on(k)])),
          selectLocked: on('selectLocked'),
          selectUnlocked: on('selectUnlocked'),
          objects: !on('objects'),
          scenarios: !on('scenarios'),
        },
      });
    });
    // (Excel's OK can't be pressed with the first box unticked.)
    whole.addEventListener('change', () => { okButton.disabled = !whole.checked; });
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON' && !okButton.disabled) { e.preventDefault(); okButton.click(); } });
    const row = el('div', { className: 'oxe-dialog-buttons' });
    row.append(okButton, button('Cancel', () => done(null)));
    dlg.append(h, wholeLine, pwLabel, pw, set, row);
  });
}

/** A box asking for a password: Confirm Password ("Reenter password to proceed.") or Unprotect Sheet ("Password:").
 *  Resolves to what was typed, or null. */
export function passwordAsk(root, { title, label }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    const h = el('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: title });
    const lab = el('label', { className: 'oxe-dialog-text oxe-dialog-label', htmlFor: `${id}-pw`, textContent: label });
    const field = passwordField(`${id}-pw`);
    field.dataset.first = '';
    const ok = () => done(field.value);
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON') { e.preventDefault(); ok(); } });
    const row = el('div', { className: 'oxe-dialog-buttons' });
    row.append(button('OK', ok), button('Cancel', () => done(null)));
    dlg.append(h, lab, field, row);
  });
}
