// Excel's message boxes and its Unhide list, as modal <dialog> elements: focus held inside, Esc cancels, focus back
// where it was on closing. Each resolves a promise with what was chosen.

let dialogCount = 0;

function modal(root, build) {
  return new Promise((resolve) => {
    const dlg = document.createElement('dialog');
    dlg.className = 'oxe-dialog';
    const back = document.activeElement;
    const done = (value) => {
      dlg.close();
      dlg.remove();
      back?.focus?.({ preventScroll: true });
      resolve(value);
    };
    dlg.addEventListener('cancel', (e) => { e.preventDefault(); done(null); });
    build(dlg, done, `oxe-dlg${++dialogCount}`);
    root.append(dlg);
    dlg.showModal();
    dlg.querySelector('[data-first]')?.focus();
  });
}

const button = (label, onClick) => {
  const b = document.createElement('button');
  b.type = 'button';
  b.className = 'oxe-dialog-button';
  b.textContent = label;
  b.addEventListener('click', onClick);
  return b;
};

/** A message with buttons (`buttons`: labels, the first one focused): resolves to the label pressed, or null for Esc. */
export function messageBox(root, text, buttons = ['OK']) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('role', 'alertdialog');
    dlg.setAttribute('aria-labelledby', `${id}-text`);
    const p = document.createElement('p');
    p.id = `${id}-text`;
    p.className = 'oxe-dialog-text';
    p.textContent = text;
    const row = document.createElement('div');
    row.className = 'oxe-dialog-buttons';
    row.append(...buttons.map((label) => button(label, () => done(label))));
    row.firstChild.dataset.first = '';
    dlg.append(p, row);
  });
}

/** Excel's Unhide box: the hidden sheets' names in a list; resolves to the index in `names` chosen, or null. */
export function unhideBox(root, names) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    const title = document.createElement('h2');
    title.id = `${id}-title`;
    title.className = 'oxe-dialog-title';
    title.textContent = 'Unhide';
    const label = document.createElement('label');
    label.className = 'oxe-dialog-text';
    label.htmlFor = `${id}-list`;
    label.textContent = 'Unhide sheet:';
    const list = document.createElement('select');
    list.id = `${id}-list`;
    list.className = 'oxe-dialog-list';
    list.size = Math.max(2, Math.min(8, names.length));
    list.append(...names.map((name, i) => Object.assign(document.createElement('option'), { value: String(i), textContent: name })));
    list.selectedIndex = 0;
    list.dataset.first = '';
    const ok = () => done(list.selectedIndex < 0 ? null : list.selectedIndex);
    list.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); ok(); } });
    list.addEventListener('dblclick', ok);
    const row = document.createElement('div');
    row.className = 'oxe-dialog-buttons';
    row.append(button('OK', ok), button('Cancel', () => done(null)));
    dlg.append(title, label, list, row);
  });
}

/** Excel's Row Height and Column Width boxes: a title, a labelled number field holding `value`, OK and Cancel. Resolves
 *  to the number typed (checked against 0 to `max`: `refusal` said under the field otherwise, and the box kept),
 *  or null for Cancel and Esc. */
export function numberBox(root, { title, label, value, max, refusal }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    const h = document.createElement('h2');
    h.id = `${id}-title`;
    h.className = 'oxe-dialog-title';
    h.textContent = title;
    const lab = document.createElement('label');
    lab.className = 'oxe-dialog-text oxe-dialog-label';
    lab.htmlFor = `${id}-field`;
    lab.textContent = label;
    const field = document.createElement('input');
    field.id = `${id}-field`;
    field.className = 'oxe-dialog-field';
    field.inputMode = 'decimal';
    field.value = value;
    field.dataset.first = '';
    const note = document.createElement('p');
    note.className = 'oxe-dialog-note';
    note.setAttribute('role', 'alert');
    const ok = () => {
      const n = Number(field.value.trim());
      if (field.value.trim() === '' || !Number.isFinite(n) || n < 0 || n > max) {
        note.textContent = refusal;
        field.focus();
        field.select();
        return;
      }
      done(n);
    };
    field.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); ok(); } });
    const row = document.createElement('div');
    row.className = 'oxe-dialog-buttons';
    row.append(button('OK', ok), button('Cancel', () => done(null)));
    dlg.append(h, lab, field, note, row);
    setTimeout(() => field.select());
  });
}
