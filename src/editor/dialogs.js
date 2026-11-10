// Excel's message boxes, its Unhide list and its boxes of choices (Insert, Delete, Paste Special), as modal <dialog>
// elements: focus held inside, Esc cancels, focus back where it was on closing. Each resolves a promise with what was
// chosen.

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
 *  to the number typed (checked against `min` (0) to `max`: `refusal` said under the field otherwise, and the box kept),
 *  or null for Cancel and Esc. */
export function numberBox(root, { title, label, value, min = 0, max, refusal }) {
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
      if (field.value.trim() === '' || !Number.isFinite(n) || n < min || n > max) {
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

/** A box of choices, as Excel's Insert, Delete and Paste Special boxes are: a title, groups of radio buttons
 *  ({ key, legend, options: labels, value: the one chosen first, columns }), check boxes ({ key, label, value }) and
 *  OK, Cancel and any other `buttons` (before OK, as Excel's Paste Link). Resolves to { button, choices: { key: label },
 *  checks: { key: true or false } }, or null for Cancel and Esc. Enter is OK. */
export function optionsBox(root, { title, groups = [], checks = [], buttons = [] }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    const h = document.createElement('h2');
    h.id = `${id}-title`;
    h.className = 'oxe-dialog-title';
    h.textContent = title;
    dlg.append(h);
    const radios = new Map();
    groups.forEach((g, gi) => {
      const set = document.createElement('fieldset');
      set.className = 'oxe-dialog-group';
      if (g.columns > 1) set.style.setProperty('--oxe-cols', String(g.columns));
      if (g.legend) set.append(Object.assign(document.createElement('legend'), { textContent: g.legend }));
      const name = `${id}-g${gi}`;
      radios.set(g.key, g.options.map((label, oi) => {
        const input = Object.assign(document.createElement('input'), { type: 'radio', name, id: `${name}-${oi}`, value: label, checked: label === g.value });
        if (label === g.value && gi === 0) input.dataset.first = '';
        const lab = Object.assign(document.createElement('label'), { htmlFor: input.id, textContent: label });
        const line = document.createElement('div');
        line.className = 'oxe-dialog-option';
        line.append(input, lab);
        set.append(line);
        return input;
      }));
      dlg.append(set);
    });
    const boxes = new Map();
    if (checks.length) {
      const line = document.createElement('div');
      line.className = 'oxe-dialog-checks';
      checks.forEach((ch, ci) => {
        const input = Object.assign(document.createElement('input'), { type: 'checkbox', id: `${id}-c${ci}`, checked: !!ch.value });
        const lab = Object.assign(document.createElement('label'), { htmlFor: input.id, textContent: ch.label });
        const one = document.createElement('div');
        one.className = 'oxe-dialog-option';
        one.append(input, lab);
        line.append(one);
        boxes.set(ch.key, input);
      });
      dlg.append(line);
    }
    const answer = (button) => done({
      button,
      choices: Object.fromEntries([...radios].map(([k, list]) => [k, list.find((x) => x.checked)?.value ?? null])),
      checks: Object.fromEntries([...boxes].map(([k, input]) => [k, input.checked])),
    });
    dlg.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && e.target.tagName !== 'BUTTON') { e.preventDefault(); answer('OK'); }
    });
    const row = document.createElement('div');
    row.className = 'oxe-dialog-buttons';
    row.append(...buttons.map((label) => button(label, () => answer(label))), button('OK', () => answer('OK')), button('Cancel', () => done(null)));
    dlg.append(row);
    if (!dlg.querySelector('[data-first]')) (dlg.querySelector('input') || row.firstChild).dataset.first = '';
  });
}
