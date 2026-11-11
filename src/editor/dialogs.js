// Excel's message boxes, its Unhide list and its boxes of choices (Insert, Delete, Paste Special), as modal <dialog>
// elements: focus held inside, Esc cancels, focus back where it was on closing. Each resolves a promise with what was
// chosen.

let dialogCount = 0;

export function modal(root, build) {
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

export const button = (label, onClick) => {
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
 *  OK (`ok`: its label, as Excel's Sort Warning has Sort), Cancel and any other `buttons` (before OK, as Excel's Paste
 *  Link), under `text` when given. Resolves to { button, choices: { key: label }, checks: { key: true or false } }, or
 *  null for Cancel and Esc. Enter is OK. */
export function optionsBox(root, { title, text = '', groups = [], checks = [], buttons = [], ok = 'OK' }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    const h = document.createElement('h2');
    h.id = `${id}-title`;
    h.className = 'oxe-dialog-title';
    h.textContent = title;
    dlg.append(h);
    if (text) {
      const p = Object.assign(document.createElement('p'), { id: `${id}-text`, className: 'oxe-dialog-text', textContent: text });
      dlg.setAttribute('aria-describedby', p.id);
      dlg.append(p);
    }
    const radios = new Map();
    groups.forEach((g, gi) => {
      const set = document.createElement('fieldset');
      set.className = 'oxe-dialog-group';
      // (Down each column, then across, as Excel's boxes list them.)
      set.style.setProperty('--oxe-rows', String(Math.ceil(g.options.length / (g.columns || 1))));
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
      if (e.key === 'Enter' && e.target.tagName !== 'BUTTON') { e.preventDefault(); answer(ok); }
    });
    const row = document.createElement('div');
    row.className = 'oxe-dialog-buttons';
    row.append(...buttons.map((label) => button(label, () => answer(label))), button(ok, () => answer(ok)), button('Cancel', () => done(null)));
    dlg.append(row);
    if (!dlg.querySelector('[data-first]')) (dlg.querySelector('input') || row.firstChild).dataset.first = '';
  });
}

/** Excel's Go To box (F5, Ctrl+G): the names and tables to go to in a list, and a Reference field (a pick fills it);
 *  resolves to the text to go to, or null for Cancel and Esc. */
export function goToBox(root, names) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    const h = Object.assign(document.createElement('h2'), { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Go To' });
    const listLabel = Object.assign(document.createElement('label'), { className: 'oxe-dialog-text', htmlFor: `${id}-list`, textContent: 'Go to:' });
    const list = Object.assign(document.createElement('select'), { id: `${id}-list`, className: 'oxe-dialog-list' });
    list.size = Math.max(4, Math.min(8, names.length));
    list.append(...names.map((n) => Object.assign(document.createElement('option'), { value: n, textContent: n })));
    const fieldLabel = Object.assign(document.createElement('label'), { className: 'oxe-dialog-text oxe-dialog-label', htmlFor: `${id}-ref`, textContent: 'Reference:' });
    const field = Object.assign(document.createElement('input'), { id: `${id}-ref`, className: 'oxe-dialog-field', type: 'text', spellcheck: false });
    field.dataset.first = '';
    list.addEventListener('change', () => { field.value = list.value; });
    list.addEventListener('dblclick', () => { if (list.value) done(list.value); });
    const ok = () => done(field.value.trim() || null);
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON') { e.preventDefault(); ok(); } });
    const row = document.createElement('div');
    row.className = 'oxe-dialog-buttons';
    row.append(button('OK', ok), button('Cancel', () => done(null)));
    dlg.append(h, listLabel, list, fieldLabel, field, row);
  });
}

/** Excel's Password box for opening a workbook saved with a password: "'name' is protected.", a Password field (its
 *  text hidden), OK and Cancel. Resolves to the password, or null. */
export function passwordBox(root, fileName) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.setAttribute('aria-describedby', `${id}-text`);
    const h = Object.assign(document.createElement('h2'), { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Password' });
    const text = Object.assign(document.createElement('p'), { id: `${id}-text`, className: 'oxe-dialog-text', textContent: `'${fileName || 'This workbook'}' is protected.` });
    const label = Object.assign(document.createElement('label'), { className: 'oxe-dialog-text oxe-dialog-label', htmlFor: `${id}-pw`, textContent: 'Password:' });
    const field = Object.assign(document.createElement('input'), { id: `${id}-pw`, className: 'oxe-dialog-field', type: 'password', autocomplete: 'off', spellcheck: false });
    field.dataset.first = '';
    const ok = () => done(field.value);
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON') { e.preventDefault(); ok(); } });
    const row = document.createElement('div');
    row.className = 'oxe-dialog-buttons';
    row.append(button('OK', ok), button('Cancel', () => done(null)));
    dlg.append(h, text, label, field, row);
  });
}

/** Excel's note being written or edited, beside the cell (`at`, a screen point): its author's name in bold, the text
 *  in a box to type in. Escape or OK keeps what was typed (Escape keeps it as Excel's note does); Cancel leaves the note
 *  as it was. Resolves to the text, or null. */
export function noteBox(root, { title, author, text = '', at }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-note-edit');
    dlg.style.setProperty('--oxe-fx', `${Math.round(at.x)}px`);
    dlg.style.setProperty('--oxe-fy', `${Math.round(at.y)}px`);
    const h = Object.assign(document.createElement('h2'), { id: `${id}-title`, className: 'oxe-dialog-title', textContent: title });
    const who = Object.assign(document.createElement('p'), { className: 'oxe-note-author', textContent: author ? `${author}:` : '' });
    const area = Object.assign(document.createElement('textarea'), { className: 'oxe-note-text', value: text, rows: 5, spellcheck: true });
    area.setAttribute('aria-labelledby', `${id}-title`);
    area.dataset.first = '';
    area.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); done(area.value); }
      if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); done(area.value); }
    });
    const row = document.createElement('div');
    row.className = 'oxe-dialog-buttons';
    row.append(button('OK', () => done(area.value)), button('Cancel', () => done(null)));
    dlg.append(h, who, area, row);
    queueMicrotask(() => { area.selectionStart = area.selectionEnd = area.value.length; });
  });
}
