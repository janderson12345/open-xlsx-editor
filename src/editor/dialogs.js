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

/** Excel's orders for a column of text, numbers or dates. */
export const SORT_ORDERS = { text: ['A to Z', 'Z to A'], number: ['Smallest to Largest', 'Largest to Smallest'], date: ['Oldest to Newest', 'Newest to Oldest'] };

/**
 * Excel's Sort box (Custom Sort...): its levels, each a column (a row, sorting left to right) and an order, with Add
 * Level, Delete Level, Copy Level, Move Up and Move Down; My data has headers; Options... (Case sensitive, the
 * orientation). `lines(leftToRight, header)` gives the columns (rows) to choose from: [{ at, label, kind: 'text' |
 * 'number' | 'date' }]. `start`: { keys: [{ at, desc }], header, matchCase, leftToRight }. Resolves to the same shape
 * as chosen, or null for Cancel and Esc.
 */
export function sortBox(root, { lines, start }) {
  return modal(root, (dlg, done, id) => {
    const now = { ...start, keys: start.keys.map((k) => ({ ...k })) };
    let current = 0;
    const make = (tag, props = {}, ...kids) => {
      const x = document.createElement(tag);
      for (const [k, v] of Object.entries(props)) { if (k === 'role' || k.startsWith('aria-')) x.setAttribute(k, v); else x[k] = v; }
      x.append(...kids);
      return x;
    };
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-sort-box');
    const head = make('input', { type: 'checkbox', id: `${id}-head` });
    const levels = make('div', { className: 'oxe-sort-levels', role: 'list', 'aria-label': 'Sort levels' });
    const tools = make('div', { className: 'oxe-sort-tools' });
    const tool = (label, run) => { const b = button(label, run); tools.append(b); return b; };
    const draw = (focus) => {
      const list = lines(now.leftToRight, now.header);
      head.checked = now.header && !now.leftToRight;
      head.disabled = now.leftToRight;
      levels.replaceChildren(...now.keys.map((k, i) => {
        const word = i === 0 ? 'Sort by' : 'Then by';
        const col = make('select', { className: 'oxe-sort-select', 'aria-label': `${word}: ${now.leftToRight ? 'row' : 'column'}` },
          ...list.map((x) => make('option', { value: String(x.at), textContent: x.label })));
        if (!list.some((x) => x.at === k.at)) k.at = list[0]?.at ?? k.at;
        col.value = String(k.at);
        const kind = () => list.find((x) => x.at === k.at)?.kind || 'text';
        const order = make('select', { className: 'oxe-sort-select', 'aria-label': `${word}: order` });
        const fill = () => { order.replaceChildren(...SORT_ORDERS[kind()].map((label, d) => make('option', { value: String(d), textContent: label }))); order.value = k.desc ? '1' : '0'; };
        fill();
        col.addEventListener('change', () => { k.at = Number(col.value); fill(); });
        order.addEventListener('change', () => { k.desc = order.value === '1'; });
        const row = make('div', { className: `oxe-sort-level${i === current ? ' oxe-on' : ''}`, role: 'listitem' }, make('span', { className: 'oxe-sort-word', textContent: word }), col, order);
        row.addEventListener('focusin', () => { if (current !== i) { current = i; levels.querySelectorAll('.oxe-sort-level').forEach((x, j) => x.classList.toggle('oxe-on', j === i)); } });
        return row;
      }));
      del.disabled = now.keys.length < 2;
      up.disabled = current === 0;
      down.disabled = current >= now.keys.length - 1;
      if (focus) levels.children[current]?.querySelector('select')?.focus();
    };
    const add = (key) => { now.keys.splice(current + 1, 0, key); current += 1; draw(true); };
    tool('Add Level', () => add({ at: lines(now.leftToRight, now.header)[0]?.at ?? 0, desc: false }));
    const del = tool('Delete Level', () => { now.keys.splice(current, 1); current = Math.min(current, now.keys.length - 1); draw(true); });
    tool('Copy Level', () => add({ ...now.keys[current] }));
    const move = (d) => { const [k] = now.keys.splice(current, 1); current += d; now.keys.splice(current, 0, k); draw(true); };
    const up = tool('Move Up', () => move(-1));
    const down = tool('Move Down', () => move(1));
    tool('Options...', async () => {
      const got = await optionsBox(root, { title: 'Sort Options', checks: [{ key: 'matchCase', label: 'Case sensitive', value: now.matchCase }],
        groups: [{ key: 'way', legend: 'Orientation', options: ['Sort top to bottom', 'Sort left to right'], value: now.leftToRight ? 'Sort left to right' : 'Sort top to bottom' }] });
      if (got) { now.matchCase = got.checks.matchCase; now.leftToRight = got.choices.way === 'Sort left to right'; draw(false); }
    });
    head.addEventListener('change', () => { now.header = head.checked; draw(false); });
    const headLine = make('div', { className: 'oxe-dialog-option oxe-sort-head' }, head, make('label', { htmlFor: head.id, textContent: 'My data has headers' }));
    const ok = () => done({ ...now, header: now.header && !now.leftToRight });
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON') { e.preventDefault(); ok(); } });
    const row = make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null)));
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Sort' }), make('div', { className: 'oxe-sort-top' }, tools, headLine), levels, row);
    draw(false);
    levels.querySelector('select').dataset.first = '';
  });
}

/**
 * A filter button's list, as Excel's: Sort A to Z and Z to A (Smallest to Largest for numbers), Clear Filter From the
 * column, a search box, and the column's values to tick, with (Select All) and (Blanks); OK keeps the rows ticked.
 * `at`: { x, y } where it opens (the button's bottom left, in the window). `choices`: filterops.js filterChoices.
 * Resolves to { do: 'asc' | 'desc' | 'clear' } or { do: 'apply', values, blank, all }, or null for Cancel and Esc.
 */
export function filterBox(root, { title, at, choices, numbers = false, filtered = false }) {
  return modal(root, (dlg, done, id) => {
    const make = (tag, props = {}, ...kids) => {
      const x = document.createElement(tag);
      for (const [k, v] of Object.entries(props)) { if (k === 'role' || k.startsWith('aria-')) x.setAttribute(k, v); else x[k] = v; }
      x.append(...kids);
      return x;
    };
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-filter-box');
    dlg.style.setProperty('--oxe-fx', `${Math.round(at.x)}px`);
    dlg.style.setProperty('--oxe-fy', `${Math.round(at.y)}px`);
    const item = (label, run, disabled = false) => { const b = button(label, run); b.classList.add('oxe-filter-item'); b.disabled = disabled; return b; };
    const head = make('div', { className: 'oxe-filter-actions' },
      item(numbers ? 'Sort Smallest to Largest' : 'Sort A to Z', () => done({ do: 'asc' })),
      item(numbers ? 'Sort Largest to Smallest' : 'Sort Z to A', () => done({ do: 'desc' })),
      item(`Clear Filter From "${title}"`, () => done({ do: 'clear' }), !filtered));
    const search = make('input', { type: 'search', className: 'oxe-dialog-field', id: `${id}-search`, placeholder: 'Search', 'aria-label': 'Search' });
    const list = make('div', { className: 'oxe-filter-list', role: 'group', 'aria-label': 'Values' });
    const entries = choices.values.map((v) => ({ label: v.text === '' ? ' ' : v.text, value: v.text, checked: v.checked }));
    if (choices.blanks != null) entries.push({ label: '(Blanks)', blank: true, checked: choices.blanks });
    const all = make('input', { type: 'checkbox', id: `${id}-all` });
    const boxes = entries.map((e, i) => {
      const input = make('input', { type: 'checkbox', id: `${id}-v${i}`, checked: e.checked });
      input.addEventListener('change', () => { sync(); });
      return { e, input, line: make('div', { className: 'oxe-dialog-option' }, input, make('label', { htmlFor: `${id}-v${i}`, textContent: e.label })) };
    });
    const okBtn = button('OK', () => apply());
    const sync = () => {
      const shown = boxes.filter((b) => !b.line.hidden);
      const on = shown.filter((b) => b.input.checked).length;
      all.checked = on === shown.length && on > 0;
      all.indeterminate = on > 0 && on < shown.length;
      okBtn.disabled = !boxes.some((b) => b.input.checked);
    };
    all.addEventListener('change', () => { for (const b of boxes) if (!b.line.hidden) b.input.checked = all.checked; sync(); });
    search.addEventListener('input', () => {
      const q = search.value.trim().toLowerCase();
      for (const b of boxes) b.line.hidden = !!q && !b.e.label.toLowerCase().includes(q);
      sync();
    });
    list.append(make('div', { className: 'oxe-dialog-option' }, all, make('label', { htmlFor: all.id, textContent: '(Select All)' })), ...boxes.map((b) => b.line));
    const apply = () => {
      if (okBtn.disabled) return;
      // (With a search typed, only the values it found count, as in Excel.)
      const picked = boxes.filter((b) => !b.line.hidden && b.input.checked);
      const every = picked.length === boxes.length;
      done({ do: 'apply', all: every, values: picked.filter((b) => !b.e.blank).map((b) => b.e.value), blank: picked.some((b) => b.e.blank) });
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON') { e.preventDefault(); apply(); } });
    const row = make('div', { className: 'oxe-dialog-buttons' }, okBtn, button('Cancel', () => done(null)));
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title oxe-filter-title', textContent: `Filter: ${title}` }), head, search, list, row);
    sync();
    search.dataset.first = '';
  });
}

/** A data validation list's choices under its cell, as a list to pick from (`items`: texts): resolves to the index
 *  picked (click, double click or Enter), or null for Esc. */
export function pickBox(root, { title, at, items }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-label', title);
    dlg.classList.add('oxe-pick-box');
    dlg.style.setProperty('--oxe-fx', `${Math.round(at.x)}px`);
    dlg.style.setProperty('--oxe-fy', `${Math.round(at.y)}px`);
    const list = document.createElement('select');
    list.id = `${id}-list`;
    list.className = 'oxe-dialog-list';
    list.setAttribute('aria-label', title);
    list.size = Math.max(2, Math.min(8, items.length));
    list.append(...items.map((t, i) => Object.assign(document.createElement('option'), { value: String(i), textContent: t })));
    list.selectedIndex = 0;
    list.dataset.first = '';
    const ok = () => done(list.selectedIndex < 0 ? null : list.selectedIndex);
    list.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); ok(); } });
    list.addEventListener('click', (e) => { if (e.target.tagName === 'OPTION') ok(); });
    dlg.append(list);
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

/**
 * Excel's Data Validation box, its three tabs: Settings (Allow, Data, the values, Ignore blank, In-cell dropdown),
 * Input Message and Error Alert; Clear All, OK and Cancel. `rule` the active cell's (dvedit.js's form: { type, op, f1,
 * f2, ignoreBlank, dropDown, prompt: { show, title, text }, error: { show, style, title, text } }), or null. Resolves to
 * the rule, 'clear' (Clear All, then OK), or null.
 */
export function validationBox(root, rule = null) {
  const r = rule || { type: 'any', op: 'between', f1: '', f2: '', ignoreBlank: true, dropDown: true, prompt: { show: true, title: '', text: '' }, error: { show: true, style: 'stop', title: '', text: '' } };
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-dv-box');
    const el = (tag, props = {}, ...kids) => { const e = Object.assign(document.createElement(tag), props); e.append(...kids); return e; };
    const field = (label, input) => { input.id ||= `${id}-${Math.random().toString(36).slice(2, 8)}`; return el('div', { className: 'oxe-dv-field' }, el('label', { htmlFor: input.id, textContent: label }), input); };
    const select = (options, value) => { const s = el('select', { className: 'oxe-dialog-list' }); s.append(...options.map(([v, t]) => el('option', { value: v, textContent: t }))); s.value = value; return s; };
    const check = (label, on) => { const c = el('input', { type: 'checkbox', checked: on }); return [c, el('label', { className: 'oxe-dv-check' }, c, ` ${label}`)]; };
    const text = (value) => el('input', { type: 'text', className: 'oxe-dialog-field', value: value || '', spellcheck: false });
    // Settings.
    const allow = select([['any', 'Any value'], ['whole', 'Whole number'], ['decimal', 'Decimal'], ['list', 'List'], ['date', 'Date'], ['time', 'Time'], ['length', 'Text length'], ['custom', 'Custom']], r.type);
    allow.dataset.first = '';
    const op = select([['between', 'between'], ['notBetween', 'not between'], ['equal', 'equal to'], ['notEqual', 'not equal to'], ['greater', 'greater than'], ['less', 'less than'], ['greaterEqual', 'greater than or equal to'], ['lessEqual', 'less than or equal to']], r.op || 'between');
    const f1 = text(r.f1);
    const f2 = text(r.f2);
    const [blank, blankRow] = check('Ignore blank', r.ignoreBlank !== false);
    const [drop, dropRow] = check('In-cell dropdown', r.dropDown !== false);
    const opField = field('Data:', op);
    const f1Field = field('Minimum:', f1);
    const f2Field = field('Maximum:', f2);
    const settings = el('div', { className: 'oxe-dv-panel' }, field('Allow:', allow), blankRow, opField, f1Field, f2Field, dropRow);
    const layout = () => {
      const t = allow.value;
      const ranged = !['any', 'list', 'custom'].includes(t);
      const two = ranged && (op.value === 'between' || op.value === 'notBetween');
      opField.hidden = !ranged;
      f1Field.hidden = t === 'any';
      f2Field.hidden = !two;
      dropRow.hidden = t !== 'list';
      blankRow.hidden = t === 'any';
      f1Field.querySelector('label').textContent = t === 'list' ? 'Source:' : t === 'custom' ? 'Formula:' : two ? 'Minimum:' : ['greater', 'greaterEqual'].includes(op.value) ? 'Minimum:' : ['less', 'lessEqual'].includes(op.value) ? 'Maximum:' : 'Value:';
    };
    allow.addEventListener('change', layout);
    op.addEventListener('change', layout);
    layout();
    // Input Message, Error Alert.
    const [showPrompt, showPromptRow] = check('Show input message when cell is selected', r.prompt?.show !== false);
    const promptTitle = text(r.prompt?.title);
    const promptText = el('textarea', { className: 'oxe-dialog-field', rows: 3, value: r.prompt?.text || '' });
    const message = el('div', { className: 'oxe-dv-panel' }, showPromptRow, field('Title:', promptTitle), field('Input message:', promptText));
    const [showError, showErrorRow] = check('Show error alert after invalid data is entered', r.error?.show !== false);
    const style = select([['stop', 'Stop'], ['warning', 'Warning'], ['information', 'Information']], r.error?.style || 'stop');
    const errorTitle = text(r.error?.title);
    const errorText = el('textarea', { className: 'oxe-dialog-field', rows: 3, value: r.error?.text || '' });
    const alertPanel = el('div', { className: 'oxe-dv-panel' }, showErrorRow, field('Style:', style), field('Title:', errorTitle), field('Error message:', errorText));
    // The tabs (a tablist, as ARIA describes one).
    const panels = [['Settings', settings], ['Input Message', message], ['Error Alert', alertPanel]];
    const list = el('div', { className: 'oxe-dv-tabs' });
    list.setAttribute('role', 'tablist');
    list.setAttribute('aria-label', 'Data Validation');
    const tabs = panels.map(([name, panel], k) => {
      const tab = el('button', { type: 'button', id: `${id}-tab${k}`, textContent: name, className: 'oxe-dv-tab' });
      tab.setAttribute('role', 'tab');
      panel.id = `${id}-panel${k}`;
      panel.setAttribute('role', 'tabpanel');
      panel.setAttribute('aria-labelledby', tab.id);
      tab.setAttribute('aria-controls', panel.id);
      return tab;
    });
    const show = (k) => tabs.forEach((t, j) => { t.setAttribute('aria-selected', String(j === k)); t.tabIndex = j === k ? 0 : -1; panels[j][1].hidden = j !== k; });
    tabs.forEach((t, k) => {
      t.addEventListener('click', () => show(k));
      t.addEventListener('keydown', (e) => {
        const step = e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0;
        if (step) { e.preventDefault(); const n = (k + step + tabs.length) % tabs.length; show(n); tabs[n].focus(); }
      });
    });
    list.append(...tabs);
    show(0);
    let cleared = false;
    const clearAll = () => {
      cleared = true;
      allow.value = 'any'; op.value = 'between'; f1.value = ''; f2.value = '';
      blank.checked = true; drop.checked = true; showPrompt.checked = true; showError.checked = true;
      promptTitle.value = ''; promptText.value = ''; style.value = 'stop'; errorTitle.value = ''; errorText.value = '';
      layout();
    };
    const ok = () => {
      if (allow.value === 'any' && cleared && !promptTitle.value && !promptText.value) { done('clear'); return; }
      done({
        type: allow.value, op: op.value, f1: f1.value, f2: f2.value, ignoreBlank: blank.checked, dropDown: drop.checked,
        prompt: { show: showPrompt.checked, title: promptTitle.value, text: promptText.value },
        error: { show: showError.checked, style: style.value, title: errorTitle.value, text: errorText.value },
      });
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName === 'INPUT' && e.target.type === 'text') { e.preventDefault(); ok(); } });
    const h = el('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Data Validation' });
    const row = el('div', { className: 'oxe-dialog-buttons' });
    row.append(button('Clear All', clearAll), button('OK', ok), button('Cancel', () => done(null)));
    dlg.append(h, list, settings, message, alertPanel, row);
  });
}
