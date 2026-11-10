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
