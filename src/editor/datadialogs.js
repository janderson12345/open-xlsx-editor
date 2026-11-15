// Excel's boxes for data (dialogs.js has the others and the modal they share): Sort, a filter's list, a list's choices,
// Data Validation and the Conditional Formatting menu's rule boxes.
import { modal, button, optionsBox } from './dialogs.js';

/** Excel's orders for a column of text, numbers or dates. */
export const SORT_ORDERS = { text: ['A to Z', 'Z to A'], number: ['Smallest to Largest', 'Largest to Smallest'], date: ['Oldest to Newest', 'Newest to Oldest'] };

/**
 * Excel's Sort box (Custom Sort...): its levels, each a column (a row, sorting left to right) and an order, with Add
 * Level, Delete Level, Copy Level, Move Up and Move Down; My data has headers; Options... (Case sensitive, the
 * orientation). Each level's Sort On: Cell Values (an order, or Custom List... typed as "High, Medium, Low"), Cell
 * Color or Font Color (a colour the column has, On Top or On Bottom). `lines(leftToRight, header)` gives the columns
 * (rows) to choose from: [{ at, label, kind: 'text' | 'number' | 'date', colors: { cellColor: [argb], fontColor: [argb] } }].
 * `start`: { keys: [{ at, desc, on, color, customOrder }], header, matchCase, leftToRight }. Resolves to the same shape
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
        const line = () => list.find((x) => x.at === k.at) || {};
        const kind = () => line().kind || 'text';
        const on = make('select', { className: 'oxe-sort-select', 'aria-label': `${word}: sort on` },
          ...[['values', 'Cell Values'], ['cellColor', 'Cell Color'], ['fontColor', 'Font Color']].map(([v, t]) => make('option', { value: v, textContent: t })));
        on.value = k.on || 'values';
        const colour = make('select', { className: 'oxe-sort-select oxe-sort-colour', 'aria-label': `${word}: colour` });
        const order = make('select', { className: 'oxe-sort-select', 'aria-label': `${word}: order` });
        const custom = make('input', { type: 'text', className: 'oxe-dialog-field oxe-sort-custom', placeholder: 'High, Medium, Low', 'aria-label': `${word}: custom list`, value: (k.customOrder || []).join(', ') });
        const fill = () => {
          const byColour = on.value !== 'values';
          colour.hidden = !byColour;
          if (byColour) {
            const have = line().colors?.[on.value] || [];
            colour.replaceChildren(...have.map((c) => { const o = make('option', { value: c, textContent: `#${c.slice(2)}` }); o.style.background = `#${c.slice(2)}`; return o; }));
            if (!have.includes(k.color)) k.color = have[0] || null;
            if (k.color) colour.value = k.color;
            order.replaceChildren(...['On Top', 'On Bottom'].map((label, d) => make('option', { value: String(d), textContent: label })));
            order.value = k.desc ? '1' : '0';
          } else {
            order.replaceChildren(...[...SORT_ORDERS[kind()], 'Custom List...'].map((label, d) => make('option', { value: d === 2 ? 'custom' : String(d), textContent: label })));
            order.value = k.customOrder ? 'custom' : k.desc ? '1' : '0';
          }
          custom.hidden = byColour || !k.customOrder;
        };
        fill();
        col.addEventListener('change', () => { k.at = Number(col.value); fill(); });
        on.addEventListener('change', () => { k.on = on.value === 'values' ? undefined : on.value; k.customOrder = undefined; fill(); });
        colour.addEventListener('change', () => { k.color = colour.value; });
        order.addEventListener('change', () => {
          if (order.value === 'custom') { k.customOrder = custom.value.split(',').map((x) => x.trim()).filter(Boolean); k.desc = false; custom.hidden = false; custom.focus(); return; }
          k.customOrder = undefined; k.desc = order.value === '1'; custom.hidden = true;
        });
        custom.addEventListener('input', () => { k.customOrder = custom.value.split(',').map((x) => x.trim()).filter(Boolean); });
        const row = make('div', { className: `oxe-sort-level${i === current ? ' oxe-on' : ''}`, role: 'listitem' }, make('span', { className: 'oxe-sort-word', textContent: word }), col, on, colour, order, custom);
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
 * A column of dates lists them by year, month and day, as Excel's list does, each year and month a tick box over the
 * dates under it. Resolves to { do: 'asc' | 'desc' | 'clear' } or { do: 'apply', values, blank, all, dates (the date
 * groups ticked: a whole year, a whole month, else days) }, or null for Cancel and Esc.
 */
export function filterBox(root, { title, at, choices, numbers = false, filtered = false, kind = numbers ? 'Number' : 'Text' }) {
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
      item(`Clear Filter From "${title}"`, () => done({ do: 'clear' }), !filtered),
      // (Excel's Text, Number or Date Filters and Filter by Color: filterbox.js's box.)
      item(`${kind} Filters...`, () => done({ do: 'custom' })));
    const search = make('input', { type: 'search', className: 'oxe-dialog-field', id: `${id}-search`, placeholder: 'Search', 'aria-label': 'Search' });
    const list = make('div', { className: 'oxe-filter-list', role: 'group', 'aria-label': 'Values' });
    const entries = choices.values.map((v) => ({ label: v.text === '' ? ' ' : v.text, value: v.text, checked: v.checked }));
    if (choices.blanks != null) entries.push({ label: '(Blanks)', blank: true, checked: choices.blanks });
    const all = make('input', { type: 'checkbox', id: `${id}-all` });
    const dated = choices.values.length > 0 && choices.values.every((v) => v.date);
    const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    choices.values.forEach((v, i) => { if (dated) { entries[i].date = v.date; entries[i].label = String(v.date.d); } });
    const boxes = entries.map((e, i) => {
      const input = make('input', { type: 'checkbox', id: `${id}-v${i}`, checked: e.checked });
      input.addEventListener('change', () => { sync(); });
      return { e, input, line: make('div', { className: `oxe-dialog-option${e.date ? ' oxe-filter-day' : ''}` }, input, make('label', { htmlFor: `${id}-v${i}`, textContent: e.label })) };
    });
    // (Dates: a year's and a month's tick boxes, each over the days under it.)
    const groups = [];
    if (dated) {
      let n = 0;
      const group = (label, members, cls) => {
        const input = make('input', { type: 'checkbox', id: `${id}-g${n++}` });
        input.addEventListener('change', () => { for (const b of members) if (!b.line.hidden) b.input.checked = input.checked; sync(); });
        const g = { input, members, line: make('div', { className: `oxe-dialog-option ${cls}` }, input, make('label', { htmlFor: input.id, textContent: label })) };
        groups.push(g);
        return g;
      };
      const days = boxes.filter((b) => b.e.date);
      for (const y of [...new Set(days.map((b) => b.e.date.y))].sort((a, b) => a - b)) {
        const inYear = days.filter((b) => b.e.date.y === y);
        const yearGroup = group(String(y), inYear, 'oxe-filter-year');
        yearGroup.months = [...new Set(inYear.map((b) => b.e.date.m))].sort((a, b) => a - b).map((m) => group(MONTH_NAMES[m - 1], inYear.filter((b) => b.e.date.m === m).sort((a, b) => a.e.date.d - b.e.date.d), 'oxe-filter-month'));
      }
    }
    const okBtn = button('OK', () => apply());
    const sync = () => {
      for (const g of groups) {
        const on = g.members.filter((b) => b.input.checked).length;
        g.input.checked = on === g.members.length;
        g.input.indeterminate = on > 0 && on < g.members.length;
        g.line.hidden = g.members.every((b) => b.line.hidden);
      }
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
    // (Dates: each year, its months, each month's days; then the rest, blanks last.)
    const ordered = dated
      ? [...groups.filter((g) => g.months).flatMap((y) => [y.line, ...y.months.flatMap((m) => [m.line, ...m.members.map((b) => b.line)])]), ...boxes.filter((b) => !b.e.date).map((b) => b.line)]
      : boxes.map((b) => b.line);
    list.append(make('div', { className: 'oxe-dialog-option' }, all, make('label', { htmlFor: all.id, textContent: '(Select All)' })), ...ordered);
    const apply = () => {
      if (okBtn.disabled) return;
      // (With a search typed, only the values it found count, as in Excel.)
      const picked = boxes.filter((b) => !b.line.hidden && b.input.checked);
      const every = picked.length === boxes.length;
      // (Dates as Excel writes them: a year whose dates are all ticked, else a month, else the days.)
      const dates = [];
      if (dated) {
        for (const y of groups.filter((g) => g.months)) {
          if (y.members.every((b) => picked.includes(b))) { dates.push({ year: y.members[0].e.date.y }); continue; }
          for (const m of y.months) {
            const on = m.members.filter((b) => picked.includes(b));
            if (on.length && on.length === m.members.length) dates.push({ year: on[0].e.date.y, month: on[0].e.date.m });
            else for (const b of on) dates.push({ year: b.e.date.y, month: b.e.date.m, day: b.e.date.d });
          }
        }
      }
      done({ do: 'apply', all: every, values: picked.filter((b) => !b.e.blank && !b.e.date).map((b) => b.e.value), blank: picked.some((b) => b.e.blank), ...(dated ? { dates } : {}) });
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

/** One of Excel's Conditional Formatting menu boxes: `text` ("Format cells that are GREATER THAN:"), its `inputs`
 *  ([{ label, value }]), a `choice` ([{ value, label }], Duplicate or Unique, a date's period; `choiceLabel` its label) and the format to apply from `presets`
 *  ([{ value, label }]). Resolves to { values, choice, preset }, or null. */
export function cfRuleBox(root, { title, text, inputs = [], choice = null, choiceLabel = 'Format cells that contain:', presets, preset }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-cf-box');
    const h = Object.assign(document.createElement('h2'), { id: `${id}-title`, className: 'oxe-dialog-title', textContent: title });
    const p = Object.assign(document.createElement('p'), { className: 'oxe-dialog-text', textContent: text });
    const fields = inputs.map((x, k) => {
      const input = Object.assign(document.createElement('input'), { type: 'text', className: 'oxe-dialog-field', value: x.value ?? '', spellcheck: false, id: `${id}-in${k}` });
      if (k === 0) input.dataset.first = '';
      const label = Object.assign(document.createElement('label'), { htmlFor: input.id, className: 'oxe-dialog-text', textContent: x.label });
      return [label, input];
    });
    const select = (opts, value, label, first) => {
      const s = Object.assign(document.createElement('select'), { className: 'oxe-dialog-list', id: `${id}-${label.replace(/\W/g, '')}` });
      s.append(...opts.map((o) => Object.assign(document.createElement('option'), { value: o.value, textContent: o.label })));
      s.value = value;
      if (first) s.dataset.first = '';
      return [Object.assign(document.createElement('label'), { htmlFor: s.id, className: 'oxe-dialog-text', textContent: label }), s];
    };
    const pick = choice ? select(choice, choice[0].value, choiceLabel, !inputs.length) : null;
    const fmt = select(presets, preset, 'with', !inputs.length && !choice);
    const ok = () => done({ values: fields.map(([, i]) => i.value), choice: pick?.[1].value ?? null, preset: fmt[1].value });
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName === 'INPUT') { e.preventDefault(); ok(); } });
    const row = document.createElement('div');
    row.className = 'oxe-dialog-buttons';
    row.append(button('OK', ok), button('Cancel', () => done(null)));
    dlg.append(h, p, ...fields.flat(), ...(pick || []), ...fmt, row);
  });
}
