// @mentions in a comment box, as Excel 365's: typing @ offers the people (the workbook's persons part and the list the
// host gives, `people`), narrowed by what follows the @; Up and Down move through them, Enter or Tab (or a click) puts
// "@Name" in the text, Esc closes the list. The box's answer carries the people picked; threadedit.js writes a mention
// for each "@Name" of theirs still in the text.

const make = (tag, props = {}) => Object.assign(document.createElement(tag), props);
let seq = 0;

/** The people to offer: the host's list first, then the workbook's, each name once. [{ name, userId, providerId }] */
export function mentionPeople(host = [], book = []) {
  const seen = new Set();
  const out = [];
  for (const p of [...host, ...book]) {
    const name = String(p?.name ?? p?.displayName ?? '').trim();
    if (!name || seen.has(name.toLowerCase())) continue;
    seen.add(name.toLowerCase());
    out.push({ name, ...(p.userId ? { userId: p.userId } : {}), ...(p.providerId ? { providerId: p.providerId } : {}) });
  }
  return out;
}

/** The @ being typed just before the caret: { at (where the @ is), query }, or null. */
export function mentionQuery(text, caret) {
  const before = text.slice(0, caret);
  const m = /(?:^|\s)@([^@\r\n]{0,40})$/.exec(before);
  if (!m) return null;
  // (A name has spaces in it, but two spaces or a space typed after no match end the search.)
  if (/\s\s/.test(m[1])) return null;
  return { at: caret - m[1].length - 1, query: m[1] };
}

/** The people whose name (or user id) has the query in it, those starting with it first; at most eight. */
export function matchPeople(people, query) {
  const q = query.toLowerCase();
  const hit = people.filter((p) => p.name.toLowerCase().includes(q) || (p.userId || '').toLowerCase().includes(q));
  const starts = (p) => (p.name.toLowerCase().startsWith(q) || p.name.toLowerCase().split(/\s+/).some((w) => w.startsWith(q)) ? 0 : 1);
  return hit.sort((a, b) => starts(a) - starts(b)).slice(0, 8);
}

/**
 * A textarea made to offer people on @: returns { el (the list, to put after the textarea), picked() (the people
 * picked, for the answer) }. `already`: people mentioned in the text before (an edit).
 */
export function mentionInput(area, people, already = []) {
  const id = `oxe-mentions-${++seq}`;
  const list = make('ul', { className: 'oxe-mention-list', id, hidden: true });
  list.setAttribute('role', 'listbox');
  list.setAttribute('aria-label', 'People to mention');
  area.setAttribute('role', 'combobox');
  area.setAttribute('aria-autocomplete', 'list');
  area.setAttribute('aria-controls', id);
  area.setAttribute('aria-expanded', 'false');
  const picked = new Map(already.filter((p) => p?.name).map((p) => [p.name, p]));
  let shown = [];
  let active = 0;
  let query = null;
  const close = () => {
    list.hidden = true;
    shown = [];
    area.setAttribute('aria-expanded', 'false');
    area.removeAttribute('aria-activedescendant');
  };
  const mark = () => {
    [...list.children].forEach((li, k) => li.setAttribute('aria-selected', String(k === active)));
    if (shown.length) area.setAttribute('aria-activedescendant', `${id}-${active}`);
  };
  const choose = (p) => {
    const end = area.selectionStart;
    const insert = `@${p.name} `;
    area.value = area.value.slice(0, query.at) + insert + area.value.slice(end);
    const caret = query.at + insert.length;
    area.setSelectionRange(caret, caret);
    picked.set(p.name, p);
    close();
    area.focus();
  };
  const update = () => {
    query = people.length ? mentionQuery(area.value, area.selectionStart) : null;
    shown = query ? matchPeople(people, query.query) : [];
    if (!shown.length) { close(); return; }
    active = Math.min(active, shown.length - 1);
    list.replaceChildren(...shown.map((p, k) => {
      const li = make('li', { id: `${id}-${k}`, className: 'oxe-mention-option', textContent: p.name });
      if (p.userId && p.userId !== p.name) li.append(make('span', { className: 'oxe-mention-id', textContent: ` ${p.userId}` }));
      li.setAttribute('role', 'option');
      li.addEventListener('mousedown', (e) => { e.preventDefault(); choose(p); });
      return li;
    }));
    list.hidden = false;
    area.setAttribute('aria-expanded', 'true');
    mark();
  };
  area.addEventListener('input', () => { active = 0; update(); });
  area.addEventListener('click', update);
  area.addEventListener('blur', close);
  area.addEventListener('keydown', (e) => {
    if (list.hidden || !shown.length) return;
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      active = (active + (e.key === 'ArrowDown' ? 1 : shown.length - 1)) % shown.length;
      mark();
    } else if ((e.key === 'Enter' && !e.ctrlKey && !e.metaKey) || e.key === 'Tab') {
      e.preventDefault();
      e.stopPropagation();
      choose(shown[active]);
    } else if (e.key === 'Escape') {
      // (The list closed, not the box.)
      e.preventDefault();
      e.stopPropagation();
      close();
    }
  });
  return { el: list, picked: () => [...picked.values()] };
}
