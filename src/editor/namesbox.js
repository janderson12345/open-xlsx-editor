// Formulas > Defined Names: Excel's Name Manager (Ctrl+F3: the names with what each refers to, its scope and comment;
// New..., Edit..., Delete, Close) and its New Name / Edit Name box (Name, Scope, Comment, Refers to). Each change is
// one step of the history (names.js through the sheetChange command); a name Excel wouldn't take gets its words.
import { modal, button, messageBox } from './dialogs.js';
import { defineName, editName, deleteName, listNames } from '../ooxml/names.js';
import { sheetPrefix } from '../calc/tokens.js';
import { colName } from '../ooxml/refs.js';

const make = (tag, props = {}, ...kids) => { const e = Object.assign(document.createElement(tag), props); e.append(...kids); return e; };

/** The New Name or Edit Name box: resolves to { name, sheet, comment, refersTo } or null. `check(answer)` gives Excel's
 *  words for one it wouldn't take (shown, the box kept open), or null. */
export function nameBox(root, { title, name = '', sheet = null, comment = '', refersTo = '', scopes, scopeFixed = false, check }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-name-box');
    const field = (key, label, value, tag = 'input') => {
      const input = make(tag, { id: `${id}-${key}`, className: 'oxe-dialog-field', value, ...(tag === 'input' ? { type: 'text', spellcheck: false } : { rows: 3 }) });
      return [make('label', { className: 'oxe-dialog-text oxe-dialog-label', htmlFor: input.id, textContent: label }), input];
    };
    const [nameLabel, nameInput] = field('name', 'Name:', name);
    nameInput.dataset.first = '';
    const scopeSelect = make('select', { id: `${id}-scope`, className: 'oxe-dialog-select', disabled: scopeFixed });
    scopeSelect.append(...scopes.map((s) => make('option', { value: s.value == null ? '' : String(s.value), textContent: s.label, selected: (s.value ?? null) === (sheet ?? null) })));
    const scopeLabel = make('label', { className: 'oxe-dialog-text oxe-dialog-label', htmlFor: scopeSelect.id, textContent: 'Scope:' });
    const [commentLabel, commentInput] = field('comment', 'Comment:', comment, 'textarea');
    const [refLabel, refInput] = field('refers', 'Refers to:', refersTo);
    const why = make('p', { className: 'oxe-dialog-text oxe-pivot-why' });
    why.setAttribute('role', 'alert');
    const ok = () => {
      const answer = { name: nameInput.value.trim(), sheet: scopeSelect.value === '' ? null : Number(scopeSelect.value), comment: commentInput.value, refersTo: refInput.value };
      const problem = check(answer);
      if (problem) { why.textContent = problem; return; }
      done(answer);
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON' && e.target.tagName !== 'TEXTAREA') { e.preventDefault(); ok(); } });
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: title }), nameLabel, nameInput, scopeLabel, scopeSelect, commentLabel, commentInput, refLabel, refInput, why,
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))));
  });
}

/** The Name Manager: the names, one chosen at a time; resolves when closed. `run(fn)` makes a change (fn(wb)) as one
 *  step of the history and returns Excel's words when it refused, or null. */
export function nameManager(root, { workbook, run, newName }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-name-manager');
    const body = make('tbody');
    const table = make('table', { className: 'oxe-pivot-table oxe-names-table' },
      make('thead', {}, make('tr', {}, ...['Name', 'Refers To', 'Scope', 'Comment'].map((h) => make('th', { scope: 'col', textContent: h })))), body);
    const list = make('div', { className: 'oxe-names-list' }, table);
    let chosen = null;
    const fill = () => {
      const names = listNames(workbook());
      body.replaceChildren(...names.map((n, k) => {
        const row = make('tr', { tabIndex: k === 0 ? 0 : -1, className: 'oxe-names-row' }, ...[n.name, n.refersTo, n.scope, n.comment].map((t) => make('td', { textContent: t })));
        row.setAttribute('aria-selected', String(chosen?.name === n.name && chosen?.sheet === n.sheet));
        const pick = () => { chosen = n; for (const r of body.children) r.setAttribute('aria-selected', String(r === row)); };
        row.addEventListener('click', pick);
        row.addEventListener('focus', pick);
        row.addEventListener('dblclick', edit);
        row.addEventListener('keydown', (e) => {
          if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); (e.key === 'ArrowDown' ? row.nextSibling : row.previousSibling)?.focus(); }
        });
        return row;
      }));
      if (!chosen && names[0]) chosen = names[0];
    };
    const scopes = () => [{ label: 'Workbook', value: null }, ...workbook().sheets.map((s, i) => ({ label: s.name, value: i })).filter((s, i) => workbook().sheets[i].kind === 'worksheet')];
    async function edit() {
      if (!chosen) return;
      const was = chosen;
      const answer = await nameBox(root, { title: 'Edit Name', name: was.name, sheet: was.sheet, comment: was.comment, refersTo: was.refersTo, scopes: scopes(), scopeFixed: true,
        check: (a) => run((wb) => editName(wb, { name: was.name, sheet: was.sheet }, { name: a.name, refersTo: a.refersTo, comment: a.comment })) });
      if (answer) chosen = { ...was, name: answer.name };
      fill();
    }
    async function add() {
      await newName();
      fill();
    }
    async function remove() {
      if (!chosen) return;
      const sure = await messageBox(root, `Are you sure you want to delete the name ${chosen.name}?`, ['OK', 'Cancel']);
      if (sure === 'OK') { run((wb) => deleteName(wb, { name: chosen.name, sheet: chosen.sheet })); chosen = null; fill(); }
    }
    fill();
    const top = make('div', { className: 'oxe-dialog-buttons oxe-names-actions' }, button('New...', add), button('Edit...', edit), button('Delete', remove));
    top.firstChild.dataset.first = '';
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Name Manager' }), top, list,
      make('div', { className: 'oxe-dialog-buttons' }, button('Close', () => done(true))));
  });
}

/** Define Name for the selection `sel` of sheet si, as Excel fills its box: the selection (absolute, with its sheet)
 *  as what it refers to, the workbook as its scope. */
export function defineNameFor(root, wb, si, sel, run) {
  const abs = (r, c) => `$${colName(c)}$${r + 1}`;
  const ref = `=${sheetPrefix(wb.sheets[si].name)}${abs(sel.r0, sel.c0)}${sel.r0 === sel.r1 && sel.c0 === sel.c1 ? '' : `:${abs(sel.r1, sel.c1)}`}`;
  const scopes = [{ label: 'Workbook', value: null }, ...wb.sheets.map((s, i) => ({ label: s.name, value: i, kind: s.kind })).filter((s) => s.kind === 'worksheet')];
  return nameBox(root, { title: 'New Name', refersTo: ref, scopes, check: (a) => run((book) => defineName(book, a)) });
}
