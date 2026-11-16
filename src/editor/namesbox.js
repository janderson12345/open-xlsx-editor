// Formulas > Defined Names: Excel's Name Manager (Ctrl+F3: the names with what each refers to, its scope and comment;
// New..., Edit..., Delete, Close) and its New Name / Edit Name box (Name, Scope, Comment, Refers to). Each change is
// one step of the history (names.js through the sheetChange command); a name Excel wouldn't take gets its words.
// Create from Selection's box too (createnames.js), Paste Name's and Apply Names' (applynames.js).
import { modal, button, messageBox, optionsBox } from './dialogs.js';
import { defineName, editName, deleteName, listNames, findName } from '../ooxml/names.js';
import { namesFromSelection, createNames, guessSides } from '../ooxml/createnames.js';
import { selectionRefers } from './selection.js';

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

/** Define Name for the selection `sel` of sheet si, as Excel fills its box: the selection (selectionRefers) as what
 *  it refers to, the workbook as its scope. */
export function defineNameFor(root, wb, si, sel, run) {
  const ref = selectionRefers(wb, si, sel);
  const scopes = [{ label: 'Workbook', value: null }, ...wb.sheets.map((s, i) => ({ label: s.name, value: i, kind: s.kind })).filter((s) => s.kind === 'worksheet')];
  return nameBox(root, { title: 'New Name', refersTo: ref, scopes, check: (a) => run((book) => defineName(book, a)) });
}

/** Create from Selection (Ctrl+Shift+F3) for the selection `sel` of sheet si: Excel's Create Names from Selection box,
 *  its sides checked as guessSides guesses; a name there already asked about first (Excel's question, its words not
 *  measured: COM's CreateNames replaced it without one), Cancel stopping it all. The names made as one step of the
 *  history through `run`. Resolves to how many were made, or null. */
export async function createNamesFor(root, wb, si, sel, run) {
  // (Several ranges: the names of each, as Excel's CreateNames made them: A2:B3 and D6:E7 by their top rows named
  // A3, B3, D7 and E7.)
  const areas = [...(sel.more || []), sel].map((a) => ({ r0: Math.min(a.r0, a.r1), c0: Math.min(a.c0, a.c1), r1: Math.max(a.r0, a.r1), c1: Math.max(a.c0, a.c1) }));
  const start = guessSides(wb, si, areas.at(-1));
  const got = await optionsBox(root, { title: 'Create Names from Selection', text: 'Create names from values in the:',
    checks: [['top', 'Top row'], ['left', 'Left column'], ['bottom', 'Bottom row'], ['right', 'Right column']].map(([key, label]) => ({ key, label, value: start[key] })) });
  if (!got) return null;
  const lists = areas.map((a) => namesFromSelection(wb, si, a, got.checks));
  const list = lists.some((x) => !x) ? null : lists.flat();
  if (!list) { await messageBox(root, "This selection isn't valid."); return null; }
  const keep = [];
  for (const x of list) {
    const had = findName(wb, x.name, null) && !keep.some((k) => k.name.toLowerCase() === x.name.toLowerCase());
    if (had) {
      const answer = await messageBox(root, `Do you want to replace the existing definition of '${x.name}'?`, ['Yes', 'No', 'Cancel']);
      if (answer !== 'Yes' && answer !== 'No') return null;
      if (answer === 'No') continue;
    }
    keep.push(x);
  }
  if (keep.length) run((book) => createNames(book, keep));
  return keep.length;
}

/** Excel's Paste Name box (F3, Use in Formula's Paste Names...): `names` in a list, OK, Cancel and, unless a cell is
 *  being edited (`list: false`, as Excel's is then dimmed), Paste List. Resolves to { name }, { list: true }, or null. */
export function pasteNameBox(root, names, { list = true } = {}) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    const label = make('label', { className: 'oxe-dialog-text oxe-dialog-label', htmlFor: `${id}-names`, textContent: 'Paste name' });
    const pick = make('select', { id: `${id}-names`, className: 'oxe-dialog-select oxe-paste-names', size: Math.min(10, Math.max(2, names.length)) });
    pick.append(...names.map((n, k) => make('option', { value: n, textContent: n, selected: k === 0 })));
    pick.dataset.first = '';
    const ok = () => (pick.value ? done({ name: pick.value }) : done(null));
    pick.addEventListener('dblclick', ok);
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON') { e.preventDefault(); ok(); } });
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Paste Name' }), label, pick,
      make('div', { className: 'oxe-dialog-buttons' }, ...[button('OK', ok), button('Cancel', () => done(null)), list ? button('Paste List', () => done({ list: true })) : null].filter(Boolean)));
  });
}

/** Excel's Apply Names box (Formulas > Define Name > Apply Names...): the names in a list to choose from (all chosen
 *  here; Excel's first choice not measured), Ignore Relative/Absolute and Use row and column names (both on, as
 *  Excel's), Omit column name if same column, Omit row name if same row (on) and the name order (Row column).
 *  Resolves to { names, ignore, rowCol, omitCol, omitRow, colFirst } or null. */
export function applyNamesBox(root, names) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-apply-names');
    const pick = make('select', { id: `${id}-names`, className: 'oxe-dialog-select', multiple: true, size: Math.min(10, Math.max(3, names.length)) });
    pick.append(...names.map((n) => make('option', { value: n, textContent: n, selected: true })));
    pick.dataset.first = '';
    const check = (key, label, on = true) => {
      const input = make('input', { type: 'checkbox', id: `${id}-${key}`, checked: on });
      return { input, line: make('div', { className: 'oxe-dialog-option' }, input, make('label', { htmlFor: input.id, textContent: label })) };
    };
    const ignore = check('ignore', 'Ignore Relative/Absolute');
    const rowCol = check('rowcol', 'Use row and column names');
    const omitCol = check('omitcol', 'Omit column name if same column');
    const omitRow = check('omitrow', 'Omit row name if same row');
    const order = ['Row column', 'Column row'].map((label, k) => {
      const input = make('input', { type: 'radio', name: `${id}-order`, id: `${id}-order${k}`, checked: k === 0 });
      return { input, line: make('div', { className: 'oxe-dialog-option' }, input, make('label', { htmlFor: input.id, textContent: label })) };
    });
    // (Excel dims the row and column options while Use row and column names is off.)
    const dim = () => [omitCol, omitRow, ...order].forEach((x) => { x.input.disabled = !rowCol.input.checked; });
    rowCol.input.addEventListener('change', dim);
    const ok = () => {
      const chosen = [...pick.selectedOptions].map((o) => o.value);
      done({ names: chosen, ignore: ignore.input.checked, rowCol: rowCol.input.checked, omitCol: omitCol.input.checked, omitRow: omitRow.input.checked, colFirst: order[1].input.checked });
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON') { e.preventDefault(); ok(); } });
    const orderSet = make('fieldset', { className: 'oxe-dialog-group' }, make('legend', { textContent: 'Name order' }), ...order.map((x) => x.line));
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Apply Names' }),
      make('label', { className: 'oxe-dialog-text oxe-dialog-label', htmlFor: pick.id, textContent: 'Apply names:' }), pick,
      make('div', { className: 'oxe-dialog-checks' }, ignore.line, rowCol.line, omitCol.line, omitRow.line), orderSet,
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))));
    dim();
  });
}
