// Review > Allow Edit Ranges: Excel's Allow Users to Edit Ranges box (the sheet's ranges, each its title and the cells
// it refers to; New..., Modify..., Delete, which can't be pressed while the sheet is protected, as in Excel) and its New
// Range / Modify Range box (Title, Refers to cells, Range password, asked for twice). ooxml/editranges.js writes them.
import { modal, button, messageBox } from './dialogs.js';
import { passwordAsk } from './protectbox.js';
import { editRanges } from '../ooxml/editranges.js';
import { parseRange, refName } from '../ooxml/refs.js';

const make = (tag, props = {}, ...kids) => { const e = Object.assign(document.createElement(tag), props); e.append(...kids); return e; };
const abs = (r, c) => refName(r, c).replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2');

/** A range's cells as the box shows them: "=$B$2:$C$3,$H$4" for sqref "B2:C3 H4". */
export function refersOf(areas) {
  return `=${areas.map((a) => (a.r0 === a.r1 && a.c0 === a.c1 ? abs(a.r0, a.c0) : `${abs(a.r0, a.c0)}:${abs(a.r1, a.c1)}`)).join(',')}`;
}

/** What the box's Refers to cells holds, as a sqref ("B2:C3 H4"), or null when it isn't cells of this sheet. */
export function sqrefOfRefers(text) {
  const parts = String(text || '').trim().replace(/^=/, '').split(',').map((p) => p.trim().replace(/\$/g, ''));
  const areas = parts.map(parseRange);
  if (!parts.length || areas.some((a) => !a)) return null;
  return parts.map((p) => p.toUpperCase()).join(' ');
}

/** The New Range or Modify Range box: resolves to { name, sqref, password } (password null: left as it was), or null.
 *  `taken(name)`: whether another range has that title. */
function rangeBox(root, { title, name, refers, modify, taken }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    const field = (key, label, value, type = 'text') => {
      const input = make('input', { id: `${id}-${key}`, className: 'oxe-dialog-field', type, value, spellcheck: false, ...(type === 'password' ? { autocomplete: 'new-password' } : {}) });
      return [make('label', { className: 'oxe-dialog-text oxe-dialog-label', htmlFor: input.id, textContent: label }), input];
    };
    const [nameLabel, nameInput] = field('name', 'Title:', name);
    nameInput.dataset.first = '';
    const [refLabel, refInput] = field('refers', 'Refers to cells:', refers);
    const [pwLabel, pwInput] = field('pw', 'Range password:', '', 'password');
    if (modify) pwInput.placeholder = '(unchanged)';
    const why = make('p', { className: 'oxe-dialog-text oxe-pivot-why' });
    why.setAttribute('role', 'alert');
    const ok = async () => {
      const answer = { name: nameInput.value.trim(), sqref: sqrefOfRefers(refInput.value), password: pwInput.value || (modify ? null : '') };
      if (!answer.name) { why.textContent = 'Enter a title for the range.'; return; }
      if (taken(answer.name)) { why.textContent = 'A range with this title already exists.'; return; }
      if (!answer.sqref) { why.textContent = 'Enter cells of this sheet, such as =$B$2:$C$3.'; return; }
      if (answer.password) {
        const again = await passwordAsk(root, { title: 'Confirm Password', label: 'Reenter password to proceed.' });
        if (again == null) return;
        if (again !== answer.password) { await messageBox(root, 'Confirmation password is not identical.'); return; }
      }
      done(answer);
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON') { e.preventDefault(); ok(); } });
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: title }), nameLabel, nameInput, refLabel, refInput, pwLabel, pwInput, why,
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))));
  });
}

/** Allow Users to Edit Ranges for sheet si: resolves when closed. `run(name, ...args)` runs an editing command
 *  ('addEditRange', 'modifyEditRange', 'deleteEditRange'); `refers` is the selection's cells, for a new range;
 *  `locked`, whether the sheet is protected (then nothing can be changed). */
export function editRangesBox(root, { workbook, si, run, refers, locked }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-name-manager');
    const body = make('tbody');
    const table = make('table', { className: 'oxe-pivot-table oxe-names-table', id: `${id}-list` },
      make('thead', {}, make('tr', {}, ...['Title', 'Refers to cells'].map((h) => make('th', { scope: 'col', textContent: h })))), body);
    const caption = make('p', { className: 'oxe-dialog-text', id: `${id}-cap`, textContent: 'Ranges unlocked by a password when sheet is protected:' });
    table.setAttribute('aria-labelledby', caption.id);
    let chosen = null;
    const ranges = () => editRanges(workbook(), si);
    const fill = () => {
      const list = ranges();
      if (chosen && !list.some((x) => x.name === chosen)) chosen = null;
      if (!chosen && list[0]) chosen = list[0].name;
      body.replaceChildren(...list.map((x, k) => {
        const row = make('tr', { tabIndex: k === 0 ? 0 : -1, className: 'oxe-names-row' }, make('td', { textContent: x.name }), make('td', { textContent: refersOf(x.areas) }));
        row.setAttribute('aria-selected', String(x.name === chosen));
        const pick = () => { chosen = x.name; for (const r of body.children) r.setAttribute('aria-selected', String(r === row)); };
        row.addEventListener('click', pick);
        row.addEventListener('focus', pick);
        row.addEventListener('dblclick', () => { if (!locked) modify(); });
        row.addEventListener('keydown', (e) => {
          if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); (e.key === 'ArrowDown' ? row.nextSibling : row.previousSibling)?.focus(); }
        });
        return row;
      }));
      modifyButton.disabled = locked || !chosen;
      deleteButton.disabled = locked || !chosen;
    };
    // (A new range's title: Range1, or the first RangeN not taken, as Excel offers.)
    const nextTitle = () => { const names = new Set(ranges().map((x) => x.name.toLowerCase())); let k = 1; while (names.has(`range${k}`)) k++; return `Range${k}`; };
    async function add() {
      const answer = await rangeBox(root, { title: 'New Range', name: nextTitle(), refers, taken: (n) => ranges().some((x) => x.name.toLowerCase() === n.toLowerCase()) });
      if (answer && run('addEditRange', answer)) chosen = answer.name;
      fill();
    }
    async function modify() {
      const was = ranges().find((x) => x.name === chosen);
      if (!was) return;
      const answer = await rangeBox(root, { title: 'Modify Range', name: was.name, refers: refersOf(was.areas), modify: true,
        taken: (n) => ranges().some((x) => x.name !== was.name && x.name.toLowerCase() === n.toLowerCase()) });
      if (answer && run('modifyEditRange', was.name, answer)) chosen = answer.name;
      fill();
    }
    const newButton = button('New...', add);
    const modifyButton = button('Modify...', modify);
    const deleteButton = button('Delete', () => { if (chosen && run('deleteEditRange', chosen)) chosen = null; fill(); });
    newButton.disabled = locked;
    const close = button('OK', () => done(true));
    close.dataset.first = '';
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Allow Users to Edit Ranges' }), caption, make('div', { className: 'oxe-names-list' }, table),
      make('div', { className: 'oxe-dialog-buttons' }, newButton, modifyButton, deleteButton), make('div', { className: 'oxe-dialog-buttons' }, close));
    fill();
  });
}

/** Excel's Unlock Range: resolves to the password typed, or null. */
export function unlockRangeBox(root) {
  return passwordAsk(root, { title: 'Unlock Range', text: 'A cell you are trying to change is password protected.', label: 'Enter the password to change this cell:' });
}
