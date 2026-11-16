// Insert > Link (Ctrl+K): Excel's Insert Hyperlink box, and Edit Hyperlink on a cell with one. Link to: Existing File or
// Web Page (Address, and Browse for File: a file picked, its name the address, as a link relative to the workbook),
// Place in This Document (Type the cell reference, the sheet or defined name), Create New Document (Name of new
// document, When to edit), E-mail Address (E-mail address, Subject); Text to display, and ScreenTip... (Set Hyperlink
// ScreenTip). Edit Hyperlink adds Remove Link. ooxml/links.js writes what it answers. A page can't read where a file is
// or make one: Browse gives only the file's name (nothing of the file is read), and a new document is the host's to
// make (index.js `onCreateDocument`).
import { modal, button } from './dialogs.js';
import { sheetPrefix } from '../calc/tokens.js';
import { parseRange } from '../ooxml/refs.js';

const make = (tag, props = {}, ...kids) => { const e = Object.assign(document.createElement(tag), props); e.append(...kids); return e; };
const SCHEME = /^[a-z][a-z0-9+.-]*:/i;

/** What the Address field holds, as the hyperlink's address: one starting www. gets http:// ("www.example.com" to
 *  "http://www.example.com"); anything else stays as typed, a file's name or path among them (report.xlsx). */
export function addressOf(text) {
  const t = String(text || '').trim();
  return /^www\./i.test(t) && !SCHEME.test(t) ? `http://${t}` : t;
}

/** Create New Document's name as the hyperlink's address: as typed, trimmed, with .xlsx when it has no extension (a
 *  new workbook, as Excel's box makes one); null when it can't be a file's name (empty, or with \ / : * ? " < > |). */
export function newDocumentName(text) {
  const t = String(text || '').trim();
  if (!t || /[\\/:*?"<>|]/.test(t) || /^\.+$/.test(t)) return null;
  return /\.[a-z0-9]{1,5}$/i.test(t) ? t : `${t}.xlsx`;
}

/** An e-mail hyperlink's address: mailto:, with ?subject= when there is one (encoded as a URL's query is). */
export const mailtoOf = (email, subject) => `mailto:${String(email).trim().replace(/^mailto:/i, '')}${subject ? `?subject=${encodeURIComponent(subject)}` : ''}`;

/** A place in the workbook as the hyperlink's location: the sheet (quoted as a formula quotes it) and the cell
 *  reference, or a defined name as it is. */
export const locationOf = (sheet, ref) => (sheet == null ? ref : `${sheetPrefix(sheet)}${ref}`);

/** Set Hyperlink ScreenTip: resolves to the text, or null. */
function screenTipBox(root, text) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    const field = make('input', { id: `${id}-tip`, className: 'oxe-dialog-field', type: 'text', value: text, spellcheck: true });
    field.dataset.first = '';
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON') { e.preventDefault(); done(field.value); } });
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Set Hyperlink ScreenTip' }),
      make('label', { className: 'oxe-dialog-text oxe-dialog-label', htmlFor: field.id, textContent: 'ScreenTip text:' }), field,
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', () => done(field.value)), button('Cancel', () => done(null))));
  });
}

/**
 * The box. `link`: the cell's hyperlink (links.js readHyperlinks' shape) for Edit Hyperlink, or null; `text`: the
 * cell's text, or null when it holds something else (a number, a formula: Text to display then can't be changed, as in
 * Excel); `sheets`: the sheet names; `names`: the defined names; `sheet`: the one shown. Resolves to { address,
 * location, tooltip, text } (text null: left as it is; with `create`: { name, editNow } for Create New Document),
 * { remove: true } (Remove Link), or null.
 */
export function linkBox(root, { link = null, text = null, sheets, names = [], sheet }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-link-box');
    let tooltip = link?.tooltip || '';
    const field = (key, label, value = '') => {
      const input = make('input', { id: `${id}-${key}`, className: 'oxe-dialog-field', type: 'text', value, spellcheck: false });
      return [make('label', { className: 'oxe-dialog-text oxe-dialog-label', htmlFor: input.id, textContent: label }), input];
    };
    const [showLabel, show] = field('show', 'Text to display:', text ?? '<<Selection in Document>>');
    show.disabled = text == null;
    // (Link to: as Excel's, the kind the cell's link is.)
    const mail = link?.address && /^mailto:/i.test(link.address);
    const kind = link?.location && !link.address ? 'place' : mail ? 'mail' : 'web';
    const kinds = [['web', 'Existing File or Web Page'], ['place', 'Place in This Document'], ['new', 'Create New Document'], ['mail', 'E-mail Address']];
    const set = make('fieldset', { className: 'oxe-dialog-group' }, make('legend', { textContent: 'Link to:' }));
    const radios = kinds.map(([k, label]) => {
      const input = make('input', { type: 'radio', name: `${id}-kind`, id: `${id}-k-${k}`, value: k, checked: k === kind });
      set.append(make('div', { className: 'oxe-dialog-option' }, input, make('label', { htmlFor: input.id, textContent: label })));
      return input;
    });
    // Existing File or Web Page
    const [addrLabel, addr] = field('addr', 'Address:', mail ? '' : link?.address || '');
    // (Browse for File: the picker's file, by its name alone, as a link relative to the workbook; the text to display
    // too when it was empty, as Excel fills it.)
    const picker = make('input', { type: 'file', hidden: true, tabIndex: -1 });
    picker.setAttribute('aria-hidden', 'true');
    picker.addEventListener('change', () => {
      const f = picker.files?.[0];
      if (!f) return;
      addr.value = f.name;
      if (!show.disabled && (!show.value || show.value === text)) show.value = f.name;
      picker.value = '';
      addr.focus();
    });
    const browse = button('Browse for File...', () => picker.click());
    const web = make('div', { className: 'oxe-link-panel' }, addrLabel, addr, make('div', { className: 'oxe-dialog-buttons' }, browse), picker);
    // Place in This Document: the cell reference, and the sheet or defined name it is on.
    const [placed, at] = (() => {
      const loc = link?.location || '';
      const bang = loc.lastIndexOf('!');
      if (bang < 0) return [names.includes(loc) ? loc : sheet, names.includes(loc) ? '' : loc || 'A1'];
      return [loc.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'"), loc.slice(bang + 1)];
    })();
    const [refLabel, ref] = field('ref', 'Type the cell reference:', at || 'A1');
    const where = make('select', { id: `${id}-where`, className: 'oxe-dialog-select', size: Math.min(8, sheets.length + names.length + 2) });
    const group = (label, items) => make('optgroup', { label }, ...items.map((v) => make('option', { value: `${label === 'Defined Names' ? 'n' : 's'}:${v}`, textContent: v, selected: v === placed })));
    where.append(group('Cell Reference', sheets), ...(names.length ? [group('Defined Names', names)] : []));
    const whereLabel = make('label', { className: 'oxe-dialog-text oxe-dialog-label', htmlFor: where.id, textContent: 'Or select a place in this document:' });
    const place = make('div', { className: 'oxe-link-panel' }, refLabel, ref, whereLabel, where);
    const isName = () => where.value.startsWith('n:');
    where.addEventListener('change', () => { ref.disabled = isName(); });
    ref.disabled = isName();
    // Create New Document: its name, where it goes (the workbook's folder: a page can't know which that is) and when to
    // edit it, as Excel's box has them.
    const [nameLabel, docName] = field('docname', 'Name of new document:', '');
    const fullPath = make('p', { className: 'oxe-dialog-text', textContent: "Full path: the workbook's folder" });
    const when = make('fieldset', { className: 'oxe-dialog-group' }, make('legend', { textContent: 'When to edit:' }));
    const later = make('input', { type: 'radio', name: `${id}-when`, id: `${id}-later`, value: 'later' });
    const now = make('input', { type: 'radio', name: `${id}-when`, id: `${id}-now`, value: 'now', checked: true });
    when.append(make('div', { className: 'oxe-dialog-option' }, later, make('label', { htmlFor: later.id, textContent: 'Edit the new document later' })),
      make('div', { className: 'oxe-dialog-option' }, now, make('label', { htmlFor: now.id, textContent: 'Edit the new document now' })));
    const newPanel = make('div', { className: 'oxe-link-panel' }, nameLabel, docName, fullPath, when);
    // E-mail Address
    const [mailLabel, email] = field('mail', 'E-mail address:', mail ? decodeURIComponent(link.address.replace(/^mailto:/i, '').split('?')[0]) : '');
    const [subjLabel, subject] = field('subj', 'Subject:', mail ? decodeURIComponent((/[?&]subject=([^&]*)/i.exec(link.address) || [])[1] || '') : '');
    const mailPanel = make('div', { className: 'oxe-link-panel' }, mailLabel, email, subjLabel, subject);
    const panels = { web, place, new: newPanel, mail: mailPanel };
    const chosen = () => radios.find((r) => r.checked).value;
    const shown = () => { for (const [k, p] of Object.entries(panels)) p.hidden = k !== chosen(); };
    for (const r of radios) r.addEventListener('change', shown);
    shown();
    const why = make('p', { className: 'oxe-dialog-text oxe-pivot-why' });
    why.setAttribute('role', 'alert');
    const ok = () => {
      const k = chosen();
      const answer = { address: '', location: '', tooltip, text: show.disabled ? null : show.value };
      if (k === 'web') {
        answer.address = addressOf(addr.value);
        if (!answer.address) { why.textContent = 'Type an address.'; return; }
      } else if (k === 'new') {
        const n = newDocumentName(docName.value);
        if (!n) { why.textContent = docName.value.trim() ? 'The file name is not valid.' : 'Type a name for the new document.'; return; }
        answer.address = n;
        answer.create = { name: n, editNow: now.checked };
      } else if (k === 'mail') {
        if (!email.value.trim()) { why.textContent = 'Type an e-mail address.'; return; }
        answer.address = mailtoOf(email.value, subject.value);
      } else if (isName()) answer.location = where.value.slice(2);
      else {
        const target = ref.value.trim().replace(/\$/g, '');
        // (Excel's words for a reference that isn't one.)
        if (!parseRange(target)) { why.textContent = 'You must type a valid cell reference or select a location in this document.'; return; }
        answer.location = locationOf(where.value.slice(2) || sheet, target.toUpperCase());
      }
      if (answer.text === '') answer.text = null;
      done(answer);
    };
    const tipButton = button('ScreenTip...', async () => { const t = await screenTipBox(root, tooltip); if (t != null) tooltip = t; });
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON' && e.target.tagName !== 'SELECT') { e.preventDefault(); ok(); } });
    (text == null ? addr : show).dataset.first = '';
    const buttons = make('div', { className: 'oxe-dialog-buttons' });
    if (link) buttons.append(button('Remove Link', () => done({ remove: true })));
    buttons.append(button('OK', ok), button('Cancel', () => done(null)));
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: link ? 'Edit Hyperlink' : 'Insert Hyperlink' }),
      set, showLabel, show, make('div', { className: 'oxe-dialog-buttons' }, tipButton), web, place, newPanel, mailPanel, why, buttons);
  });
}
