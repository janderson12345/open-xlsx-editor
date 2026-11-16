// Data > Queries & Connections > Edit Links: Excel's Edit Links box, the workbook's links to other workbooks
// (ooxml/extlinks.js) with their source, type, update and status; Update Values (from the workbook when it is open in
// an editor on the page, else from a file chosen), Change Source... (another file: its name, and its values), Break
// Link (the formulas that use it become their values, after Excel's question) and Close. Each change is one step of
// the history (`run`, the sheetChange command).
import { modal, button, messageBox } from './dialogs.js';
import { readLinks, refreshLink, updateLink, breakLinkSteps } from '../ooxml/extlinks.js';
import { Workbook } from '../ooxml/workbook.js';
import { openBook } from './pagebooks.js';

const make = (tag, props = {}, ...kids) => { const e = Object.assign(document.createElement(tag), props); e.append(...kids); return e; };
/** Excel's question before Break Link. */
export const BREAK_QUESTION = 'Breaking links permanently converts formulas and external references to their existing values. Because this cannot be undone, you may want to save a version of this file with a new name.\n\nAre you sure you want to break the links?';

/** A workbook file chosen with the browser's file picker: { name, wb }, or null (none chosen, or not one it reads). */
function chooseFile(root) {
  return new Promise((resolve) => {
    const input = make('input', { type: 'file', accept: '.xlsx,.xlsm,.xltx,.xltm', className: 'oxe-links-file' });
    input.hidden = true;
    input.addEventListener('change', async () => {
      const file = input.files?.[0];
      input.remove();
      if (!file) { resolve(null); return; }
      try { resolve({ name: file.name, wb: Workbook.fromBytes(new Uint8Array(await file.arrayBuffer())) }); } catch { resolve(null); }
    });
    input.addEventListener('cancel', () => { input.remove(); resolve(null); });
    root.append(input);
    input.click();
  });
}

/** Excel's Edit Links box over the workbook's links; resolves when closed. `run(fn)` makes a change (fn(wb)) as one
 *  step of the history; `announce` says what happened. */
export function editLinksBox(root, { workbook, run, announce = () => {} }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-links-box');
    const body = make('tbody');
    const table = make('table', { className: 'oxe-pivot-table oxe-links-table' },
      make('thead', {}, make('tr', {}, ...['Source', 'Type', 'Update', 'Status'].map((h) => make('th', { scope: 'col', textContent: h })))), body);
    const where = make('p', { className: 'oxe-dialog-text oxe-links-where' });
    let chosen = 0;
    const fill = () => {
      const links = readLinks(workbook());
      if (chosen >= links.length) chosen = links.length - 1;
      body.replaceChildren(...links.map((l, k) => {
        // (OK when the workbook is open in an editor here, as Excel says for an open source; else Unknown, as before
        // Check Status.)
        const status = openBook(l.name) ? 'OK' : 'Unknown';
        const row = make('tr', { tabIndex: k === Math.max(chosen, 0) ? 0 : -1, className: 'oxe-links-row' }, ...[l.name || l.target || '', 'Worksheet', 'A', status].map((t) => make('td', { textContent: t })));
        row.setAttribute('aria-selected', String(k === chosen));
        const pick = () => { chosen = k; for (const r of body.children) r.setAttribute('aria-selected', String(r === row)); where.textContent = `Location: ${l.target || ''}`; };
        row.addEventListener('click', pick);
        row.addEventListener('focus', pick);
        row.addEventListener('keydown', (e) => {
          const step = e.key === 'ArrowDown' ? 1 : e.key === 'ArrowUp' ? -1 : 0;
          if (!step) return;
          e.preventDefault();
          (body.children[k + step] || row).focus();
        });
        return row;
      }));
      const link = links[chosen];
      where.textContent = link ? `Location: ${link.target || ''}` : 'This workbook has no links to other workbooks.';
      for (const b of [update, change, breaks]) b.disabled = !link;
    };
    const current = () => readLinks(workbook())[chosen] || null;
    /** The source of `link`: the workbook open here by its name, else one chosen from a file. */
    const sourceOf = async (link) => {
      const open = openBook(link.name);
      if (open) return { name: link.name, wb: open };
      announce(`${link.name} isn't open here: choose the file to update from`);
      return chooseFile(root);
    };
    const update = button('Update Values', async () => {
      const link = current();
      if (!link) return;
      const src = await sourceOf(link);
      if (!src) return;
      run((wb) => { refreshLink(wb, link.index, src.wb); });
      announce(`Updated the values from ${link.name}`);
      fill();
    });
    const change = button('Change Source...', async () => {
      const link = current();
      if (!link) return;
      const src = await chooseFile(root);
      if (!src) return;
      run((wb) => { updateLink(wb, link.index, { target: src.name }); refreshLink(wb, link.index, src.wb); });
      announce(`The link now goes to ${src.name}`);
      fill();
    });
    const breaks = button('Break Link', async () => {
      const link = current();
      if (!link) return;
      if (await messageBox(root, BREAK_QUESTION, ['Break Links', 'Cancel']) !== 'Break Links') return;
      run((wb) => { for (const st of breakLinkSteps(wb, link.index)) wb.apply(st.si, st.contents); });
      announce(`Broke the link to ${link.name}`);
      fill();
    });
    const close = button('Close', () => done(true));
    close.dataset.first = '';
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Edit Links' }), make('div', { className: 'oxe-names-list' }, table), where,
      make('div', { className: 'oxe-dialog-buttons oxe-links-buttons' }, update, change, breaks),
      make('div', { className: 'oxe-dialog-buttons' }, close));
    fill();
  });
}
