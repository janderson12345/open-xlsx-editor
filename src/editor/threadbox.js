// Excel's comment thread (Review > New Comment, Ctrl+Alt+M; a thread's Reply, Edit and Delete): a cell's thread, each
// comment with its author and time, Edit for each and Delete for a reply, a box to start the thread or reply to it, and
// Resolve thread (Reopen once resolved), Delete Thread. Resolves to what was asked: { post: text }, { edit: k, text },
// { remove: k }, { resolve: true | false }, { removeAll: true }, or
// null; the editor makes it (threadedit.js) and opens the box again on the thread.
import { modal, button } from './dialogs.js';

const make = (tag, props = {}, ...kids) => { const e = Object.assign(document.createElement(tag), props); e.append(...kids.filter((k) => k != null)); return e; };
/** A comment's time as the thread shows it (its dT, local). */
const when = (dT) => { const d = new Date(dT); return Number.isNaN(d.getTime()) ? '' : d.toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' }); };

/** The thread box for cell `ref`: `items` the thread's comments ([{ author, text, date }], none for a new one). */
export function threadBox(root, { ref, items = [], at, readOnly = false, resolved = false }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-note-edit', 'oxe-thread-box');
    dlg.style.setProperty('--oxe-fx', `${Math.round(at.x)}px`);
    dlg.style.setProperty('--oxe-fy', `${Math.round(at.y)}px`);
    const title = make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: items.length ? `Comment thread on ${ref}${resolved ? ' (Resolved)' : ''}` : `New comment on ${ref}` });
    const list = make('ol', { className: 'oxe-thread-list' });
    items.forEach((x, k) => {
      const head = make('p', { className: 'oxe-thread-head' }, make('strong', { textContent: x.author || 'Unknown' }), x.date ? ` ${when(x.date)}` : '');
      const body = make('p', { className: 'oxe-thread-text', textContent: x.text });
      const acts = make('div', { className: 'oxe-thread-acts' });
      if (!readOnly && !resolved) {
        const edit = button('Edit', () => {
          // (Edited in place: the text in a box of its own, Save or Cancel.)
          const area = make('textarea', { className: 'oxe-note-text', value: x.text, rows: 3 });
          area.setAttribute('aria-label', `Edit ${k ? `reply ${k}` : 'the comment'}`);
          const save = () => done({ edit: k, text: area.value });
          area.addEventListener('keydown', (e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); save(); } });
          body.replaceWith(area);
          acts.replaceChildren(button('Save', save), button('Cancel', () => { area.replaceWith(body); acts.replaceChildren(...own); }));
          area.focus();
        });
        edit.setAttribute('aria-label', `Edit ${k ? `reply ${k}` : 'the comment'} by ${x.author || 'Unknown'}`);
        const own = [edit];
        if (k) {
          const del = button('Delete', () => done({ remove: k }));
          del.setAttribute('aria-label', `Delete reply ${k} by ${x.author || 'Unknown'}`);
          own.push(del);
        }
        acts.append(...own);
      }
      list.append(make('li', { className: 'oxe-thread-item' }, head, body, acts));
    });
    const area = make('textarea', { className: 'oxe-note-text', rows: 3, spellcheck: true, placeholder: items.length ? 'Reply…' : 'Start a conversation' });
    area.setAttribute('aria-label', items.length ? 'Reply' : 'Comment');
    area.dataset.first = '';
    const post = () => { if (area.value.trim()) done({ post: area.value }); };
    area.addEventListener('keydown', (e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); post(); } });
    const row = make('div', { className: 'oxe-dialog-buttons' });
    // (A resolved thread: Reopen, no reply, as Excel's; an open one: Resolve thread.)
    if (!readOnly && !resolved) row.append(button(items.length ? 'Reply' : 'Post', post));
    if (!readOnly && items.length) row.append(resolved ? button('Reopen', () => done({ resolve: false })) : button('Resolve thread', () => done({ resolve: true })));
    if (!readOnly && items.length) row.append(button('Delete Thread', () => done({ removeAll: true })));
    row.append(button('Close', () => done(null)));
    dlg.append(...[title, items.length ? list : null, readOnly || resolved ? null : area, row].filter(Boolean));
    if (readOnly || resolved) row.firstChild.dataset.first = '';
  });
}
