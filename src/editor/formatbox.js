// Excel's Format Cells box (Ctrl+1): its Number, Alignment, Font, Border, Fill and Protection tabs, opened on the
// active cell's format. It resolves to formatops.js formatCells's change, holding only what was changed in the box
// (a setting left alone stays as each cell has it, as in Excel), or null when cancelled.
import { modal, button } from './dialogs.js';
import { make } from './formatbox-kit.js';
import { numberTab, alignmentTab, fontTab } from './formattabs.js';
import { borderTab, fillTab, protectionTab } from './formattabs-more.js';

/** The Format Cells box. `look`: { xf, font, fontName, fontNames, fill, border, code, cell, date1904, many (more than
 *  one cell selected), part (characters selected in a cell being edited: the Font tab alone, as Excel's box has then) };
 *  `start`: the tab shown first (Number, Alignment, Font, Border, Fill, Protection). */
export function formatCellsBox(root, look, start = 'Number') {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-format-box');
    const changed = {};
    const all = [['Number', numberTab], ['Alignment', alignmentTab], ['Font', fontTab], ['Border', borderTab], ['Fill', fillTab], ['Protection', protectionTab]];
    const tabs = all.filter(([name]) => !look.part || name === 'Font')
      .map(([name, build]) => ({ name, ...build(look, (changed[name] = {})) }));
    const list = make('div', { className: 'oxe-fc-tabs', role: 'tablist' });
    const panels = tabs.map((t, k) => {
      const panel = make('div', { className: 'oxe-fc-panel', role: 'tabpanel', id: `${id}-p${k}`, hidden: t.name !== start }, t.el);
      const tab = make('button', { type: 'button', className: 'oxe-fc-tab', role: 'tab', id: `${id}-t${k}`, textContent: t.name, tabIndex: t.name === start ? 0 : -1 });
      tab.setAttribute('aria-controls', panel.id);
      tab.setAttribute('aria-selected', String(t.name === start));
      panel.setAttribute('aria-labelledby', tab.id);
      tab.addEventListener('click', () => pick(k));
      tab.addEventListener('keydown', (e) => {
        const step = e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0;
        if (step) { e.preventDefault(); pick((k + step + tabs.length) % tabs.length, true); }
      });
      list.append(tab);
      return panel;
    });
    const pick = (k, focus = false) => {
      [...list.children].forEach((tab, j) => { tab.setAttribute('aria-selected', String(j === k)); tab.tabIndex = j === k ? 0 : -1; panels[j].hidden = j !== k; });
      if (focus) list.children[k].focus();
    };
    list.children[Math.max(0, tabs.findIndex((t) => t.name === start))].dataset.first = '';
    const ok = () => {
      const change = {};
      for (const t of tabs) if (Object.keys(changed[t.name]).length) Object.assign(change, t.get());
      done(Object.keys(change).length ? change : null);
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName !== 'BUTTON' && e.target.tagName !== 'SELECT') { e.preventDefault(); ok(); } });
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Format Cells' }), list, ...panels,
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))));
  });
}
