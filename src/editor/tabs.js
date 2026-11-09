// The sheet tabs under the grid, as Excel shows them: one for each sheet whose tab isn't hidden, in the workbook's
// order, in the sheet's tab colour, the sheet shown marked, and a New sheet button after them. An ARIA tablist: Left,
// Right, Home and End move between the tabs and show the sheet (as a click does); a chart sheet's tab is there but
// can't be shown yet. Ctrl+Page Up and Page Down in the grid change sheet from there (commands.js). A right click,
// Shift+F10 or the Menu key on a tab opens its menu (sheetmenu.js); a double click renames it in place; dragging a tab
// moves the sheet.
import { div } from './grid.js';

/** Black or white, whichever reads better on `hex` ('#RRGGBB'; WCAG relative luminance). */
export function textOn(hex) {
  const lin = (i) => {
    const v = parseInt(hex.slice(i, i + 2), 16) / 255;
    return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
  };
  const l = 0.2126 * lin(1) + 0.7152 * lin(3) + 0.0722 * lin(5);
  return (l + 0.05) / 0.05 >= 1.05 / (l + 0.05) ? '#000000' : '#ffffff';
}

/** A strip at the bottom of `root`. `controls` is the id of the grid it shows; `on` = { select(i), menu(i, point,
 *  byKeyboard), rename(i), move(from, to), add() }. */
export function createSheetTabs(root, { controls, on }) {
  const bar = div('oxe-tabbar');
  const strip = div('oxe-tabs', 'tablist');
  strip.setAttribute('aria-label', 'Sheets');
  const add = document.createElement('button');
  add.type = 'button';
  add.className = 'oxe-tab-add';
  add.textContent = '+';
  add.title = 'New sheet (Shift+F11)';
  add.setAttribute('aria-label', 'New sheet');
  add.addEventListener('click', () => on.add());
  bar.append(strip, add);
  root.append(bar);
  let tabs = []; // [{ index, el }]
  let dragged = false;

  const mark = (i) => {
    for (const t of tabs) {
      const sel = t.index === i;
      t.el.setAttribute('aria-selected', String(sel));
      t.el.tabIndex = sel ? 0 : -1;
      if (sel) t.el.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
    }
  };
  const tabAt = (target) => target.closest?.('.oxe-tab');
  strip.addEventListener('click', (e) => {
    const el = tabAt(e.target);
    if (el && !dragged) on.select(Number(el.dataset.sheet));
    dragged = false;
  });
  strip.addEventListener('dblclick', (e) => { const el = tabAt(e.target); if (el) on.rename(Number(el.dataset.sheet)); });
  strip.addEventListener('contextmenu', (e) => {
    const el = tabAt(e.target);
    if (!el) return;
    e.preventDefault();
    // (From the keyboard the event comes at 0, 0 where there is no pointer: Shift+F10 is handled below instead.)
    if (e.button !== 2 && e.clientX === 0 && e.clientY === 0) return;
    on.menu(Number(el.dataset.sheet), { x: e.clientX, y: e.clientY }, false);
  });
  strip.addEventListener('keydown', (e) => {
    const at = tabs.findIndex((t) => t.el === document.activeElement);
    if (at < 0) return;
    if ((e.key === 'F10' && e.shiftKey) || e.key === 'ContextMenu') {
      e.preventDefault();
      const r = tabs[at].el.getBoundingClientRect();
      on.menu(tabs[at].index, { x: r.left, y: r.top }, true);
      return;
    }
    const to = { ArrowLeft: at - 1, ArrowRight: at + 1, Home: 0, End: tabs.length - 1 }[e.key];
    if (to == null) return;
    e.preventDefault();
    const t = tabs[Math.max(0, Math.min(tabs.length - 1, to))];
    t.el.focus();
    on.select(t.index);
  });

  // Dragging a tab: past 5 pixels it moves; the gap it would go into is marked; let go, the sheet moves there.
  let drag = null;
  const dropAt = (x) => {
    for (const t of tabs) {
      const r = t.el.getBoundingClientRect();
      if (x < r.left + r.width / 2) return { before: t };
    }
    return { after: tabs[tabs.length - 1] };
  };
  const unmark = () => strip.querySelectorAll('.oxe-drop-before, .oxe-drop-after').forEach((el) => el.classList.remove('oxe-drop-before', 'oxe-drop-after'));
  // (A formula being typed keeps focus when a tab is pressed: it goes on, pointing at the other sheet.)
  strip.addEventListener('mousedown', (e) => { if (e.button === 0 && on.keepFocus?.()) e.preventDefault(); });
  strip.addEventListener('pointerdown', (e) => {
    const el = tabAt(e.target);
    dragged = false; // (a drag's click may never come: each press starts afresh)
    if (e.button !== 0 || !el) return;
    drag = { from: Number(el.dataset.sheet), x: e.clientX, moved: false, id: e.pointerId };
  });
  strip.addEventListener('pointermove', (e) => {
    if (!drag || (!drag.moved && Math.abs(e.clientX - drag.x) < 5)) return;
    if (!drag.moved) { drag.moved = true; strip.setPointerCapture?.(drag.id); }
    unmark();
    const d = dropAt(e.clientX);
    if (d.before) d.before.el.classList.add('oxe-drop-before');
    else d.after.el.classList.add('oxe-drop-after');
  });
  const endDrag = (e) => {
    if (!drag) return;
    const { from, moved } = drag;
    drag = null;
    unmark();
    if (!moved || e.type === 'pointercancel') return;
    dragged = true;
    const d = dropAt(e.clientX);
    const k = d.before ? d.before.index : d.after.index;
    const to = d.before ? (k > from ? k - 1 : k) : (k > from ? k : k + 1);
    if (to !== from) on.move(from, to);
  };
  strip.addEventListener('pointerup', endDrag);
  strip.addEventListener('pointercancel', endDrag);

  return {
    el: bar,
    /** Show a workbook's tabs, [{ index, name, color ('#RRGGBB' or null), disabled }], with sheet `active` marked.
     *  A tab that had focus keeps it on the active tab. */
    set(list, active) {
      const focused = strip.contains(document.activeElement);
      tabs = list.map(({ index, name, color, disabled }) => {
        const el = document.createElement('button');
        el.type = 'button';
        el.className = 'oxe-tab';
        el.setAttribute('role', 'tab');
        el.setAttribute('aria-controls', controls);
        el.dataset.sheet = String(index);
        el.textContent = name;
        if (color) {
          el.classList.add('oxe-tab-colored');
          el.style.setProperty('--oxe-tab-color', color);
          el.style.setProperty('--oxe-tab-text', textOn(color));
        }
        if (disabled) {
          el.setAttribute('aria-disabled', 'true');
          el.title = "Chart sheets can't be shown yet";
        }
        return { index, el };
      });
      strip.replaceChildren(...tabs.map((t) => t.el));
      mark(active);
      if (focused) tabs.find((t) => t.index === active)?.el.focus();
    },
    /** Mark sheet i as the one shown. */
    activate: mark,
    /** Focus sheet i's tab. */
    focus(i) { tabs.find((t) => t.index === i)?.el.focus(); },
    /** Rename sheet i in place: a text box over its tab. Enter or leaving it commits through `commit(name)`, which
     *  resolves false when the name was refused (the box stays, its text selected, as in Excel); Esc gives up. */
    editName(i, commit) {
      const t = tabs.find((x) => x.index === i);
      if (!t) return;
      const was = t.el.textContent;
      const box = document.createElement('input');
      box.className = 'oxe-tab-edit';
      box.value = was;
      box.spellcheck = false;
      box.setAttribute('aria-label', 'Sheet name');
      box.size = Math.max(6, was.length + 2);
      t.el.hidden = true;
      t.el.after(box);
      let busy = false;
      const finish = async (keep) => {
        if (busy) return;
        busy = true;
        if (keep && box.value !== was && !(await commit(box.value))) {
          busy = false;
          box.focus();
          box.select();
          return;
        }
        box.remove();
        t.el.hidden = false;
        if (!keep || box.value === was) t.el.focus();
      };
      box.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') { e.preventDefault(); finish(true); }
        else if (e.key === 'Escape') { e.preventDefault(); finish(false); }
      });
      box.addEventListener('blur', () => finish(true));
      box.focus();
      box.select();
    },
  };
}
