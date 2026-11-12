// The sheet tab's menu, as Excel's right-click menu on a tab: Insert, Delete, Rename, Move Left, Move Right, Tab
// Color, Hide, Unhide. A role=menu of buttons: Up, Down, Home and End move; Right (or Enter) on Tab Color opens the
// colours (the theme's ten, ten standard ones, No Color), Left or Esc goes back; Esc in the menu, or Tab, closes it and
// focus goes back where it was. A click outside closes it.

/** Excel's standard colours, as its Tab Color palette names them. */
export const STANDARD_COLORS = [['Dark Red', 'C00000'], ['Red', 'FF0000'], ['Orange', 'FFC000'], ['Yellow', 'FFFF00'],
  ['Light Green', '92D050'], ['Green', '00B050'], ['Light Blue', '00B0F0'], ['Blue', '0070C0'], ['Dark Blue', '002060'],
  ['Purple', '7030A0']];
/** The theme's colours by their theme index, as the palette's first row names them. */
export const THEME_COLORS = ['Background 1', 'Text 1', 'Background 2', 'Text 2', 'Accent 1', 'Accent 2', 'Accent 3',
  'Accent 4', 'Accent 5', 'Accent 6'];

/** A menu in `root`. `open(entries, point, { back, byKeyboard })`: entries are { label, run, disabled, colors } or
 *  'separator'; `colors` (a list of { label, swatch: '#RRGGBB' or null, run }) makes the entry open a colour menu. */
export function createMenu(root) {
  let stack = []; // the open menus, innermost last: { el, opener }
  let back = null;

  const items = (el) => [...el.querySelectorAll(':scope > button:not([disabled]), :scope > .oxe-swatches > button')];
  const closeAll = (refocus) => {
    for (const m of stack) m.el.remove();
    stack = [];
    document.removeEventListener('mousedown', outside, true);
    if (refocus) back?.focus?.({ preventScroll: true });
  };
  const outside = (e) => { if (!stack.some((m) => m.el.contains(e.target))) closeAll(false); };

  function place(el, x, y) {
    el.style.left = '0px';
    el.style.top = '0px';
    const w = el.offsetWidth;
    const h = el.offsetHeight;
    // (Tabs are at the bottom: the menu opens upwards when there's no room below.)
    el.style.left = `${Math.max(4, Math.min(x, window.innerWidth - w - 4))}px`;
    el.style.top = `${Math.max(4, y + h > window.innerHeight - 4 ? y - h : y)}px`;
  }

  function build(entries, label, opener) {
    const el = document.createElement('div');
    el.className = 'oxe-menu';
    el.setAttribute('role', 'menu');
    el.setAttribute('aria-label', label);
    const swatches = entries.filter((e) => e !== 'separator' && e.swatch !== undefined);
    const grid = swatches.length ? document.createElement('div') : null;
    if (grid) grid.className = 'oxe-swatches';
    for (const entry of entries) {
      if (entry === 'separator') { const hr = document.createElement('hr'); hr.setAttribute('role', 'separator'); el.append(hr); continue; }
      // (A heading, as Excel's error menu has: its name in bold, not an item.)
      if (entry.heading) { const h = document.createElement('div'); h.className = 'oxe-menuhead'; h.setAttribute('role', 'presentation'); h.textContent = entry.label; el.append(h); continue; }
      const b = document.createElement('button');
      b.type = 'button';
      b.setAttribute('role', 'menuitem');
      b.disabled = !!entry.disabled;
      if (entry.swatch !== undefined) {
        b.className = 'oxe-swatch';
        b.setAttribute('aria-label', entry.label);
        b.title = entry.label;
        if (entry.swatch) b.style.background = entry.swatch;
        else b.textContent = entry.label;
        if (!entry.swatch) { b.classList.add('oxe-swatch-none'); el.append(b); } else grid.append(b);
      } else {
        b.className = 'oxe-menuitem';
        b.textContent = entry.label;
        el.append(b);
      }
      if (entry.colors) {
        b.setAttribute('aria-haspopup', 'menu');
        b.setAttribute('aria-expanded', 'false');
        b.classList.add('oxe-menuitem-sub');
      }
      b.addEventListener('click', () => {
        if (entry.colors) { openSub(b, entry); return; }
        closeAll(true);
        entry.run();
      });
    }
    if (grid) el.insertBefore(grid, el.firstChild);
    el.addEventListener('keydown', (e) => {
      const list = items(el);
      const i = list.indexOf(document.activeElement);
      const go = (k) => { e.preventDefault(); list[(k + list.length) % list.length]?.focus(); };
      if (e.key === 'ArrowDown') go(i + 1);
      else if (e.key === 'ArrowUp') go(i - 1);
      else if (e.key === 'Home') go(0);
      else if (e.key === 'End') go(list.length - 1);
      else if (e.key === 'ArrowRight' && document.activeElement?.getAttribute('aria-haspopup')) { e.preventDefault(); document.activeElement.click(); }
      else if ((e.key === 'ArrowLeft' || e.key === 'Escape') && opener) { e.preventDefault(); e.stopPropagation(); closeSub(); }
      else if (e.key === 'Escape') { e.preventDefault(); closeAll(true); }
      else if (e.key === 'Tab') { e.preventDefault(); closeAll(true); }
    });
    return el;
  }

  function openSub(opener, entry) {
    if (stack.length > 1) closeSub();
    const el = build(entry.colors, entry.label, opener);
    root.append(el);
    stack.push({ el, opener });
    opener.setAttribute('aria-expanded', 'true');
    const r = opener.getBoundingClientRect();
    place(el, r.right, r.top);
    items(el)[0]?.focus();
  }
  function closeSub() {
    const m = stack.pop();
    if (!m) return;
    m.el.remove();
    m.opener.setAttribute('aria-expanded', 'false');
    m.opener.focus();
  }

  return {
    open(entries, { x, y }, { label = 'Menu', byKeyboard = false } = {}) {
      closeAll(false);
      back = document.activeElement;
      const el = build(entries, label, null);
      root.append(el);
      stack = [{ el, opener: null }];
      place(el, x, byKeyboard ? y - 2 : y);
      items(el)[0]?.focus();
      document.addEventListener('mousedown', outside, true);
    },
    close: () => closeAll(false),
    get isOpen() { return stack.length > 0; },
  };
}

/** The entries of sheet i's tab menu. `act` = { run(command, ...args), rename(i), unhide(), remove(i) (asks first) };
 *  `color({ theme })` gives a theme colour's swatch. Read-only, every entry is there but off. */
export function tabMenuEntries(wb, i, { act, color, readOnly }) {
  const shown = wb.sheets.map((_, k) => k).filter((k) => wb.sheets[k].state === 'visible');
  const at = shown.indexOf(i);
  const colors = [
    ...THEME_COLORS.map((label, theme) => ({ label, swatch: color({ theme }), run: () => act.run('tabColor', i, { theme }) })),
    ...STANDARD_COLORS.map(([label, hex]) => ({ label, swatch: `#${hex}`, run: () => act.run('tabColor', i, { rgb: `FF${hex}` }) })),
    { label: 'No Color', swatch: null, run: () => act.run('tabColor', i, null) },
  ];
  return [
    { label: 'Insert', disabled: readOnly, run: () => act.run('insertSheet') },
    { label: 'Delete', disabled: readOnly, run: () => act.remove(i) },
    { label: 'Rename', disabled: readOnly, run: () => act.rename(i) },
    { label: 'Move Left', disabled: readOnly || at <= 0, run: () => act.run('moveSheet', i, shown[at - 1]) },
    { label: 'Move Right', disabled: readOnly || at < 0 || at >= shown.length - 1, run: () => act.run('moveSheet', i, shown[at + 1]) },
    { label: 'Move or Copy...', disabled: readOnly || wb.sheets[i].kind !== 'worksheet', run: () => act.moveCopy?.(i) },
    { label: 'Tab Color', disabled: readOnly || wb.sheets[i].kind !== 'worksheet', colors },
    'separator',
    { label: 'Hide', disabled: readOnly, run: () => act.run('hideSheet', i) },
    { label: 'Unhide...', disabled: readOnly || !wb.sheets.some((s) => s.state === 'hidden'), run: () => act.unhide() },
  ];
}
