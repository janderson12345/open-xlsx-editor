// The sheet tabs under the grid, as Excel shows them: one for each sheet whose tab isn't hidden, in the workbook's
// order, in the sheet's tab colour, the sheet shown marked. An ARIA tablist: Left, Right, Home and End move between the
// tabs and show the sheet (as a click does); a chart sheet's tab is there but can't be shown yet. Ctrl+Page Up and Page
// Down in the grid change sheet from there (commands.js).
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

/** A strip at the bottom of `root`. `onSelect(i)` when a tab is chosen; `controls` the id of the grid it shows. */
export function createSheetTabs(root, { onSelect, controls }) {
  const strip = div('oxe-tabs', 'tablist');
  strip.setAttribute('aria-label', 'Sheets');
  root.append(strip);
  let tabs = []; // [{ index, el }]

  const mark = (i) => {
    for (const t of tabs) {
      const on = t.index === i;
      t.el.setAttribute('aria-selected', String(on));
      t.el.tabIndex = on ? 0 : -1;
      if (on) t.el.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
    }
  };
  strip.addEventListener('click', (e) => {
    const el = e.target.closest('.oxe-tab');
    if (el) onSelect(Number(el.dataset.sheet));
  });
  strip.addEventListener('keydown', (e) => {
    const at = tabs.findIndex((t) => t.el === document.activeElement);
    if (at < 0) return;
    const to = { ArrowLeft: at - 1, ArrowRight: at + 1, Home: 0, End: tabs.length - 1 }[e.key];
    if (to == null) return;
    e.preventDefault();
    const t = tabs[Math.max(0, Math.min(tabs.length - 1, to))];
    t.el.focus();
    onSelect(t.index);
  });

  return {
    el: strip,
    /** Show a workbook's tabs, [{ index, name, color ('#RRGGBB' or null), disabled }], with sheet `active` marked. */
    set(list, active) {
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
    },
    /** Mark sheet i as the one shown. */
    activate: mark,
  };
}
