// The ribbon's building blocks: elements, buttons, groups, the one drop-down menu, select fields. From open-docx-editor
// (its ribbonkit.js), with commands as plain functions and the sheet, not a document, getting focus back.

// Line icons drawn in the text colour, so they show in light and dark mode (colour emoji can be near-black).
const ICON_PATHS = {
  '↶': 'M9 14L4 9l5-5 M4 9h11a5 5 0 0 1 0 10h-3',
  '↷': 'M15 14l5-5-5-5 M20 9H9a5 5 0 0 0 0 10h3',
  '🪣': 'M5 11l7-7 7 7-7 7z M5 11h14 M19 15c0 1.5 1 2.5 1 3.5a1 1 0 0 1-2 0c0-1 1-2 1-3.5z',
  '⊞': 'M4 4h16v16H4z M4 12h16 M12 4v16',
  '⏎': 'M4 7h12a4 4 0 0 1 0 8H8 M11 12l-3 3 3 3 M4 11h6',
  '⇤': 'M4 5v14 M20 7H9 M20 12H12 M20 17H9 M12 9l-3 3 3 3',
  '⇥': 'M4 5v14 M20 7H9 M20 12H12 M20 17H9 M9 9l3 3-3 3',
  '🗑': 'M4 7h16 M9 7V4h6v3 M6 7l1 13h10l1-13',
  '⊕': 'M12 5v14 M5 12h14',
  '✎': 'M4 20h4L19 9l-4-4L4 16z M14 6l4 4',
  '⌫': 'M9 5h11v14H9l-6-7z M12 9l5 6 M17 9l-5 6',
  'A⁺': 'M3 20l6-15 6 15 M5.5 14h7 M18 9V3 M15 6h6',
  'A⁻': 'M3 20l6-15 6 15 M5.5 14h7 M15 6h6',
  '.0←': 'M3 19h.01 M7 13v6 M11 13v6 M15 13v6 M20 6H9 M12 3L9 6l3 3',
  '.0→': 'M3 19h.01 M7 13v6 M11 13v6 M9 6h11 M17 3l3 3-3 3',
  '⊟': 'M3 5h18v14H3z M6 12h12 M9 9l-3 3 3 3 M15 9l3 3-3 3',
  '⫷': 'M4 6h16 M4 10h10 M4 14h16 M4 18h10',
  '≡': 'M4 6h16 M7 10h10 M4 14h16 M7 18h10',
  '⫸': 'M4 6h16 M10 10h10 M4 14h16 M10 18h10',
  '⤒': 'M5 4h14 M12 20V8 M8 12l4-4 4 4',
  '↕': 'M5 12h14 M12 3v6 M9 6l3 3 3-3 M12 21v-6 M9 18l3-3 3 3',
  '⤓': 'M5 20h14 M12 4v12 M8 12l4 4 4-4',
};
export function iconize(root) {
  root.querySelectorAll('.oxe-ico, .oxe-ctxico').forEach((e) => {
    const d = ICON_PATHS[e.textContent.trim()];
    if (!d) return;
    e.innerHTML = `<svg viewBox="0 0 24 24" width="1.1em" height="1.1em" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" style="vertical-align:-0.15em"><path d="${d}"/></svg>`;
  });
}

export function el(tag, attrs = {}, ...children) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    // Controls act on mouse-down (so the sheet keeps its selection); Enter and Space do the same from the keyboard.
    if (k === 'onmousedown' && v) e.addEventListener('keydown', (ev) => {
      if (ev.key !== 'Enter' && ev.key !== ' ') return;
      const r = e.getBoundingClientRect();
      ev.preventDefault();
      e.oxeKeyAt = performance.now();
      v({ preventDefault() {}, stopPropagation() {}, target: e, currentTarget: e, clientX: r.left, clientY: r.bottom, button: 0, detail: 1, viaKeyboard: true });
    });
    // A screen reader in browse mode activates with a bare click (no mouse-down), which acts like a key press.
    if (k === 'onmousedown' && v) e.addEventListener('click', (ev) => {
      if (ev.detail !== 0 || performance.now() - (e.oxeKeyAt || -1e9) < 500 || performance.now() - (e.oxeDownAt || -1e9) < 1000) return;
      const r = e.getBoundingClientRect();
      v({ preventDefault() {}, stopPropagation() {}, target: e, currentTarget: e, clientX: r.left, clientY: r.bottom, button: 0, detail: 1, viaKeyboard: true });
    });
    if (k === 'onmousedown' && v) e.addEventListener('mousedown', () => { e.oxeDownAt = performance.now(); }, true);
    else if (v != null && v !== false && !k.startsWith('on')) e.setAttribute(k, v === true ? '' : v);
  }
  e.append(...children.filter((c) => c != null && c !== false));
  return e;
}

/** A tab or group label as a feature name: 'Font' -> 'font', 'Page Layout' -> 'page-layout'. */
export const slug = (t) => String(t || '').toLowerCase().replace(/&/g, 'and').replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');

/**
 * The controls of one ribbon, bound to its container (where the menu is placed). `home()` gives focus back to the
 * sheet. A command is a function of no arguments.
 */
export function ribbonKit(container, home) {
  let lastPointer = -1e9;
  const onPointer = () => { lastPointer = performance.now(); };
  document.addEventListener('pointerdown', onPointer, true);
  /** A command as a mouse-down handler. */
  const run = (command) => (e) => {
    e.preventDefault();
    closeMenu();
    command();
    // From the keyboard, focus stays on the button (as in Excel) so its new state is heard; a click returns to the sheet.
    // A screen reader's click (Space in browse mode) has no pointer behind it, so it counts as keyboard too, and so
    // does pressing a button that already had focus. (Not an item of a menu, which has just closed: back to the sheet.)
    const keys = e.viaKeyboard || performance.now() - lastPointer > 500 || document.activeElement === e.target?.closest?.('button');
    const inMenu = !!e.target?.closest?.('.oxe-pop');
    if (keys && !inMenu && e.target?.isConnected) e.target.closest('button')?.focus(); else home();
  };

  /** A ribbon button: `icon` is a short glyph, `label` shows under it when `big`. */
  const btn = (icon, title, command, { label = null, big = false, cls = '' } = {}) =>
    el('button', { type: 'button', title, 'aria-label': label && !title.toLowerCase().includes(String(label).toLowerCase()) ? `${label}: ${title}` : title, class: `${big ? 'oxe-big' : ''} ${cls}`.trim() || null, onmousedown: run(command) },
      el('span', { class: 'oxe-ico', 'aria-hidden': 'true' }, icon), label && el('span', { class: 'oxe-lbl' }, label));
  /** A toggle button (Bold, Wrap Text...): `aria-pressed` follows its `on` class. */
  const toggle = (icon, title, command, opts) => {
    const b = btn(icon, title, command, opts);
    b.setAttribute('aria-pressed', 'false');
    b.setOn = (on) => { b.classList.toggle('on', !!on); b.setAttribute('aria-pressed', String(!!on)); };
    return b;
  };
  const grp = (label, ...kids) => el('div', { class: 'oxe-group', 'data-feature': slug(label) }, el('div', { class: 'oxe-rrow' }, ...kids), label && el('div', { class: 'oxe-glabel' }, label));
  const col = (...kids) => el('div', { class: 'oxe-rcol' }, ...kids);
  const line = (...kids) => el('div', { class: 'oxe-rline' }, ...kids);

  // --- Drop-down menus (one open at a time) ---
  const pop = el('div', { class: 'oxe-pop', hidden: true, role: 'menu' });
  let popOwner = null;
  function closeMenu(refocus = false) { const o = popOwner; pop.hidden = true; popOwner = null; o?.setAttribute('aria-expanded', 'false'); if (refocus) o?.focus(); }
  pop.addEventListener('keydown', (e) => {
    const items = [...pop.querySelectorAll('button:not([disabled]), input, select')];
    const i = items.indexOf(document.activeElement);
    if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); closeMenu(true); }
    else if (e.key === 'Tab') closeMenu(false);
    else if (e.key === 'ArrowDown' || e.key === 'ArrowRight') { e.preventDefault(); items[(i + 1) % items.length]?.focus(); }
    else if (e.key === 'ArrowUp' || e.key === 'ArrowLeft') { e.preventDefault(); items[(i - 1 + items.length) % items.length]?.focus(); }
  });
  /** Opens the menu under `owner` with the elements `build()` returns (a second press on the owner closes it). */
  function openMenu(owner, build) {
    if (popOwner === owner) return closeMenu();
    pop.replaceChildren(...build());
    iconize(pop);
    pop.hidden = false;
    popOwner = owner;
    owner.setAttribute('aria-expanded', 'true');
    pop.setAttribute('aria-label', owner.getAttribute('aria-label') || owner.textContent);
    pop.querySelectorAll('button').forEach((b) => { if (!b.getAttribute('role')) b.setAttribute('role', 'menuitem'); if (!b.getAttribute('aria-label') && b.title && !b.textContent.trim()) b.setAttribute('aria-label', b.title); });
    if (owner.contains(document.activeElement) || owner === document.activeElement) setTimeout(() => pop.querySelector('button:not([disabled]), input, select')?.focus(), 300); // after the reader has said "expanded"
    const r = owner.getBoundingClientRect();
    const host = container.getBoundingClientRect();
    pop.style.left = `${Math.max(0, Math.min(r.left - host.left, host.width - 240))}px`;
    pop.style.top = `${r.bottom - host.top + 2}px`;
  }
  const onOutside = (e) => { if (!pop.hidden && !pop.contains(e.target) && !popOwner?.contains(e.target)) closeMenu(); };
  document.addEventListener('mousedown', onOutside);
  const item = (label, command, { on = false, hint = '', icon = null } = {}) => el('button', { type: 'button', class: `oxe-item${on ? ' on' : ''}`, onmousedown: run(command) },
    icon && el('span', { class: 'oxe-ctxico', 'aria-hidden': 'true' }, icon), el('span', {}, label), hint && el('small', {}, hint));
  /** A button with a ▾ that opens `items()` (an array of elements). */
  const menu = (icon, title, items, opts = {}) => {
    const b = el('button', { type: 'button', title, 'aria-label': title, 'aria-haspopup': 'menu', 'aria-expanded': 'false', class: `oxe-menu-btn ${opts.big ? 'oxe-big' : ''} ${opts.cls || ''}`.trim(), onmousedown: (e) => { e.preventDefault(); openMenu(b, items); } },
      el('span', { class: 'oxe-ico' }, icon), opts.label && el('span', { class: 'oxe-lbl' }, opts.label), el('span', { class: 'oxe-caret' }, '▾'));
    return b;
  };
  /** A split button: the part that runs `command` (the last choice) and a ▾ that opens the choices. */
  const split = (icon, title, command, items, opts = {}) => {
    const main = btn(icon, title, command, opts);
    const more = el('button', { type: 'button', title: `${title} options`, 'aria-label': `${title} options`, 'aria-haspopup': 'menu', 'aria-expanded': 'false', class: 'oxe-split-more', onmousedown: (e) => { e.preventDefault(); openMenu(more, items); } }, el('span', { class: 'oxe-caret' }, '▾'));
    const wrap = el('span', { class: 'oxe-split' }, main, more);
    wrap.main = main;
    return wrap;
  };
  /** A select whose choice goes to `onPick(value)`, then back to the sheet. */
  const pick = (title, cls, options, onPick) => el('select', { title, 'aria-label': title, class: cls, onchange: (e) => { onPick(e.target.value); home(); } }, ...options);

  const destroy = () => {
    document.removeEventListener('pointerdown', onPointer, true);
    document.removeEventListener('mousedown', onOutside);
  };
  return { run, btn, toggle, grp, col, line, pop, openMenu, closeMenu, item, menu, split, pick, destroy };
}
