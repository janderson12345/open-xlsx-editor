// The ribbon: Excel's tabs over the sheet, each tab's groups from a file of its own (hometab.js), as open-docx-editor's
// toolbar.js lays them out: groups that don't fit go behind "More", a one-line simplified ribbon remembered per
// browser, F6 between the sheet and the ribbon, Esc back to the sheet, and `features` to hide a tab or a group by name.
import { el, iconize, slug, ribbonKit } from './ribbonkit.js';

/**
 * A ribbon in `container`. `tabs`: { name: (kit) => { groups: [elements], update(look) } }; `home()` gives focus back
 * to the sheet. Returns { update(look), setFeatures(f), features, showTab(name), destroy() }.
 */
export function createRibbon(container, { tabs, home }) {
  const kit = ribbonKit(container, home);
  const { pop, closeMenu } = kit;
  const built = Object.fromEntries(Object.entries(tabs).map(([name, make]) => [name, make(kit)]));
  let current = Object.keys(built)[0];
  const panels = {};
  const tabBtns = {};
  const strip = el('div', { class: 'oxe-rtabs' });
  const tablist = el('div', { class: 'oxe-rtablist', role: 'tablist', 'aria-label': 'Ribbon tabs' });
  strip.append(tablist);
  const ribbonId = `oxe-rb${Math.random().toString(36).slice(2, 7)}`;
  const body = el('div', { class: 'oxe-ribbon' });
  let seq = 0;
  for (const [name, tab] of Object.entries(built)) {
    const id = `${ribbonId}-${seq++}`;
    panels[name] = el('div', { class: 'oxe-panel', role: 'tabpanel', id: `${id}-p`, 'aria-labelledby': `${id}-t`, hidden: name !== current }, ...tab.groups);
    tabBtns[name] = el('button', { type: 'button', role: 'tab', id: `${id}-t`, 'aria-controls': `${id}-p`, 'aria-selected': String(name === current), tabindex: name === current ? '0' : '-1', class: `oxe-rtab${name === current ? ' on' : ''}`, onmousedown: (e) => { e.preventDefault(); showTab(name); } }, name);
    tabBtns[name].addEventListener('keydown', (e) => {
      const names = Object.keys(tabBtns).filter((k) => !tabBtns[k].hidden);
      const i = names.indexOf(name);
      const to = { ArrowRight: names[(i + 1) % names.length], ArrowLeft: names[(i - 1 + names.length) % names.length], Home: names[0], End: names.at(-1) }[e.key];
      if (to) { e.preventDefault(); showTab(to); tabBtns[to].focus(); }
    });
    tablist.append(tabBtns[name]);
    body.append(panels[name]);
  }

  // Groups that don't fit on the row are hidden behind "More"; it opens the rest as further rows.
  const moreBtn = el('button', { type: 'button', class: 'oxe-more', title: 'More commands', 'aria-label': 'More commands', 'aria-expanded': 'false', hidden: true,
    onmousedown: (e) => { e.preventDefault(); const on = !body.classList.contains('oxe-expand'); body.classList.toggle('oxe-expand', on); moreBtn.setAttribute('aria-expanded', String(on)); fitRibbon(); } }, '⋯');
  body.append(moreBtn);
  function fitRibbon() {
    const panel = panels[current];
    Object.values(panels).forEach((p) => [...p.children].forEach((g) => g.classList.remove('oxe-cut')));
    if (!panel || panel.hidden) return;
    const room = panel.clientWidth - 36;
    const cut = [...panel.children].filter((g) => !g.hidden && g.offsetLeft + g.offsetWidth > room);
    const over = cut.length > 0 || (body.classList.contains('oxe-expand') && panel.scrollHeight > (panel.firstElementChild?.offsetHeight || 0) + 8);
    moreBtn.hidden = !over;
    if (!over) { body.classList.remove('oxe-expand'); moreBtn.setAttribute('aria-expanded', 'false'); }
    if (!body.classList.contains('oxe-expand')) cut.forEach((g) => g.classList.add('oxe-cut'));
  }
  const resize = typeof ResizeObserver === 'function' ? new ResizeObserver(() => fitRibbon()) : null;
  resize?.observe(body);

  function showTab(name) {
    current = name;
    closeMenu();
    for (const k of Object.keys(panels)) {
      panels[k].hidden = k !== name;
      tabBtns[k].classList.toggle('on', k === name);
      tabBtns[k].setAttribute('aria-selected', String(k === name));
      tabBtns[k].tabIndex = k === name ? 0 : -1;
    }
    fitRibbon();
  }

  // Simplified ribbon (one line) or the classic one; remembered per browser.
  let simple = false;
  try { simple = localStorage.getItem('oxe-ribbon') === 'simple'; } catch { /* storage blocked */ }
  const layoutBtn = el('button', { type: 'button', class: 'oxe-ribbon-mode', onmousedown: (e) => { e.preventDefault(); setSimple(!simple); } });
  function setSimple(on) {
    simple = on;
    container.classList.toggle('oxe-simple', on);
    layoutBtn.textContent = on ? '⌄' : '⌃';
    layoutBtn.title = on ? 'Classic ribbon' : 'Simplified ribbon';
    layoutBtn.setAttribute('aria-label', layoutBtn.title);
    try { localStorage.setItem('oxe-ribbon', on ? 'simple' : 'classic'); } catch { /* storage blocked */ }
    fitRibbon();
  }
  strip.append(el('span', { class: 'oxe-rtabs-fill' }), layoutBtn);
  setSimple(simple);

  container.classList.add('oxe-toolbar');
  container.replaceChildren(strip, body, pop);
  iconize(container);

  // F6 moves between the sheet and the ribbon, as in Excel; Esc in the ribbon goes back to the sheet.
  const onF6 = (e) => {
    if (e.key !== 'F6' || !container.isConnected || container.offsetParent === null) return;
    const root = container.parentElement;
    if (!root?.contains(document.activeElement) && !container.contains(document.activeElement)) return;
    e.preventDefault();
    if (container.contains(document.activeElement)) home();
    else tabBtns[current]?.focus();
  };
  document.addEventListener('keydown', onF6);
  container.addEventListener('keydown', (e) => { if (e.key === 'Escape' && pop.hidden) { e.preventDefault(); home(); } });

  /** Ribbon features off or on: a tab by name ({ home: false }) or a group by its label on any tab ({ font: false }). */
  let features = {};
  function applyFeatures() {
    let first = null;
    for (const name of Object.keys(panels)) {
      const off = features[slug(name)] === false;
      tabBtns[name].hidden = off;
      if (!off && !first) first = name;
      panels[name].querySelectorAll('.oxe-group').forEach((g) => { g.hidden = features[g.dataset.feature] === false; });
    }
    if (features[slug(current)] === false && first) showTab(first);
    fitRibbon();
  }

  return {
    el: container,
    /** Each tab's controls set to the active cell's look (formatops.js activeLook, plus what the tab needs). */
    update(look) { for (const tab of Object.values(built)) tab.update?.(look); },
    setFeatures(f) { features = { ...features, ...f }; applyFeatures(); },
    get features() { return { ...features }; },
    showTab,
    destroy() {
      kit.destroy();
      resize?.disconnect();
      document.removeEventListener('keydown', onF6);
      container.remove();
    },
  };
}
