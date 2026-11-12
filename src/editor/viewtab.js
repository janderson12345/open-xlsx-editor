// The View tab, Excel's groups in Excel's order as far as they are here: Show (Gridlines, Formula Bar, Headings: Excel's
// check boxes, here buttons that stay pressed while on), Zoom (Zoom..., 100%, Zoom to Selection) and Window (Freeze Panes, Freeze Top Row, Freeze First Column, and Unfreeze Panes once frozen, as Excel's menu turns).
import { el } from './ribbonkit.js';

/** The View tab. `act(name, ...args)` runs a command. */
export function viewTab(kit, { act }) {
  const { btn, toggle, grp, col, item, menu } = kit;
  const f = (name, ...args) => () => act(name, ...args);
  let frozen = false;
  const zoomLabel = el('span', { class: 'oxe-zoomlabel', 'aria-live': 'off' }, '100%');
  const shows = [['showGridlines', '▦', 'Gridlines', 'showGridLines'], ['showFormulaBar', '▭', 'Formula Bar', 'formulaBar'], ['showHeadings', '⊞', 'Headings', 'showHeadings']]
    .map(([name, icon, label, key]) => { const b = toggle(icon, label, f(name), { label }); b.setOn(true); return { key, b }; });
  const groups = [
    grp('Show', col(...shows.map((s) => s.b))),
    grp('Zoom', btn('🔎', 'Zoom (10% to 400%)', f('zoomBox'), { label: 'Zoom', big: true }), col(btn('1:1', '100%', f('zoom', 100), { label: '100%' }),
      btn('⤢', 'Zoom to Selection', f('zoomSelection'), { label: 'Zoom to Selection' })), zoomLabel),
    grp('Window', menu('❄', 'Freeze Panes', () => [
      frozen ? item('Unfreeze Panes', f('freeze', 'none')) : item('Freeze Panes', f('freeze', 'panes'), { hint: 'Keep the rows above and columns left of the active cell in view' }),
      item('Freeze Top Row', f('freeze', 'topRow')), item('Freeze First Column', f('freeze', 'firstColumn'))], { label: 'Freeze Panes', big: true })),
  ];
  return {
    groups,
    /** From the sheet shown: { view: { freeze, zoom, showGridLines, showHeadings, formulaBar } }. */
    update(look) {
      const v = look.view;
      if (!v) return;
      frozen = !!(v.freeze.rows || v.freeze.cols);
      zoomLabel.textContent = `${v.zoom}%`;
      for (const s of shows) s.b.setOn(v[s.key] !== false);
    },
  };
}
