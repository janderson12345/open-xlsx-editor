// The Page Layout tab, Excel's groups as far as they are here: Sheet Options (Gridlines and Headings, each with View,
// the same switch as View > Show, and Print: ooxml/printopts.js). Excel's check boxes; here buttons that stay pressed
// while ticked.

/** The Page Layout tab. `act(name, ...args)` runs a command. */
export function pageLayoutTab(kit, { act }) {
  const { grp, col, toggle } = kit;
  const opts = [['showGridlines', 'Gridlines: View', 'showGridLines'], ['printGridlines', 'Gridlines: Print', 'printGridLines'],
    ['showHeadings', 'Headings: View', 'showHeadings'], ['printHeadings', 'Headings: Print', 'printHeadings']]
    .map(([name, title, key]) => ({ key, b: toggle(name.startsWith('print') ? '⎙' : '👁', title, () => act(name), { label: title.split(': ')[1] }) }));
  const groups = [grp('Sheet Options', col(opts[0].b, opts[1].b), col(opts[2].b, opts[3].b))];
  return {
    groups,
    /** From the sheet shown: { view: { showGridLines, showHeadings, printGridLines, printHeadings } }. */
    update(look) {
      const v = look.view;
      if (v) for (const o of opts) o.b.setOn(v[o.key]);
    },
  };
}
