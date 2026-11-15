// The Page Layout tab, Excel's groups as far as they are here: Page Setup (Print Area, Breaks and Print Titles:
// ooxml/pagesetupedit.js, pagesetupbox.js) and Sheet Options (Gridlines and Headings, each with View, the same switch as View > Show,
// and Print: ooxml/printopts.js). Excel's check boxes; here buttons that stay pressed while ticked.

/** The Page Layout tab. `act(name, ...args)` runs a command; `hasPrintArea()`: whether the sheet shown has one (Excel
 *  then offers Add to Print Area). */
export function pageLayoutTab(kit, { act, hasPrintArea = () => false }) {
  const { grp, col, toggle, menu, item, btn } = kit;
  const opts = [['showGridlines', 'Gridlines: View', 'showGridLines'], ['printGridlines', 'Gridlines: Print', 'printGridLines'],
    ['showHeadings', 'Headings: View', 'showHeadings'], ['printHeadings', 'Headings: Print', 'printHeadings']]
    .map(([name, title, key]) => ({ key, b: toggle(name.startsWith('print') ? '⎙' : '👁', title, () => act(name), { label: title.split(': ')[1] }) }));
  const setup = grp('Page Setup',
    menu('▣', 'Print Area', () => [item('Set Print Area', () => act('printArea', 'set')), item('Clear Print Area', () => act('printArea', 'clear')),
      ...(hasPrintArea() ? [item('Add to Print Area', () => act('printArea', 'add'))] : [])], { label: 'Print Area', big: true }),
    menu('⤓', 'Breaks', () => [item('Insert Page Break', () => act('pageBreak', 'insert')), item('Remove Page Break', () => act('pageBreak', 'remove')),
      item('Reset All Page Breaks', () => act('pageBreak', 'reset'))], { label: 'Breaks', big: true }),
    btn('▤', 'Print Titles', () => act('printTitlesBox'), { label: 'Print Titles', big: true }));
  const groups = [setup, grp('Sheet Options', col(opts[0].b, opts[1].b), col(opts[2].b, opts[3].b))];
  return {
    groups,
    /** From the sheet shown: { view: { showGridLines, showHeadings, printGridLines, printHeadings } }. */
    update(look) {
      const v = look.view;
      if (v) for (const o of opts) o.b.setOn(v[o.key]);
    },
  };
}
