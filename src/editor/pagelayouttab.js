// The Page Layout tab, Excel's groups as far as they are here: Page Setup (Print Area, Breaks and Print Titles:
// ooxml/pagesetupedit.js, pagesetupbox.js) and Sheet Options (Gridlines and Headings, each with View, the same switch as View > Show,
// and Print: ooxml/printopts.js). Excel's check boxes; here buttons that stay pressed while ticked.

import { MARGINS } from '../ooxml/pagesetupedit.js';

/** Excel's Size list, as far as the print layout knows the papers: [label, paperSize code]. */
const SIZES = [['Letter 8.5" x 11"', 1], ['Legal 8.5" x 14"', 5], ['Executive 7.25" x 10.5"', 7], ['Tabloid 11" x 17"', 3], ['A3 11.69" x 16.54"', 8],
  ['A4 8.27" x 11.69"', 9], ['A5 5.83" x 8.27"', 11], ['B5 (JIS) 7.17" x 10.12"', 13]];
const PAGES = [['Automatic', 0], ...Array.from({ length: 9 }, (_, k) => [`${k + 1} page${k ? 's' : ''}`, k + 1])];

/** The Page Layout tab. `act(name, ...args)` runs a command; `hasPrintArea()`: whether the sheet shown has one (Excel
 *  then offers Add to Print Area); `page()`: the sheet shown's page setup (sheet.js readPageSetup), to tick the
 *  current choices. */
export function pageLayoutTab(kit, { act, hasPrintArea = () => false, page = () => null }) {
  const { grp, col, toggle, menu, item, btn } = kit;
  const opts = [['showGridlines', 'Gridlines: View', 'showGridLines'], ['printGridlines', 'Gridlines: Print', 'printGridLines'],
    ['showHeadings', 'Headings: View', 'showHeadings'], ['printHeadings', 'Headings: Print', 'printHeadings']]
    .map(([name, title, key]) => ({ key, b: toggle(name.startsWith('print') ? '⎙' : '👁', title, () => act(name), { label: title.split(': ')[1] }) }));
  const tick = (on) => ({ on, icon: on ? '✓' : ' ' });
  const same = (m, p) => !!m && Object.keys(p).every((k) => Math.abs(m[k] - p[k]) < 1e-9);
  const marginNames = [['Normal', 'normal', 'Top 0.75" Bottom 0.75" Left 0.7" Right 0.7"'], ['Wide', 'wide', 'Top 1" Bottom 1" Left 1" Right 1"'], ['Narrow', 'narrow', 'Top 0.75" Bottom 0.75" Left 0.25" Right 0.25"']];
  // (Width and height: Automatic is 0; fit to pages off is Automatic both ways.)
  const fitNow = () => { const p = page(); return p && p.fitToPage ? { width: p.fitToWidth, height: p.fitToHeight } : { width: 0, height: 0 }; };
  const setup = grp('Page Setup',
    menu('▭', 'Margins', () => marginNames.map(([label, key, hint]) => item(label, () => act('pageLayout', { margins: key }), { hint, ...tick(same(page()?.margins, MARGINS[key])) })), { label: 'Margins', big: true }),
    menu('⇄', 'Orientation', () => [['Portrait', false], ['Landscape', true]].map(([label, v]) => item(label, () => act('pageLayout', { landscape: v }), tick(!!page()?.landscape === v))), { label: 'Orientation', big: true }),
    menu('▯', 'Size', () => SIZES.map(([label, code]) => item(label, () => act('pageLayout', { paperSize: code }), tick((page()?.paperSize ?? 1) === code))), { label: 'Size', big: true }),
    menu('▣', 'Print Area', () => [item('Set Print Area', () => act('printArea', 'set')), item('Clear Print Area', () => act('printArea', 'clear')),
      ...(hasPrintArea() ? [item('Add to Print Area', () => act('printArea', 'add'))] : [])], { label: 'Print Area', big: true }),
    menu('⤓', 'Breaks', () => [item('Insert Page Break', () => act('pageBreak', 'insert')), item('Remove Page Break', () => act('pageBreak', 'remove')),
      item('Reset All Page Breaks', () => act('pageBreak', 'reset'))], { label: 'Breaks', big: true }),
    btn('▤', 'Print Titles', () => act('printTitlesBox'), { label: 'Print Titles', big: true }));
  const fitGroup = grp('Scale to Fit', col(
    menu('↔', 'Width', () => PAGES.map(([label, n]) => item(label, () => act('pageLayout', { fit: { ...fitNow(), width: n } }), tick(fitNow().width === n))), { label: 'Width' }),
    menu('↕', 'Height', () => PAGES.map(([label, n]) => item(label, () => act('pageLayout', { fit: { ...fitNow(), height: n } }), tick(fitNow().height === n))), { label: 'Height' }),
    btn('%', 'Scale', () => act('scaleBox'), { label: 'Scale' })));
  const groups = [setup, fitGroup, grp('Sheet Options', col(opts[0].b, opts[1].b), col(opts[2].b, opts[3].b))];
  return {
    groups,
    /** From the sheet shown: { view: { showGridLines, showHeadings, printGridLines, printHeadings } }. */
    update(look) {
      const v = look.view;
      if (v) for (const o of opts) o.b.setOn(v[o.key]);
    },
  };
}
