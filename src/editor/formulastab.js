// The Formulas tab, Excel's groups as far as they are here: Function Library (Insert Function, Shift+F3; AutoSum and
// its choices; Recently Used and each category's functions, More Functions for the rest: functionbox.js and
// calc/catalog.js), Defined Names (Name Manager, Ctrl+F3, and Define Name: namesbox.js and names.js), Formula Auditing
// (Trace Precedents, Trace Dependents, Remove Arrows: tracing.js; Show Formulas, Ctrl+`) and Calculation (Calculation Options, Calculate Now, Calculate Sheet: ooxml/calcmode.js).
import { el } from './ribbonkit.js';
import { inCategory, functionInfo, syntaxOf, DEFAULT_RECENT } from '../calc/catalog.js';

/** The Formulas tab. `act(name, ...args)` runs a command; `recent()` the Most Recently Used functions; `calcMode()`
 *  the workbook's calculation mode. */
export function formulasTab(kit, { act, recent = () => null, calcMode = () => 'auto' }) {
  const { grp, btn, toggle, col, menu, split, item, openMenu, pop } = kit;
  const fn = (name) => item(name, () => act('insertFunction', name));
  const titled = (b, name) => { b.title = syntaxOf(name); return b; };
  const list = (names) => () => names.map((n) => titled(fn(n), n));
  const category = (c) => list(inCategory(c).map((f) => f.name));
  const lists = [['$', 'Financial'], ['?', 'Logical'], ['A', 'Text'], ['◷', 'Date & Time'], ['🔍', 'Lookup & Reference'], ['θ', 'Math & Trig']];
  const auto = (f) => () => act('autoSum', f);
  // (More Functions: its categories, each opening its functions in the same place.)
  let more = null;
  const moreItems = () => ['Statistical', 'Engineering', 'Information', 'Compatibility', 'Web'].map((c) =>
    item(`${c} ▸`, () => setTimeout(() => {
      openMenu(more, () => [...category(c)(), el('hr'), item('Insert Function...', () => act('insertFunction'))]);
      pop.querySelector('button')?.focus();
    })));
  more = menu('⋯', 'More Functions', moreItems, { label: 'More Functions' });
  const showBtn = toggle('⨍', 'Show Formulas (Ctrl+`)', () => act('showFormulas'), { label: 'Show Formulas' });
  const groups = [
    grp('Function Library', btn('ƒx', 'Insert Function (Shift+F3)', () => act('insertFunction'), { label: 'Insert Function', big: true }),
      split('∑', 'AutoSum', auto('sum'), () => [item('Sum', auto('sum'), { hint: 'Alt+=' }), item('Average', auto('average')), item('Count Numbers', auto('count')),
        item('Max', auto('max')), item('Min', auto('min')), el('hr'), item('More Functions...', () => act('insertFunction'))], { label: 'AutoSum', big: true }),
      menu('★', 'Recently Used', () => [...list((recent() || DEFAULT_RECENT).filter((n) => functionInfo(n)))(), el('hr'), item('Insert Function...', () => act('insertFunction'))], { label: 'Recently Used', big: true }),
      ...lists.map(([icon, c]) => menu(icon, c, () => [...category(c)(), el('hr'), item('Insert Function...', () => act('insertFunction'))], { label: c, big: true })),
      col(more)),
    grp('Defined Names', btn('🏷', 'Name Manager (Ctrl+F3)', () => act('nameManager'), { label: 'Name Manager', big: true }),
      col(btn('＋', 'Define Name', () => act('defineName'), { label: 'Define Name' }),
        btn('⊞', 'Create from Selection (Ctrl+Shift+F3)', () => act('createNames'), { label: 'Create from Selection' }))),
    grp('Formula Auditing', col(btn('⇢', 'Trace Precedents', () => act('trace', 'precedents'), { label: 'Trace Precedents' }),
      btn('⇠', 'Trace Dependents', () => act('trace', 'dependents'), { label: 'Trace Dependents' }),
      split('✕', 'Remove Arrows', () => act('removeArrows'), () => [item('Remove Arrows', () => act('removeArrows')),
        item('Remove Precedent Arrows', () => act('removeArrows', 'precedents')), item('Remove Dependent Arrows', () => act('removeArrows', 'dependents'))], { label: 'Remove Arrows' })),
      col(showBtn, btn('⚠', 'Error Checking', () => act('errorChecking'), { label: 'Error Checking' }),
        btn('ƒ⇣', 'Evaluate Formula', () => act('evaluateFormula'), { label: 'Evaluate Formula' })),
      btn('👁', 'Watch Window', () => act('watchWindow'), { label: 'Watch Window', big: true })),
    grp('Calculation', menu('⚙', 'Calculation Options', () => [['auto', 'Automatic'], ['autoNoTable', 'Automatic Except for Data Tables'], ['manual', 'Manual']]
      .map(([m, label]) => item(label, () => act('calcMode', m), { on: calcMode() === m, icon: calcMode() === m ? '✓' : ' ' })), { label: 'Calculation Options', big: true }),
    col(btn('▦', 'Calculate Now (F9)', () => act('calculate', 'now'), { label: 'Calculate Now' }),
      btn('▤', 'Calculate Sheet (Shift+F9)', () => act('calculate', 'sheet'), { label: 'Calculate Sheet' }))),
  ];
  return {
    groups,
    /** From the sheet shown: Show Formulas pressed while on. */
    update(look) { showBtn.setOn(look.view?.showFormulas); },
  };
}
