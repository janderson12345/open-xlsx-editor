// The Data tab, Excel's groups in Excel's order as far as they are here: Queries & Connections (Refresh All, its
// menu Refresh All and Refresh: the workbook's PivotTables, or the one at the active cell, pivot.js) and Sort & Filter
// (Sort A to Z, Sort Z to A, Sort..., Filter, Clear, Reapply: the commands Home > Editing's menu runs too) and Data
// Tools (Data Validation..., dvedit.js).

/** The Data tab. `act(name, ...args)` runs a command. */
export function dataTab(kit, { act }) {
  const { btn, grp, col, item, split } = kit;
  const f = (name, ...args) => () => act(name, ...args);
  const groups = [
    grp('Queries & Connections', split('⟳', 'Refresh All (Ctrl+Alt+F5)', f('refreshPivots', true), () => [
      item('Refresh All', f('refreshPivots', true), { hint: 'Ctrl+Alt+F5' }), item('Refresh', f('refreshPivots', false), { hint: 'Alt+F5' })], { label: 'Refresh All', big: true })),
    grp('Sort & Filter',
      col(btn('A↓', 'Sort A to Z', f('sortAsc')), btn('Z↓', 'Sort Z to A', f('sortDesc'))),
      btn('⇅', 'Sort...', f('sortBox'), { label: 'Sort', big: true }),
      btn('⏷', 'Filter (Ctrl+Shift+L)', f('filterToggle'), { label: 'Filter', big: true }),
      col(btn('✕', 'Clear', f('filterClear'), { label: 'Clear' }), btn('↻', 'Reapply', f('filterReapply'), { label: 'Reapply' }), btn('⚙', 'Advanced...', f('advancedFilterBox'), { label: 'Advanced' }))),
    grp('Data Tools', btn('⫼', 'Text to Columns...', f('textToColumnsBox'), { label: 'Text to Columns', big: true }), col(btn('⚡', 'Flash Fill (Ctrl+E)', f('flashFill'), { label: 'Flash Fill' }),
      btn('⊖', 'Remove Duplicates...', f('removeDuplicatesBox'), { label: 'Remove Duplicates' })), btn('☑', 'Data Validation...', f('validationBox'), { label: 'Data Validation', big: true })),
  ];
  return { groups, update() {} };
}
