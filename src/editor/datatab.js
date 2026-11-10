// The Data tab, Excel's groups in Excel's order as far as they are here: Queries & Connections (Refresh All, its
// menu Refresh All and Refresh: the workbook's PivotTables, or the one at the active cell, pivot.js).

/** The Data tab. `act(name, ...args)` runs a command. */
export function dataTab(kit, { act }) {
  const { grp, item, split } = kit;
  const f = (name, ...args) => () => act(name, ...args);
  const groups = [
    grp('Queries & Connections', split('⟳', 'Refresh All (Ctrl+Alt+F5)', f('refreshPivots', true), () => [
      item('Refresh All', f('refreshPivots', true), { hint: 'Ctrl+Alt+F5' }), item('Refresh', f('refreshPivots', false), { hint: 'Alt+F5' })], { label: 'Refresh All', big: true })),
  ];
  return { groups, update() {} };
}
