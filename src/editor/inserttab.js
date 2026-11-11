// The Insert tab, Excel's groups as far as they are here: Tables (PivotTable: its boxes, pivotbox.js, and the
// PivotTable made, pivotnew.js).

/** The Insert tab. `act(name, ...args)` runs a command. */
export function insertTab(kit, { act }) {
  const { grp, btn } = kit;
  const groups = [
    grp('Tables', btn('▦', 'PivotTable', () => act('pivotBox'), { label: 'PivotTable', big: true })),
  ];
  return { groups, update() {} };
}
