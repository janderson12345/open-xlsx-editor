// The Formulas tab, Excel's groups as far as they are here: Defined Names (Name Manager, Ctrl+F3, and Define Name:
// namesbox.js and names.js).

/** The Formulas tab. `act(name, ...args)` runs a command. */
export function formulasTab(kit, { act }) {
  const { grp, btn, col } = kit;
  const groups = [
    grp('Defined Names', btn('🏷', 'Name Manager (Ctrl+F3)', () => act('nameManager'), { label: 'Name Manager', big: true }),
      col(btn('＋', 'Define Name', () => act('defineName'), { label: 'Define Name' }))),
  ];
  return { groups, update() {} };
}
