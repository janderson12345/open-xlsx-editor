// The Data tab's Outline group's actions that ask first (acts.js): Group and Ungroup on a selection that isn't whole
// rows or columns (Excel's box asks which), the Settings box and the Subtotal box. Each runs an outline command
// (commands-outline.js), one step of the history.
import { MAX_ROWS, MAX_COLS, colName } from '../ooxml/refs.js';
import { currentRegion } from '../ooxml/special.js';
import { groupAxisBox, outlineSettingsBox, subtotalBox } from './outlinebox.js';

export const OUTLINE_ACTS = new Set(['groupBox', 'ungroupBox', 'outlineSettingsBox', 'subtotalBox']);

/** `env` is acts.js's; returns run(name, args, i). */
export function createOutlineActs(env) {
  const { root, state, editing } = env;
  const back = () => state.grid.el.focus({ preventScroll: true });
  return (name, args, i) => {
    const s = editing().selection;
    if (name === 'groupBox' || name === 'ungroupBox') {
      // (Whole rows or columns go at once; a range asks Rows or Columns, as Excel's Group box does.)
      const cmd = name === 'groupBox' ? 'group' : 'ungroup';
      const rows = s.c0 === 0 && s.c1 >= MAX_COLS - 1;
      const cols = s.r0 === 0 && s.r1 >= MAX_ROWS - 1;
      if (rows !== cols) return editing().run(cmd, rows ? 'row' : 'col');
      return groupAxisBox(root, name === 'groupBox' ? 'Group' : 'Ungroup').then((axis) => { back(); return axis ? editing().run(cmd, axis) : false; });
    }
    if (name === 'outlineSettingsBox') {
      const o = state.workbook.sheet(i).outline || { below: true, right: true };
      return outlineSettingsBox(root, o).then((got) => { back(); return got ? editing().run('outlineSettings', got) : false; });
    }
    if (name === 'subtotalBox') {
      // (The list: the selection, or the table round the active cell; its first row's headings name its columns.)
      const sheet = state.workbook.sheet(i);
      const area = s.r0 === s.r1 && s.c0 === s.c1 ? currentRegion(sheet, s.ar, s.ac) : { r0: s.r0, c0: s.c0, r1: Math.min(s.r1, sheet.maxRow), c1: Math.min(s.c1, sheet.maxCol) };
      const columns = [];
      for (let c = area.c0; c <= area.c1; c++) { const v = sheet.cells.get(area.r0)?.get(c)?.v; columns.push({ c, label: v == null || v === '' ? `Column ${colName(c)}` : String(v) }); }
      return subtotalBox(root, { columns, below: sheet.outline?.below !== false }).then((got) => {
        back();
        if (!got) return false;
        return got.remove ? editing().run('removeSubtotals', area) : editing().run('subtotal', area, got);
      });
    }
    return undefined;
  };
}
