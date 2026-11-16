// The parts of the editor made with the first workbook opened (index.js open()): the grid with its buttons and charts'
// picker, the editing attached to it (editor/interact.js) with what it calls back for, and the sheet tabs. `env` is
// mountXlsxEditor's shared state and helpers, as one object.
import { createGrid } from './grid.js';
import { chartPicker } from './chartpick.js';
import { readNotes } from '../ooxml/notes.js';
import { attachEditing } from './interact.js';
import { unlockRangeBox } from './rangesbox.js';
import { createSheetTabs } from './tabs.js';
import { lastCopy, setLastCopy } from './pagebooks.js';

/** Make the grid, its editing and the tabs in `env.root` for workbook `wb` showing sheet `i` (`sheet`, its model). */
export function createParts(env, wb, i, sheet) {
  const { root, state, options, act, openFilter, fillOptionsMenu, listMenu, chartMenu, showConditional, showSheet, sheetsChanged, alert, followLink, editing, lookChanged, restyled, showClip, lines, headerMenu, cellMenu, filterMenu, askInvalid, showValidation, allowed, openMenu, statusMenu, sheetActs } = env;
    state.grid = createGrid(root, { sheet, label: wb.sheets[i].name, looks: state.looks, readOnly: state.readOnly });
    state.grid.onFilterButton((c, button) => openFilter(c, button.getBoundingClientRect(), button.dataset.table || null));
    state.grid.onErrorButton((point) => act('errorMenu', point));
    state.grid.onOutline((o) => (o.level ? act('outlineLevel', o.axis, o.level) : act('outlineToggle', o.axis, o.line)));
    state.grid.onFillOptions((point) => fillOptionsMenu(point));
    state.grid.onDropdown((button) => listMenu(button.getBoundingClientRect()));
    state.picker = chartPicker({ grid: state.grid, charts: () => state.charts || [], act, menu: (point, k) => chartMenu(point, k), readOnly: () => state.readOnly });
    state.grid.onChartsDrawn(() => state.picker?.redrawn());
    state.grid.setNotes(readNotes(wb, i));
    showConditional();
    state.editing = attachEditing(root, {
      workbook: () => state.workbook,
      sheetIndex: () => state.sheetIndex,
      showSheet,
      sheetsChanged,
      alert,
      // (Excel's Unlock Range, for a change to an Allow Edit Range with a password: rangesbox.js.)
      followLink,
      unlockRange: (name) => unlockRangeBox(root).then((pw) => pw != null && editing().run('unlockEditRange', name, pw)),
      grid: state.grid,
      looks: () => state.looks,
      readOnly: () => state.readOnly,
      changed: () => { showConditional(); options.onChange?.(); },
      edited: (e) => { for (const f of [...state.onEdit]) f(e); },
      lookChanged,
      restyled,
      clipShown: (c) => { state.clip = c; showClip(); },
      // (The last copy made in any editor on the page, and letting one go: pagebooks.js.)
      sharedClip: () => lastCopy(),
      dropShared: (c) => setLastCopy(null, c),
      reloaded: () => { state.grid.setModel(state.workbook.sheet(state.sheetIndex)); lookChanged(); },
      lines,
      headerMenu,
      cellMenu,
      pasteSpecial: () => act('pasteSpecialBox'),
      link: () => act('linkBox'),
      insertTable: () => act('tableInsert'),
      autoSum: (fn) => act('autoSum', fn),
      filterMenu: () => filterMenu() || listMenu(),
      askInvalid: (info) => askInvalid(info),
      selectionShown: () => { for (const f of [...state.onSelect]) f(); return showValidation(); },
      goToBox: () => act('goToBox'),
      refreshData: (all) => act('refreshData', all),
      nameManager: () => act('nameManager'),
      createNames: () => act('createNames'),
      insertFunction: () => act('insertFunction'),
      functionArguments: () => act('functionArguments'),
      pasteNames: () => act('pasteNames'),
      errorMenu: () => { const at = state.grid.errorButtonAt(); return at ? act('errorMenu', at) : false; },
      painting: () => act('painting'),
      commentBox: () => act('commentBox'),
      outline: (dir) => act(dir > 0 ? 'groupBox' : 'ungroupBox'),
      flashFill: () => act('flashFill'),
      saveAs: () => act('saveAsBox'),
      chartKey: (e) => state.picker?.onKey(e),
      formatCells: () => act('formatCellsBox'),
      painted: () => act('paintFormats'),
      drawing: () => act('drawing'),
      drawn: (range, side) => act('drawn', range, side),
      stopDrawing: () => act('stopDrawing'),
      toggleNote: () => { const s = editing().selection; return state.grid.toggleNote(s.ar, s.ac); },
      editNote: () => act('editNote'),
      fitWidth: (c) => state.grid.fitWidth(c),
      fitHeight: (r) => state.grid.fitHeight(r),
      // (A fill handle's fill: Auto Fill Options shown for it.)
      filled: (fill) => { state.lastFill = fill; state.grid.showFillOptions(fill.to); },
      find: (replace) => state.find?.open(replace),
      shownText: (r, c) => state.grid.shownText(r, c),
      allowed,
    });
    state.editing.setReadOnly(state.readOnly);
    state.tabs = createSheetTabs(root, {
      controls: state.grid.el.id,
      on: {
        select: (k) => editing().run('showSheet', k),
        menu: openMenu,
        statusMenu,
        rename: sheetActs.rename,
        move: (from, to) => editing().run('moveSheet', from, to),
        add: () => editing().run('insertSheet', { after: true }),
        keepFocus: () => !!state.editing?.keepsFocus(),
      },
    });
}
