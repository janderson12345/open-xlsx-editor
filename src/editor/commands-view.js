// The commands on the sheet's view (commands.js shares its state with them): Freeze Panes, Show Formulas, the zoom
// and View > Show (Gridlines, Formula Bar, Headings). View changes, kept in the file (all but the Formula Bar, which
// Excel keeps as its own setting, not the workbook's) and not in the undo history, as in Excel.

/** The view commands: `env` is createCommands's shared state and helpers. */
export function viewCommands(env) {
  const { commands, ctx, state, sheet } = env;
  return {
    /** Freeze Panes ('panes': above and left of the active cell; again when frozen, Unfreeze Panes, as Excel's menu
     *  turns), Freeze Top Row ('topRow'), Freeze First Column ('firstColumn') or Unfreeze ('none'). A view change, kept
     *  in the file but not in the undo history, as in Excel. */
    freeze(kind) {
      const wb = ctx.workbook();
      const si = ctx.sheetIndex();
      const now = wb.sheet(si).freeze || { rows: 0, cols: 0 };
      const s = state.sel;
      let freeze = { rows: 0, cols: 0 };
      if (kind === 'topRow') freeze = { rows: 1, cols: 0 };
      else if (kind === 'firstColumn') freeze = { rows: 0, cols: 1 };
      else if (kind === 'panes' && !(now.rows || now.cols)) {
        // (Excel freezes the middle of the window when A1 is active; here that needs another cell.)
        if (!s.ar && !s.ac) { ctx.announce('Select the cell below and to the right of what should stay in place'); return false; }
        freeze = { rows: s.ar, cols: s.ac };
      }
      const topLeft = ctx.grid.view().topLeft;
      wb.setView(si, { freeze, active: { r: s.ar, c: s.ac }, topLeft });
      ctx.grid.setView({ freeze });
      ctx.changed?.();
      ctx.announce(freeze.rows || freeze.cols ? `Frozen: ${freeze.rows} row${freeze.rows === 1 ? '' : 's'} and ${freeze.cols} column${freeze.cols === 1 ? '' : 's'}` : 'Panes unfrozen');
      return true;
    },
    /** Formulas > Show Formulas (Ctrl+`): on, off, or turned over (no argument); a view change, not undone. */
    showFormulas(on) {
      const wb = ctx.workbook();
      const si = ctx.sheetIndex();
      const now = on ?? !wb.sheet(si).showFormulas;
      wb.setView(si, { showFormulas: now });
      ctx.reloaded?.();
      ctx.grid.scrollIntoView(state.focus.r, state.focus.c);
      ctx.changed?.();
      ctx.announce(now ? 'Showing formulas' : 'Showing values');
      return true;
    },
    /** The sheet's zoom in percent (10 to 400, as Excel allows); a view change, not undone. */
    zoom(percent) {
      const z = Math.round(Math.min(400, Math.max(10, percent)));
      ctx.workbook().setView(ctx.sheetIndex(), { zoom: z });
      ctx.grid.setView({ zoom: z });
      ctx.grid.scrollIntoView(state.focus.r, state.focus.c);
      ctx.changed?.();
      ctx.announce(`Zoom ${z}%`);
      return true;
    },
    /** Zoom to Selection: the zoom at which the selection fills the window (cut to the used range when it runs past
     *  it), as near as whole percent allows. */
    zoomToSelection() {
      const s = state.sel;
      const sh = sheet();
      const r1 = s.r1 - s.r0 > 10000 ? Math.max(s.r0, Math.min(s.r1, sh.maxRow)) : s.r1;
      const c1 = s.c1 - s.c0 > 1000 ? Math.max(s.c0, Math.min(s.c1, sh.maxCol)) : s.c1;
      const g = ctx.grid.geometry();
      const w = g.cols.start(c1 + 1) - g.cols.start(s.c0);
      const h = g.rows.start(r1 + 1) - g.rows.start(s.r0);
      const view = ctx.grid.el;
      // (The headings scale with the sheet, so they count with it.)
      const z = Math.floor(Math.min(view.clientWidth / (Math.max(1, w) + g.rowHeadW), view.clientHeight / (Math.max(1, h) + g.colHeadH)) * 100);
      commands.zoom(z);
      ctx.grid.scrollToCell(s.r0, s.c0);
      return true;
    },
    /** View > Gridlines: on, off, or turned over (no argument), on the sheet shown. */
    showGridlines(on) {
      const now = on ?? !sheet().showGridLines;
      ctx.workbook().setView(ctx.sheetIndex(), { showGridLines: now });
      ctx.reloaded?.();
      ctx.changed?.();
      ctx.announce(now ? 'Gridlines shown' : 'Gridlines hidden');
      return true;
    },
    /** View > Headings: the row numbers and column letters on, off, or turned over (no argument), on the sheet shown. */
    showHeadings(on) {
      const now = on ?? sheet().showHeadings === false;
      ctx.workbook().setView(ctx.sheetIndex(), { showHeadings: now });
      ctx.reloaded?.();
      ctx.grid.scrollIntoView(state.focus.r, state.focus.c);
      ctx.changed?.();
      ctx.announce(now ? 'Headings shown' : 'Headings hidden');
      return true;
    },
    /** View > Formula Bar: on, off, or turned over (no argument). The editor's, for every sheet; not in the file. */
    showFormulaBar(on) {
      const now = on ?? !ctx.formulaBar();
      ctx.formulaBar(now);
      ctx.changed?.();
      ctx.announce(now ? 'Formula bar shown' : 'Formula bar hidden');
      return true;
    },
  };
}
