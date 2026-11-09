// Keyboard and mouse on the grid, as Excel answers them: arrows (Shift extends, Ctrl jumps), Tab, Enter, Home, End,
// Page Up and Down (with Ctrl, the next or previous sheet), F2 and double click to edit, typing to replace, Delete to
// empty, Ctrl+Z and Ctrl+Y; click and drag to select. Everything goes through the named commands; the cell editor and
// formula bar enter what is typed.
import { createCommands } from './commands.js';
import { createCellEditor } from './celleditor.js';
import { createFormulaBar } from './formulabar.js';
import { announce } from './live.js';
import { refName } from '../ooxml/refs.js';
import { cellText } from '../ooxml/numfmt.js';
import { editText } from '../ooxml/input.js';

const ARROWS = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };

/** Make the grid in `root` editable. `ctx` = { workbook(), sheetIndex(), showSheet(i), grid, looks(), readOnly() }. */
export function attachEditing(root, ctx) {
  const { grid } = ctx;
  const view = grid.el;
  let cmd = null;
  const cellAt = (r, c) => ctx.workbook().sheet(ctx.sheetIndex()).cells.get(r)?.get(c) || null;
  const bar = createFormulaBar(root, {
    onEnter: (text) => { cmd.run('enter', text); view.focus(); },
    onCancel: () => { refreshBar(); view.focus(); },
  });
  const editor = createCellEditor({
    grid,
    onFinish: (text, move) => {
      // (An entry refused, a formula with a mistake in it, goes back into the editor to be put right, as in Excel.)
      if (!cmd.run('enter', text)) { reopen(text); return; }
      if (move) cmd.run('move', move[0], move[1]);
      view.focus({ preventScroll: true });
    },
    onCancel: () => view.focus({ preventScroll: true }),
  });
  cmd = createCommands({
    ...ctx,
    announce,
    selected: (sel, note) => {
      refreshBar();
      const cell = cellAt(sel.ar, sel.ac);
      const range = sel.r0 !== sel.r1 || sel.c0 !== sel.c1 ? `, ${refName(sel.r0, sel.c0)} to ${refName(sel.r1, sel.c1)} selected` : '';
      announce(`${note}${refName(sel.ar, sel.ac)}${cell && cell.v != null ? `, ${cellText(cell)}` : ', blank'}${range}`);
    },
    showSheet: (i) => { editor.close(); ctx.showSheet(i); },
    changed: () => refreshBar(),
  });
  function refreshBar() {
    const s = cmd ? cmd.selection : { ar: 0, ac: 0 };
    bar.show(refName(s.ar, s.ac), editText(cellAt(s.ar, s.ac)));
  }
  const reopen = (text) => {
    const s = cmd.selection;
    const cell = cellAt(s.ar, s.ac);
    editor.open(s.ar, s.ac, text, 'edit', ctx.looks().classOf ? ctx.looks().classOf(cell?.s || 0) : '');
  };
  const startEdit = (mode, seed = null) => {
    if (ctx.readOnly()) { announce('The workbook is read-only'); return; }
    const s = cmd.selection;
    const cell = cellAt(s.ar, s.ac);
    if (ctx.workbook().sheet(ctx.sheetIndex()).arrays.some((a) => s.ar >= a.r0 && s.ar <= a.r1 && s.ac >= a.c0 && s.ac <= a.c1)) {
      announce("You can't change part of an array");
      return;
    }
    const cls = ctx.looks().classOf ? ctx.looks().classOf(cell?.s || 0) : '';
    editor.open(s.ar, s.ac, seed != null ? seed : editText(cell), mode, cls);
  };

  view.addEventListener('keydown', (e) => {
    if (e.target !== view) return;
    const ctrl = e.ctrlKey || e.metaKey;
    const key = e.key;
    let handled = true;
    if (ARROWS[key]) cmd.run('move', ARROWS[key][0], ARROWS[key][1], { extend: e.shiftKey, jump: ctrl });
    else if (key === 'Tab') cmd.run('move', 0, e.shiftKey ? -1 : 1);
    else if (key === 'Enter') cmd.run('move', e.shiftKey ? -1 : 1, 0);
    else if (key === 'Home') cmd.run('home', { sheetStart: ctrl });
    else if (key === 'End' && ctrl) cmd.run('end');
    else if ((key === 'PageDown' || key === 'PageUp') && ctrl) cmd.run('nextSheet', key === 'PageDown' ? 1 : -1);
    else if (key === 'PageDown' || key === 'PageUp') cmd.run('page', key === 'PageDown' ? 1 : -1, { extend: e.shiftKey });
    else if (key === 'F2') startEdit('edit');
    else if (key === 'Delete') cmd.run('clear');
    else if (key === 'Backspace') startEdit('enter', '');
    else if (ctrl && (key === 'z' || key === 'Z') && !e.shiftKey) cmd.run('undo');
    else if (ctrl && (key === 'y' || ((key === 'z' || key === 'Z') && e.shiftKey))) cmd.run('redo');
    else if (key.length === 1 && !ctrl && !e.altKey) startEdit('enter', key);
    else handled = false;
    if (handled) e.preventDefault();
  });

  let dragging = false;
  view.addEventListener('mousedown', (e) => {
    if (e.button !== 0 || e.target.closest('.oxe-editor')) return;
    const at = grid.hitTest(e.clientX, e.clientY);
    if (!at || at.r < 0 || at.c < 0) return;
    e.preventDefault();
    view.focus({ preventScroll: true });
    if (e.shiftKey) cmd.run('extendTo', at.r, at.c);
    else cmd.run('select', at.r, at.c);
    // (A double click edits. Taken from the click count: each click redraws the cells, so the browser's dblclick,
    // which wants both clicks on one element, doesn't come.)
    if (e.detail === 2 && !e.shiftKey) { startEdit('edit'); return; }
    dragging = true;
  });
  const onMove = (e) => {
    if (!dragging) return;
    const at = grid.hitTest(e.clientX, e.clientY);
    if (at && at.r >= 0 && at.c >= 0) cmd.run('extendTo', at.r, at.c);
  };
  const onUp = () => { dragging = false; };
  window.addEventListener('mousemove', onMove);
  window.addEventListener('mouseup', onUp);

  return {
    commands: cmd,
    editor,
    bar,
    setReadOnly(flag) { bar.setReadOnly(flag); grid.setReadOnly(flag); if (flag) editor.close(); },
    reset() { editor.close(); cmd.reset(); },
    destroy() {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    },
  };
}
