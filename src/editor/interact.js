// Keyboard and mouse on the grid, as Excel answers them: arrows (Shift extends, Ctrl jumps), Tab, Enter, Home, End,
// Page Up and Down (with Ctrl, the next or previous sheet), F2 and double click to edit, typing to replace, Delete to
// empty, Ctrl+Z and Ctrl+Y, Shift+F11 for a new sheet; click and drag to select. Everything goes through the named
// commands; the cell editor and formula bar enter what is typed. While a formula is typed, the arrows, clicks and drags
// point at cells (pointing.js); a sheet tab clicked then shows that sheet and the formula goes on in the formula bar,
// entered into its own cell, on its own sheet, with Enter (as in Excel).
import { createCommands } from './commands.js';
import { createCellEditor } from './celleditor.js';
import { createFormulaBar } from './formulabar.js';
import { createPointing } from './pointing.js';
import { attachFunctionTip } from './fntip.js';
import { announce } from './live.js';
import { refName, MAX_ROWS, MAX_COLS } from '../ooxml/refs.js';
import { cellText } from '../ooxml/numfmt.js';
import { editText } from '../ooxml/input.js';
import { legacyAt } from '../calc/legacyat.js';
import { formatCodeOf } from '../ooxml/styles.js';

const FONT_KEYS = { b: 'b', 2: 'b', i: 'i', 3: 'i', u: 'u', 4: 'u', 5: 'strike' };
const ARROWS = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };

/** Make the grid in `root` editable. `ctx` = { workbook(), sheetIndex(), showSheet(i), grid, looks(), readOnly() }. */
export function attachEditing(root, ctx) {
  const { grid } = ctx;
  const view = grid.el;
  let cmd = null;
  const cellAt = (r, c) => ctx.workbook().sheet(ctx.sheetIndex()).cells.get(r)?.get(c) || null;
  // (A formula carried to the bar while another sheet is shown: { si, r, c }, the cell it is for.)
  let carried = null;
  const bar = createFormulaBar(root, {
    onEnter: (text) => {
      if (carried) { finishCarried(text); return; }
      commit(text, null, () => view.focus());
    },
    onCancel: () => {
      if (carried) { const o = carried; carried = null; pointing.attach(null); cmd.run('showSheet', o.si); }
      refreshBar();
      view.focus();
    },
    onGo: (text) => {
      if (cmd.run('nameBox', text) !== false) view.focus();
    },
    onKey: (e) => pointing.key(e),
    onFocus: () => pointing.attach(bar.input),
    onBlur: () => {
      // (Leaving the bar while a formula is carried enters it, as leaving the cell editor does.)
      if (carried) finishCarried(bar.input.value);
      else if (pointing.control === bar.input) pointing.attach(null);
    },
  });
  const editor = createCellEditor({
    grid,
    onKey: (e) => {
      if ((e.ctrlKey || e.metaKey) && (e.key === 'PageDown' || e.key === 'PageUp') && pointing.ready) { cmd.run('nextSheet', e.key === 'PageDown' ? 1 : -1); return true; }
      // (Shift+F3: Insert Function, its choice typed at the caret.)
      if (e.key === 'F3' && e.shiftKey && !e.ctrlKey && !e.altKey && ctx.insertFunction) { ctx.insertFunction(); return true; }
      return pointing.key(e);
    },
    onClose: () => { if (pointing.control === editor.el) pointing.attach(null); },
    onFinish: (text, move) => {
      // (An entry refused, a formula with a mistake in it, goes back into the editor to be put right, as in Excel.)
      commit(text, null, () => {
        if (move) cmd.run('move', move[0], move[1]);
        view.focus({ preventScroll: true });
      });
    },
    onCancel: () => view.focus({ preventScroll: true }),
  });
  // (Inside a function's parentheses, its syntax under the cell, as Excel shows it.)
  attachFunctionTip(editor.el, grid.overlay);
  cmd = createCommands({
    ...ctx,
    announce,
    selected: (sel, note) => {
      refreshBar();
      ctx.lookChanged?.();
      const cell = cellAt(sel.ar, sel.ac);
      const range = sel.r0 !== sel.r1 || sel.c0 !== sel.c1 ? `, ${refName(sel.r0, sel.c0)} to ${refName(sel.r1, sel.c1)} selected` : '';
      // (A data validation input message is read with the cell, as it shows beside it.)
      const prompt = ctx.selectionShown?.() || '';
      announce(`${note}${refName(sel.ar, sel.ac)}${cell && cell.v != null ? `, ${cellText(cell)}` : ', blank'}${range}${prompt}`);
    },
    showSheet: (i) => {
      if (pointing.ready && !carried) carry();
      else if (editor.isOpen) editor.commit();
      ctx.showSheet(i);
      if (carried) pointing.sheetShown();
    },
    changed: () => { refreshBar(); ctx.lookChanged?.(); ctx.changed?.(); },
  });
  const pointing = createPointing({
    grid,
    workbook: ctx.workbook,
    shown: ctx.sheetIndex,
    origin: () => {
      if (carried) return carried;
      const at = editor.at || { r: cmd.selection.ar, c: cmd.selection.ac };
      return { si: ctx.sheetIndex(), ...at };
    },
    active: () => ({ r: cmd.selection.ar, c: cmd.selection.ac }),
    say: announce,
    keys: () => (pointing.control === editor.el && editor.mode === 'enter') || (pointing.control === bar.input && !!carried),
  });
  /** Carry the formula being typed to the bar: another sheet is about to be shown. */
  function carry() {
    const box = pointing.control;
    const text = box.value;
    const caret = box.selectionStart;
    carried = box === editor.el ? { si: ctx.sheetIndex(), ...editor.at } : { si: ctx.sheetIndex(), r: cmd.selection.ar, c: cmd.selection.ac };
    if (box === editor.el) editor.close();
    bar.input.value = text;
    bar.input.focus();
    bar.input.setSelectionRange(caret, caret);
    pointing.attach(bar.input);
  }
  /** Enter a carried formula into its cell, on its sheet (shown again). Refused, it goes back into the cell's editor. */
  function finishCarried(text) {
    const o = carried;
    carried = null;
    pointing.attach(null);
    cmd.run('showSheet', o.si);
    cmd.run('select', o.r, o.c);
    commit(text, { r: o.r, c: o.c }, () => { view.focus({ preventScroll: true }); refreshBar(); });
  }
  /** Enter `text` (into `at`, or the active cell), then `after()`. Refused (a formula with a mistake), it goes back into
   *  the editor; refused by the cell's data validation, Excel's question first (`ctx.askInvalid`): Retry (or No) back
   *  into the editor, Yes or OK entered anyway, Cancel dropped. */
  async function commit(text, at, after) {
    const done = cmd.run('enter', text, at);
    if (done && done.invalid) {
      const how = (await ctx.askInvalid?.(done.invalid)) || 'cancel';
      if (how === 'force') { cmd.run('enter', text, at, { force: true }); after(); }
      else if (how === 'retry') reopen(text);
      else { refreshBar(); view.focus({ preventScroll: true }); }
      return;
    }
    if (!done) { reopen(text); return; }
    after();
  }
  /** A cell's text as Excel's formula bar has it (a date as a date, a percentage as one). */
  // (An old formula with the @s Excel 365 shows where it cuts a range to one value, at its cell: calc/legacyat.js.)
  const barText = (cell, r, c) => editText(cell, formatCodeOf(ctx.workbook().styles, cell?.s || 0), ctx.workbook().date1904,
    (shown) => legacyAt(ctx.workbook().calc, ctx.sheetIndex(), r, c, shown));
  function refreshBar() {
    const s = cmd ? cmd.selection : { ar: 0, ac: 0 };
    bar.show(refName(s.ar, s.ac), barText(cellAt(s.ar, s.ac), s.ar, s.ac));
  }
  const reopen = (text) => {
    const s = cmd.selection;
    const cell = cellAt(s.ar, s.ac);
    editor.open(s.ar, s.ac, text, 'edit', ctx.looks().classOf ? ctx.looks().classOf(cell?.s || 0) : '');
    pointing.attach(editor.el);
  };
  const startEdit = (mode, seed = null) => {
    if (ctx.readOnly()) { announce('The workbook is read-only'); return; }
    cmd.run('dropClip'); // (typing lets a copy go, as in Excel)
    const s = cmd.selection;
    const cell = cellAt(s.ar, s.ac);
    if (ctx.workbook().sheet(ctx.sheetIndex()).arrays.some((a) => s.ar >= a.r0 && s.ar <= a.r1 && s.ac >= a.c0 && s.ac <= a.c1)) {
      announce("You can't change part of an array");
      return;
    }
    const cls = ctx.looks().classOf ? ctx.looks().classOf(cell?.s || 0) : '';
    editor.open(s.ar, s.ac, seed != null ? seed : barText(cell, s.ar, s.ac), mode, cls);
    pointing.attach(editor.el);
  };

  view.addEventListener('keydown', (e) => {
    if (e.target !== view) return;
    // (A chart picked, or Ctrl+Alt+5 to pick one: its keys first, chartpick.js.)
    if (ctx.chartKey?.(e)) { e.preventDefault(); return; }
    const ctrl = e.ctrlKey || e.metaKey;
    const key = e.key;
    let handled = true;
    // (Alt+Down on a heading cell of a filter opens its list, as in Excel.)
    if (e.altKey && !ctrl && key === 'ArrowDown' && ctx.filterMenu?.()) { /* opened */ }
    else if (ARROWS[key]) cmd.run('move', ARROWS[key][0], ARROWS[key][1], { extend: e.shiftKey, jump: ctrl });
    else if (key === 'Tab') cmd.run('move', 0, e.shiftKey ? -1 : 1);
    // (Format Painter waiting: Enter paints the selection, as a click does.)
    else if (key === 'Enter' && ctx.painting?.()) ctx.painted();
    else if (key === 'Enter') cmd.run('move', e.shiftKey ? -1 : 1, 0);
    else if (key === 'Home') cmd.run('home', { sheetStart: ctrl });
    else if (key === 'End' && ctrl) cmd.run('end');
    else if ((key === 'PageDown' || key === 'PageUp') && ctrl) cmd.run('nextSheet', key === 'PageDown' ? 1 : -1);
    else if (key === 'PageDown' || key === 'PageUp') cmd.run('page', key === 'PageDown' ? 1 : -1, { extend: e.shiftKey });
    else if (key === 'F5' && e.altKey && !e.shiftKey) cmd.run('refreshPivots', ctrl);
    else if ((key === 'F5' && !ctrl && !e.shiftKey && !e.altKey) || (ctrl && !e.shiftKey && !e.altKey && key.toLowerCase() === 'g')) ctx.goToBox?.();
    // (Alt+Shift+F10: the error button's menu, on a cell with an error indicator.)
    else if (key === 'F10' && e.altKey && e.shiftKey) ctx.errorMenu?.();
    // (Ctrl+`: Show Formulas, on or off.)
    else if (ctrl && !e.altKey && !e.shiftKey && e.code === 'Backquote') cmd.run('showFormulas');
    // (F9 Calculate Now, Shift+F9 Calculate Sheet, Ctrl+Alt+F9 every formula.)
    else if (key === 'F9') cmd.run('calculate', ctrl && e.altKey ? 'all' : e.shiftKey ? 'sheet' : 'now');
    else if (key === 'F3' && ctrl && !e.shiftKey && !e.altKey) ctx.nameManager?.();
    else if (key === 'F3' && e.shiftKey && !ctrl && !e.altKey) ctx.insertFunction?.();
    // (Ctrl+Alt+M: New Comment, or the active cell's thread to reply to.)
    else if (ctrl && e.altKey && !e.shiftKey && key.toLowerCase() === 'm') ctx.commentBox?.();
    // (Ctrl+E: Flash Fill.)
    else if (ctrl && !e.altKey && !e.shiftKey && key.toLowerCase() === 'e') ctx.flashFill?.();
    // (Ctrl+1: Format Cells. Taken from the key's code, as Ctrl+Shift+1 is "!".)
    else if (ctrl && !e.altKey && !e.shiftKey && e.code === 'Digit1') ctx.formatCells?.();
    // (Shift+F2: Excel's New Note, or Edit Note on a cell with one.)
    else if (key === 'F2' && e.shiftKey) ctx.editNote?.();
    else if (key === 'F2') startEdit('edit');
    else if (key === 'F11' && e.shiftKey) cmd.run('insertSheet');
    else if (key === 'Delete') cmd.run('clear');
    else if (key === 'Escape') { if (cmd.run('dropClip')) announce('Copy cancelled'); }
    else if (key === 'Backspace') startEdit('enter', '');
    // (Shift+Space the rows, Ctrl+Space the columns, Ctrl+A the sheet; Ctrl++ inserts, Ctrl+- deletes, as in Excel.)
    else if (key === ' ' && e.shiftKey && !ctrl) cmd.run('selectRows');
    else if (key === ' ' && ctrl) cmd.run('selectColumns');
    else if (ctrl && (key === 'a' || key === 'A') && !e.shiftKey) cmd.run('selectAll');
    else if (ctrl && (key === '+' || (key === '=' && e.shiftKey))) ctx.lines?.(1);
    else if (ctrl && key === '-') ctx.lines?.(-1);
    else if (ctrl && !e.shiftKey && (key === 'f' || key === 'h')) ctx.find?.(key === 'h');
    else if (ctrl && !e.shiftKey && !e.altKey && (key.toLowerCase() === 'd' || key.toLowerCase() === 'r')) { if (ctx.allowed?.('fill') !== false) cmd.run('fill', key.toLowerCase() === 'd' ? 'down' : 'right'); }
    else if (ctrl && e.altKey && !e.shiftKey && key.toLowerCase() === 'v') ctx.pasteSpecial?.();
    else if (e.altKey && !ctrl && key === '=') ctx.autoSum?.('sum');
    else if (ctrl && e.shiftKey && !e.altKey && key.toLowerCase() === 'l') { if (ctx.allowed?.('filterToggle') !== false) cmd.run('filterToggle'); }
    // (Ctrl+9 hides the rows, Ctrl+0 the columns; with Shift they show again. Taken from the key's code: Shift+9 is "(".)
    else if (ctrl && (e.code === 'Digit9' || e.code === 'Digit0')) cmd.run('hideLines', e.code === 'Digit9' ? 'row' : 'col', !e.shiftKey);
    else if (ctrl && (key === 'z' || key === 'Z') && !e.shiftKey) cmd.run('undo');
    else if (ctrl && (key === 'y' || ((key === 'z' || key === 'Z') && e.shiftKey))) cmd.run('redo');
    // (Ctrl+C, X and V: the browser's copy and paste go on after the key, so its default isn't stopped.)
    else if (ctrl && !e.altKey && !e.shiftKey && /^[cxv]$/i.test(key)) { if (key.toLowerCase() === 'v') keyPaste(); else keyCopy(key.toLowerCase() === 'x'); handled = false; }
    // (Excel's font shortcuts: Ctrl+B or 2 bold, I or 3 italic, U or 4 underline, 5 strikethrough.)
    else if (ctrl && !e.shiftKey && !e.altKey && FONT_KEYS[key.toLowerCase()]) { if (ctx.allowed?.('toggleFont') !== false) cmd.run('toggleFont', FONT_KEYS[key.toLowerCase()]); }
    else if (key.length === 1 && !ctrl && !e.altKey) startEdit('enter', key);
    else handled = false;
    if (handled) e.preventDefault();
  });

  // Copy, cut and paste through the browser's clipboard events (Ctrl+C, X, V, and the browser's own menu): the cells go
  // on the system clipboard as text, tabs between cells and a line for each row, as Excel puts them.
  /** The text of the cells a copy holds, as the grid shows them. */
  const clipText = (clip) => {
    const lines = [];
    for (let r = clip.rect.r0; r <= clip.rect.r1; r++) {
      const line = [];
      for (let c = clip.rect.c0; c <= clip.rect.c1; c++) line.push(grid.shownText(r, c) ?? '');
      lines.push(line.join('\t'));
    }
    return `${lines.join('\r\n')}\r\n`;
  };
  const norm = (t) => t.replace(/\r\n?/g, '\n').replace(/\n$/, '');
  /** Copy or cut the selection: kept here, and its text returned for the system clipboard (null if refused). */
  const takeClip = (cut) => {
    const clip = cmd.run(cut ? 'cut' : 'copy');
    if (!clip) return null;
    const text = clipText(clip);
    clip.text = norm(text);
    announce(`${cut ? 'Cut' : 'Copied'} ${refName(clip.rect.r0, clip.rect.c0)}${clip.rect.r1 > clip.rect.r0 || clip.rect.c1 > clip.rect.c0 ? ` to ${refName(clip.rect.r1, clip.rect.c1)}` : ''}`);
    return text;
  };
  const pasteText = (text) => cmd.run('paste', text == null ? null : norm(text));
  // Ctrl+C and Ctrl+X take the cells at the key, and the copy event that follows puts their text on the clipboard (a
  // browser may send none to a sheet with no text selected: the async clipboard is tried too). Ctrl+V waits a moment
  // for the paste event; without one it reads the async clipboard, or pastes what was copied here.
  let pending = null;
  let pasteTimer = 0;
  const keyCopy = (cut) => {
    if (ctx.allowed?.('copy') === false) return;
    pending = takeClip(cut);
    if (pending != null) navigator.clipboard?.writeText(pending).catch(() => {});
  };
  const keyPaste = () => {
    if (ctx.allowed?.('paste') === false) return;
    clearTimeout(pasteTimer);
    pasteTimer = setTimeout(async () => {
      let text = null;
      try { text = await navigator.clipboard.readText(); } catch { /* not allowed: what was copied here */ }
      pasteText(text);
    }, 80);
  };
  const onCopy = (e) => {
    if (document.activeElement !== view || ctx.allowed?.('copy') === false) return;
    const text = pending ?? takeClip(e.type === 'cut');
    pending = null;
    if (text == null) return;
    e.clipboardData.setData('text/plain', text);
    e.preventDefault();
  };
  const onPaste = (e) => {
    if (document.activeElement !== view || ctx.allowed?.('paste') === false) return;
    clearTimeout(pasteTimer);
    e.preventDefault();
    pasteText(e.clipboardData.getData('text/plain'));
  };
  document.addEventListener('copy', onCopy);
  document.addEventListener('cut', onCopy);
  document.addEventListener('paste', onPaste);

  let dragging = false;
  let pointingDrag = false;
  // The fill handle: dragged out of the selection down, up, right or left (the way the pointer is furthest out), the
  // cells it will fill outlined; let go, they are filled (Ctrl held: the other way of filling).
  let filling = null;
  const fillTarget = (at) => {
    const s = cmd.selection;
    const down = at.r - s.r1;
    const up = s.r0 - at.r;
    const right = at.c - s.c1;
    const left = s.c0 - at.c;
    const most = Math.max(down, up, right, left);
    if (most <= 0) return null;
    if (most === down) return { ...s, r1: at.r };
    if (most === up) return { ...s, r0: at.r };
    if (most === right) return { ...s, c1: at.c };
    return { ...s, c0: at.c };
  };
  view.addEventListener('mousedown', (e) => {
    if (e.button !== 0 || e.target.closest('.oxe-editor') || pointing.control || !grid.fillHandleAt(e.clientX, e.clientY)) return;
    e.preventDefault();
    e.stopImmediatePropagation();
    view.focus({ preventScroll: true });
    filling = { to: null };
  }, true);
  view.addEventListener('mousedown', (e) => {
    if (e.button !== 0 || e.target.closest('.oxe-editor')) return;
    const at = grid.hitTest(e.clientX, e.clientY);
    if (!at) return;
    // (A row or column heading selects the whole line, Shift extends; the corner selects the sheet.)
    if ((at.r < 0 || at.c < 0) && !pointing.control) {
      e.preventDefault();
      view.focus({ preventScroll: true });
      if (at.r < 0 && at.c < 0) cmd.run('selectAll');
      else if (at.r < 0) cmd.run('selectLines', 'col', at.c, e.shiftKey);
      else cmd.run('selectLines', 'row', at.r, e.shiftKey);
      return;
    }
    if (at.r < 0 || at.c < 0) return;
    e.preventDefault();
    // (A formula being typed where a reference can go: the cell is pointed at, and the text box keeps focus.)
    if (pointing.control && pointing.press(at.r, at.c, e.shiftKey)) { pointingDrag = true; return; }
    view.focus({ preventScroll: true });
    if (e.shiftKey) cmd.run('extendTo', at.r, at.c);
    else if (e.ctrlKey || e.metaKey) cmd.run('addArea', at.r, at.c); // (Ctrl: another range, dragged out as one is)
    else cmd.run('select', at.r, at.c);
    // (A double click edits. Taken from the click count: each click redraws the cells, so the browser's dblclick,
    // which wants both clicks on one element, doesn't come.)
    if (e.detail === 2 && !e.shiftKey) { startEdit('edit'); return; }
    dragging = true;
  });
  // Right-click on a row or column heading: its menu (Insert, Delete, Clear Contents), the line selected first unless
  // it is already part of the selection, as in Excel. On a cell, the cell's menu (Cut, Copy, Paste, Paste Special...,
  // Insert..., Delete..., Clear Contents), the cell selected first unless it is in the selection.
  view.addEventListener('contextmenu', (e) => {
    const at = grid.hitTest(e.clientX, e.clientY);
    if (at && at.r >= 0 && at.c >= 0) {
      e.preventDefault();
      const s = cmd.selection;
      if (!(at.r >= s.r0 && at.r <= s.r1 && at.c >= s.c0 && at.c <= s.c1)) cmd.run('select', at.r, at.c);
      ctx.cellMenu?.({ x: e.clientX, y: e.clientY });
      return;
    }
    if (!at || (at.r < 0 && at.c < 0)) return;
    e.preventDefault();
    const axis = at.r < 0 ? 'col' : 'row';
    const s = cmd.selection;
    const whole = axis === 'col' ? s.r0 === 0 && s.r1 >= MAX_ROWS - 1 && at.c >= s.c0 && at.c <= s.c1 : s.c0 === 0 && s.c1 >= MAX_COLS - 1 && at.r >= s.r0 && at.r <= s.r1;
    if (!whole) cmd.run('selectLines', axis, axis === 'col' ? at.c : at.r);
    ctx.headerMenu?.(axis, { x: e.clientX, y: e.clientY });
  });
  // A heading edge: the resize cursor over it; dragged, the row or column resized (a guide and Excel's tip while
  // dragging); double-clicked, AutoFit. Captured before the heading's own selecting.
  let resizing = null;
  view.addEventListener('mousedown', (e) => {
    if (e.button !== 0 || pointing.control) return;
    const edge = grid.edgeAt(e.clientX, e.clientY);
    if (!edge) return;
    e.preventDefault();
    e.stopImmediatePropagation();
    view.focus({ preventScroll: true });
    if (ctx.readOnly()) return;
    if (e.detail === 2) { cmd.run('fitLines', edge.axis, edge.index); return; }
    resizing = { ...edge, from: edge.axis === 'col' ? e.clientX : e.clientY, now: edge.size };
    grid.showGuide(edge.axis, edge.index, edge.size);
  }, true);
  // Ctrl+wheel zooms the sheet by 10% a notch.
  view.addEventListener('wheel', (e) => {
    if (!(e.ctrlKey || e.metaKey) || !e.deltaY) return;
    e.preventDefault();
    if (ctx.allowed?.('zoom') === false) return;
    cmd.run('zoom', grid.view().zoom + (e.deltaY < 0 ? 10 : -10));
  }, { passive: false });
  view.addEventListener('mousemove', (e) => {
    if (resizing || dragging || pointingDrag || filling) return;
    const edge = e.buttons ? null : grid.edgeAt(e.clientX, e.clientY);
    view.classList.toggle('oxe-resize-col', edge?.axis === 'col');
    view.classList.toggle('oxe-resize-row', edge?.axis === 'row');
    view.classList.toggle('oxe-fill-cursor', !edge && !e.buttons && grid.fillHandleAt(e.clientX, e.clientY));
  });
  const onMove = (e) => {
    if (filling) {
      const at = grid.hitTest(e.clientX, e.clientY);
      if (!at || at.r < 0 || at.c < 0) return;
      filling.to = fillTarget(at);
      grid.showFill(filling.to);
      return;
    }
    if (resizing) {
      resizing.now = Math.max(0, resizing.size + (resizing.axis === 'col' ? e.clientX : e.clientY) - resizing.from);
      grid.showGuide(resizing.axis, resizing.index, resizing.now);
      return;
    }
    if (!dragging && !pointingDrag) return;
    const at = grid.hitTest(e.clientX, e.clientY);
    if (!at || at.r < 0 || at.c < 0) return;
    if (pointingDrag) pointing.drag(at.r, at.c);
    else cmd.run('extendTo', at.r, at.c);
  };
  const onUp = (e) => {
    if (filling) {
      const { to } = filling;
      filling = null;
      grid.showFill(null);
      if (to) cmd.run('autoFill', to, { ctrl: e.ctrlKey || e.metaKey });
      return;
    }
    if (resizing) {
      const r = resizing;
      resizing = null;
      grid.showGuide(null);
      if (Math.round(r.now) !== Math.round(r.size)) cmd.run('resizeLines', r.axis, r.index, r.now);
    }
    // (Format Painter waiting: the cells just selected get the copied formats.)
    const selected = dragging;
    dragging = false;
    pointingDrag = false;
    if (selected && ctx.painting?.()) ctx.painted();
  };
  window.addEventListener('mousemove', onMove);
  window.addEventListener('mouseup', onUp);

  return {
    commands: cmd,
    /** The active cell opened for editing with `text` (AutoSum's proposal, waiting for Enter as in Excel). */
    editWith(text) { startEdit('enter', text); },
    /** Insert Function's choice: typed at the caret of the cell being edited, or the active cell opened with
     *  "=NAME(". `hold` keeps the editor open while the box is up. */
    get editorOpen() { return editor.isOpen; },
    holdEditor(on) { editor.hold(on); },
    typeFunction(name) {
      if (!editor.isOpen) { startEdit('enter', `=${name}(`); return; }
      editor.insert(editor.el.value === '' ? `=${name}(` : `${name}(`);
    },
    /** The ribbon's Copy and Cut: the clipboard written through the async API (a button has no clipboard event). */
    copyOut(cut = false) {
      const text = takeClip(cut);
      if (text != null) navigator.clipboard?.writeText(text).catch(() => {});
      view.focus({ preventScroll: true });
    },
    /** Format Painter's copy: the selection kept, outlined, as Copy keeps it but leaving the system clipboard
     *  alone. False if it couldn't be copied. */
    pickFormats() { return takeClip(false) != null; },
    /** The ribbon's Paste: the system clipboard's text if the browser lets it be read, else what was copied here. */
    async pasteIn() {
      let text = null;
      try { text = await navigator.clipboard.readText(); } catch { /* not allowed: what was copied here */ }
      if (text == null && !cmd.selection) return;
      if (!pasteText(text) && text == null) announce('Use Ctrl+V to paste from another program');
      view.focus({ preventScroll: true });
    },
    editor,
    bar,
    /** Whether focus should stay where a formula is typed when a sheet tab is pressed (the formula goes on there). */
    keepsFocus: () => pointing.ready || !!carried,
    setReadOnly(flag) { bar.setReadOnly(flag); grid.setReadOnly(flag); grid.setFillHandle(!flag && ctx.allowed?.('autoFill') !== false); if (flag) editor.close(); },
    reset() { editor.close(); cmd.reset(); },
    destroy() {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
      document.removeEventListener('copy', onCopy);
      document.removeEventListener('cut', onCopy);
      document.removeEventListener('paste', onPaste);
      clearTimeout(pasteTimer);
    },
  };
}
