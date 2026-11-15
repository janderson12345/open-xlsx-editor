// Draw Borders (Home > Font > Borders): Draw Border, Draw Border Grid and Erase Border as modes, and the Line Color and
// Line Style they draw with. Measured over COM (tools/excel-compare/excel-drawborders.ps1, GetPressedMso after each
// step): one mode at a time, a second press of a mode's button ends it, picking another switches to it, and a cell
// selected ends it (typing a value over COM didn't). While a mode is on, the grid's drag draws instead of selecting
// (interact.js) and Esc ends it; each drag is the editor's drawBorders command, one step of the history.
// The mode, the line and the colour are view state: they never reach the workbook or the undo history.
import { announce } from './live.js';
import { colorBox } from './colorbox.js';
import { recentColors } from '../ooxml/recentcolors.js';
import { LINES } from './formatbox-kit.js';

const NAMES = { draw: 'Draw Border', grid: 'Draw Border Grid', erase: 'Erase Border' };
/** The ribbon's actions this module answers (acts.js hands them over). */
export const DRAW_ACTS = new Set(['drawBorderMode', 'borderLineColor', 'borderLineStyle', 'borderLine', 'drawing', 'drawn', 'stopDrawing', 'moreColorsBox', 'recentColors']);

/** `env` is mountXlsxEditor's shared state and helpers (index.js, as acts.js gets them). */
export function borderDrawing(env) {
  const { root, state, editing } = env;
  // (Thin and automatic to start with, as Excel's Line Style and Line Color.)
  const now = { mode: null, line: { style: 'thin', color: null } };
  const show = () => {
    root.classList.toggle('oxe-drawing-borders', now.mode === 'draw' || now.mode === 'grid');
    root.classList.toggle('oxe-erasing-borders', now.mode === 'erase');
    state.ribbon?.drawing?.(now.mode);
  };
  const setMode = (mode) => {
    if (mode && state.readOnly) return false;
    const was = now.mode;
    now.mode = mode && mode !== was ? mode : null;
    show();
    if (now.mode) announce(`${NAMES[now.mode]} on: drag over the cells, Esc to stop`);
    else if (was) announce(`${NAMES[was]} off`);
    return true;
  };
  // (A cell selected ends the mode, as Excel's Range.Select did; the drawing itself selects nothing.)
  state.onSelect.add(() => { if (now.mode) setMode(null); });
  // (Picking a line colour or style with no mode on turns Draw Border on, as Excel's menus do; COM can't pick them,
  // so this is Excel's documented behaviour, not a measurement.)
  const lineChanged = () => { if (!now.mode) setMode('draw'); return true; };
  const lineOut = () => (now.line.color ? { style: now.line.style, color: now.line.color } : { style: now.line.style });
  /** More Colors: the box on `target`'s colour now ('font', 'fill' or 'line'); the colour picked (Recent Colors too). */
  const moreColors = async (target, current = null) => {
    const rgb = await colorBox(root, { current: current?.rgb || (target === 'line' ? now.line.color?.rgb : null) });
    if (!rgb) return false;
    if (target === 'line') {
      now.line.color = { rgb };
      editing().run('recentColor', rgb);
      lineChanged();
      return now.line.color;
    }
    return editing().run('moreColor', target, rgb) ? { rgb } : false;
  };
  return {
    act(name, ...args) {
      if (name === 'drawing') return now.mode;
      if (name === 'borderLine') return { ...now.line };
      if (name === 'recentColors') return state.workbook ? recentColors(state.workbook) : [];
      if (name === 'drawBorderMode') return setMode(args[0] || null);
      if (name === 'stopDrawing') return now.mode ? setMode(null) : false;
      if (name === 'borderLineColor') { now.line.color = args[0] || null; return lineChanged(); }
      if (name === 'borderLineStyle') {
        if (!LINES.some(([v]) => v === args[0])) return false;
        now.line.style = args[0];
        return lineChanged();
      }
      if (name === 'moreColorsBox') return moreColors(args[0], args[1]);
      // (A drag let go: its range, or a gridline clicked, `side` of a one-cell range.)
      if (name === 'drawn') return now.mode ? editing().run('drawBorders', now.mode, args[0], { side: args[1] || null, line: lineOut() }) : false;
      return false;
    },
  };
}
