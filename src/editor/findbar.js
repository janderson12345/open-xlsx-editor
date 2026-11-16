// Find and Replace: the bar under the ribbon (Ctrl+F, Ctrl+H), as open-docx-editor's findbar.js, with Excel's options
// (Match case, Match entire cell contents, Within, Search, Look in) and Find All. The search itself is findops.js.
import { el } from './ribbonkit.js';
import { findAll, findNext, replaceChanges, searchText } from '../ooxml/findops.js';
import { colName } from '../ooxml/refs.js';
import { ruleFormatBox } from './cfformatbox.js';
import { describeFormat, formatFromCell } from '../ooxml/findformat.js';

const NOT_FOUND = "We couldn't find what you were looking for. Select Options for more ways to search.";
const NOTHING_TO_REPLACE = "We couldn't find anything to replace. Select Options for more ways to search.";

/** A cell picked on the sheet (Choose Format From Cell): the active cell once a click on the grid lets go, or when
 *  Enter is pressed there; null on Esc. `here()` gives the active cell, { si, r, c }. */
export function pickCell(grid, here) {
  return new Promise((resolve) => {
    const finish = (hit) => {
      grid.removeEventListener('mouseup', up, true);
      grid.removeEventListener('keydown', key, true);
      grid.classList.remove('oxe-picking');
      resolve(hit);
    };
    // (After the click has made the cell the active one.)
    const up = () => setTimeout(() => finish(here()), 0);
    const key = (e) => {
      if (e.key !== 'Enter' && e.key !== 'Escape') return;
      e.preventDefault();
      e.stopPropagation();
      finish(e.key === 'Enter' ? here() : null);
    };
    grid.addEventListener('mouseup', up, true);
    grid.addEventListener('keydown', key, true);
    grid.classList.add('oxe-picking');
    grid.focus({ preventScroll: true });
  });
}

/**
 * The bar, hidden until opened. `ctx` = { workbook(), from() (the active cell, { si, r, c }), within() (the
 * selection's areas to look in, { si, areas }, or null with one cell selected), go(hit) (that cell selected, its
 * sheet shown; inside the selection, the active cell moved there), apply(steps) (replacements as one step of the undo history), home() (focus back to the
 * sheet), readOnly(), pickCell() (a cell picked on the sheet) }. Returns { el, open(replace), close() }.
 */
export function createFindBar(ctx) {
  let seq = 0;
  const id = `oxe-find${Math.random().toString(36).slice(2, 7)}`;
  const field = (cls, label, attrs = {}) => el('input', { type: 'text', class: cls, 'aria-label': label, placeholder: label, ...attrs });
  const check = (label) => {
    const box = el('input', { type: 'checkbox', id: `${id}-${++seq}` });
    return [box, el('label', { class: 'oxe-find-opt', for: box.id }, box, label)];
  };
  const choice = (label, options) => {
    const sel = el('select', { 'aria-label': label, title: label }, ...options.map(([v, t]) => el('option', { value: v }, t)));
    return [sel, el('label', { class: 'oxe-find-opt' }, `${label} `, sel)];
  };
  const findBox = field('oxe-findbox', 'Find what');
  const replBox = field('oxe-findbox', 'Replace with');
  const [matchCase, caseOpt] = check('Match case');
  const [whole, wholeOpt] = check('Match entire cell contents');
  const [within, withinOpt] = choice('Within:', [['sheet', 'Sheet'], ['workbook', 'Workbook']]);
  const [order, orderOpt] = choice('Search:', [['rows', 'By Rows'], ['cols', 'By Columns']]);
  const [lookIn, lookOpt] = choice('Look in:', [['formulas', 'Formulas'], ['values', 'Values'], ['notes', 'Notes'], ['comments', 'Comments']]);
  const info = el('span', { class: 'oxe-findinfo', role: 'status' }); // (read out by screen readers, 4.1.3)
  // (Find Format and Replace Format: Excel's Format... beside each box, its format shown, Clear to take it off.)
  const formats = { find: null, replace: null };
  const formatControl = (key, label) => {
    const shown = el('span', { class: 'oxe-find-format', 'aria-live': 'polite' }, 'No Format Set');
    const set = (f) => {
      formats[key] = f && Object.keys(f).length ? f : null;
      const f2 = formats[key];
      shown.textContent = describeFormat(f2, f2?.from ? `${colName(f2.from.c)}${f2.from.r + 1}` : '');
      // (A preview of the format, as Excel's box shows one: its fill, ink and weight.)
      const fill = f2?.from ? (/^[0-9A-F]{6}$/.test(f2.fill) ? f2.fill : null) : f2?.fill && f2.fill !== 'none' ? f2.fill.slice(-6) : null;
      shown.style.setProperty('--oxe-find-fill', fill ? `#${fill}` : 'transparent');
      shown.style.setProperty('--oxe-find-ink', f2?.color ? `#${f2.color.slice(-6)}` : 'inherit');
      shown.classList.toggle('oxe-find-format-set', !!f2);
      shown.classList.toggle('oxe-find-format-bold', !!f2?.bold);
      shown.classList.toggle('oxe-find-format-italic', !!f2?.italic);
      clear.hidden = !f2;
    };
    const clear = el('button', { type: 'button', title: `Clear ${label}`, hidden: true, onmousedown: (e) => { e.preventDefault(); set(null); } }, 'Clear');
    const pick = el('button', { type: 'button', title: label, onmousedown: (e) => { e.preventDefault(); ctx.hold?.(true); ruleFormatBox(ctx.root).then((f) => { if (f) set(f); ctx.hold?.(false); findBox.focus(); }); } }, 'Format...');
    // (Choose Format From Cell: the next cell clicked, or the one Enter is pressed on, gives its whole format.)
    const fromCell = el('button', { type: 'button', title: `Choose Format From Cell (${label})`, onmousedown: (e) => {
      e.preventDefault();
      if (!ctx.pickCell) return;
      say('Click the cell whose format to use, or select it and press Enter (Esc to cancel).');
      ctx.pickCell().then((hit) => {
        if (hit) set(formatFromCell(ctx.workbook(), hit.si, hit.r, hit.c));
        say('');
        (key === 'replace' ? replBox : findBox).focus();
      });
    } }, 'From Cell...');
    return el('span', { class: 'oxe-find-formats' }, shown, pick, fromCell, clear);
  };
  const opts = (forReplace = false) => ({
    matchCase: matchCase.checked, whole: whole.checked, within: within.value, byColumns: order.value === 'cols',
    lookIn: forReplace ? 'formulas' : lookIn.value,
    // (More than one cell selected, Within: Sheet: Find and Replace look there only, every range of it, as Excel's.)
    ...(() => { const got = within.value === 'sheet' && ctx.within?.(); return got ? { areas: got.areas, areasSheet: got.si } : {}; })(),
    ...(formats.find ? { findFormat: formats.find } : {}), ...(forReplace && formats.replace ? { replaceFormat: formats.replace } : {}),
  });
  const say = (text) => { info.textContent = text; };
  const button = (label, title, fn) => el('button', { type: 'button', title, onmousedown: (e) => { e.preventDefault(); fn(); } }, label);

  /** Find Next (dir 1) or Find Previous (-1): the match selected, the bar keeping focus. */
  function next(dir = 1) {
    const hit = findNext(ctx.workbook(), ctx.from(), findBox.value, opts(), dir);
    if (!hit) { say(findBox.value ? NOT_FOUND : ''); return false; }
    ctx.go(hit);
    say('');
    return true;
  }
  /** Find All: how many, and Excel's list of them (sheet, cell, value, formula), each going to its cell. */
  function all() {
    const wb = ctx.workbook();
    const hits = findAll(wb, ctx.from().si, findBox.value, opts());
    say(hits.length ? `${hits.length} cell(s) found` : NOT_FOUND);
    list.replaceChildren(...hits.slice(0, 500).map((hit) => {
      const cell = wb.sheet(hit.si).cells.get(hit.r)?.get(hit.c);
      const value = searchText(wb, hit.si, cell, 'values');
      const formula = cell?.f != null ? searchText(wb, hit.si, cell, 'formulas') : '';
      const row = el('button', { type: 'button', class: 'oxe-findall-row', onmousedown: (e) => { e.preventDefault(); ctx.go(hit); } },
        el('span', {}, wb.sheets[hit.si].name), el('span', {}, `$${colName(hit.c)}$${hit.r + 1}`), el('span', {}, value), el('span', {}, formula));
      row.setAttribute('aria-label', `${wb.sheets[hit.si].name} ${colName(hit.c)}${hit.r + 1}${value ? `, ${value}` : ''}${formula ? `, ${formula}` : ''}`);
      return row;
    }));
    list.hidden = !hits.length;
  }
  /** Replace: the active cell, if it matches, then on to the next match. */
  function replaceOne() {
    if (ctx.readOnly()) return;
    const at = ctx.from();
    const steps = replaceChanges(ctx.workbook(), [at], findBox.value, replBox.value, opts(true));
    if (steps.count) ctx.apply(steps);
    const hit = findNext(ctx.workbook(), at, findBox.value, opts(true), 1);
    if (hit && !(hit.si === at.si && hit.r === at.r && hit.c === at.c)) ctx.go(hit);
    say(steps.count || hit ? '' : NOTHING_TO_REPLACE);
  }
  function replaceAll() {
    if (ctx.readOnly()) return;
    const wb = ctx.workbook();
    const steps = replaceChanges(wb, findAll(wb, ctx.from().si, findBox.value, opts(true)), findBox.value, replBox.value, opts(true));
    if (!steps.count) { say(NOTHING_TO_REPLACE); return; }
    ctx.apply(steps);
    say(`All done. We made ${steps.count} replacement${steps.count === 1 ? '' : 's'}.${steps.refused ? ` ${steps.refused} left as they were.` : ''}`);
  }

  // (Find All's list, Excel's columns: Sheet, Cell, Value, Formula.)
  const list = el('div', { class: 'oxe-findall', role: 'group', 'aria-label': 'Cells found', hidden: true });
  const replRow = el('div', { class: 'oxe-replrow' }, replBox, formatControl('replace', 'Replace Format'), button('Replace', 'Replace', replaceOne), button('Replace All', 'Replace All', replaceAll));
  const bar = el('div', { class: 'oxe-findbar', role: 'search', 'aria-label': 'Find and Replace', hidden: true },
    el('div', { class: 'oxe-findrow' }, findBox, formatControl('find', 'Find Format'),
      button('↑', 'Find Previous (Shift+Enter)', () => next(-1)), button('↓', 'Find Next (Enter)', () => next(1)), button('Find All', 'Find All', all),
      button('✕', 'Close (Esc)', () => close())),
    replRow,
    el('div', { class: 'oxe-findrow oxe-findopts' }, caseOpt, wholeOpt, withinOpt, orderOpt, lookOpt),
    info, list);
  findBox.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); next(e.shiftKey ? -1 : 1); }
  });
  replBox.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); replaceOne(); }
  });
  bar.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') { e.preventDefault(); close(); }
    const ctrl = e.ctrlKey || e.metaKey;
    if (ctrl && (e.key === 'f' || e.key === 'h')) { e.preventDefault(); open(e.key === 'h'); }
  });

  function open(replace = false) {
    const withReplace = replace && !ctx.readOnly();
    bar.hidden = false;
    replRow.hidden = !withReplace;
    // (Replace looks in formulas only, as Excel's Replace does.)
    lookIn.disabled = withReplace;
    if (withReplace) lookIn.value = 'formulas';
    say('');
    list.hidden = true;
    // (At once, not after a moment: typing straight after Ctrl+H would land in whichever box had focus by then.)
    findBox.focus();
    findBox.select();
  }
  function close() {
    bar.hidden = true;
    ctx.home();
  }
  return { el: bar, open, close, get isOpen() { return !bar.hidden; } };
}
