// Notes and threaded comments on the sheet (ooxml/notes.js), as Excel shows them: a small red triangle in the top-right
// corner of a cell with a note, a purple one for a comment thread; the note (its author in bold) or the thread (each
// author, date and text) in a box beside the cell while the pointer rests on it, or after Shift+F2. Assistive
// technology hears the note as the cell's description.
import { div, setVars } from './dom.js';

/** What a screen reader hears for a note or a thread. */
export function describeNote(n) {
  if (n.kind === 'thread') return `Comment thread. ${n.items.map((x) => `${x.author}: ${x.text}`).join('. ')}`;
  return `Note. ${n.text.replace(/\s*\n\s*/g, ' ')}`;
}

/** The box shown beside cell (r, c). */
function noteBox(st, n, r, c) {
  const box = div(n.kind === 'thread' ? 'oxe-note-box oxe-thread-box' : 'oxe-note-box');
  setVars(box, { '--oxe-x': st.cols.start(c + 1) + 10, '--oxe-y': st.rows.start(r) });
  if (n.kind === 'thread') {
    for (const item of n.items) {
      const one = div('oxe-thread-item');
      const who = document.createElement('strong');
      who.textContent = item.author;
      const when = document.createElement('span');
      when.className = 'oxe-thread-date';
      when.textContent = item.date ? new Date(item.date).toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' }) : '';
      const text = div('oxe-thread-text');
      text.textContent = item.text;
      one.append(who, ' ', when, text);
      box.append(one);
    }
  } else {
    for (const run of n.runs) {
      const s = document.createElement(run.bold ? 'strong' : 'span');
      s.textContent = run.text;
      box.append(s);
    }
  }
  return box;
}

/**
 * Draw the marks of `st.notes` (Map "r,c" => note) that fall in window `win` into `st.noteLayer`, the box of the note
 * open (`st.noteOpen`, "r,c") with them, and each drawn cell's description into `st.noteTexts` (ids the cells point
 * to with aria-describedby).
 */
export function drawNoteMarks(st, win) {
  const layer = st.noteLayer;
  if (!layer) return;
  const marks = [];
  const texts = [];
  for (const [key, n] of st.notes || []) {
    const [r, c] = key.split(',').map(Number);
    if (!win || r < win.r0 || r > win.r1 || c < win.c0 || c > win.c1 || st.rows.size(r) < 1 || st.cols.size(c) < 1) continue;
    const m = div(n.kind === 'thread' ? 'oxe-note-mark oxe-thread-mark' : 'oxe-note-mark');
    setVars(m, { '--oxe-x': st.cols.start(c + 1), '--oxe-y': st.rows.start(r) });
    marks.push(m);
    if (st.noteTexts && st.idBase) {
      const d = document.createElement('div');
      d.id = `${st.idBase}-note-${r}-${c}`;
      d.textContent = describeNote(n);
      texts.push(d);
    }
    if (st.noteOpen === key) marks.push(noteBox(st, n, r, c));
  }
  layer.replaceChildren(...marks);
  if (st.noteTexts && st.idBase) st.noteTexts.replaceChildren(...texts);
}
