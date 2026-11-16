// Notes and threaded comments on the sheet (ooxml/notes.js), as Excel shows them: a small red triangle in the top-right
// corner of a cell with a note, a purple one for a comment thread; the note (its author in bold) or the thread (each
// author, date and text) in a box beside the cell while the pointer rests on it, or after Shift+F2. Assistive
// technology hears the note as the cell's description.
import { div, setVars } from './dom.js';

/** A comment's text into `el`, each @mention in it (ooxml/notes.js: start, length) highlighted as Excel shows one. */
export function mentionText(el, text, mentions = []) {
  let at = 0;
  for (const m of [...mentions].sort((a, b) => a.start - b.start)) {
    if (m.start < at || m.start + m.length > text.length) continue;
    el.append(text.slice(at, m.start));
    const s = document.createElement('span');
    s.className = 'oxe-mention';
    s.textContent = text.slice(m.start, m.start + m.length);
    el.append(s);
    at = m.start + m.length;
  }
  el.append(text.slice(at));
  return el;
}

/** What a screen reader hears for a note or a thread. */
export function describeNote(n) {
  if (n.kind === 'thread') return `Comment thread${n.resolved ? ', resolved' : ''}. ${n.items.map((x) => `${x.author}: ${x.text}`).join('. ')}`;
  return `Note. ${n.text.replace(/\s*\n\s*/g, ' ')}`;
}

/** The box shown beside cell (r, c): a note where its VML shape puts it, at its size; else (a thread, or a note with
 *  no shape) beside the cell. */
function noteBox(st, n, r, c) {
  const box = div(n.kind === 'thread' ? 'oxe-note-box oxe-thread-box' : 'oxe-note-box');
  const a = n.anchor;
  if (a) {
    const x0 = st.cols.start(a.c0) + a.dx0;
    const y0 = st.rows.start(a.r0) + a.dy0;
    setVars(box, { '--oxe-x': x0, '--oxe-y': y0, '--oxe-nw': Math.max(20, st.cols.start(a.c1) + a.dx1 - x0), '--oxe-nh': Math.max(12, st.rows.start(a.r1) + a.dy1 - y0) });
    box.classList.add('oxe-note-sized');
  } else setVars(box, { '--oxe-x': st.cols.start(c + 1) + 10, '--oxe-y': st.rows.start(r) });
  if (n.kind === 'thread') {
    for (const item of n.items) {
      const one = div('oxe-thread-item');
      const who = document.createElement('strong');
      who.textContent = item.author;
      const when = document.createElement('span');
      when.className = 'oxe-thread-date';
      when.textContent = item.date ? new Date(item.date).toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' }) : '';
      const text = div('oxe-thread-text');
      mentionText(text, item.text, item.mentions);
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
    const m = div(n.kind === 'thread' ? `oxe-note-mark oxe-thread-mark${n.resolved ? ' oxe-thread-done' : ''}` : 'oxe-note-mark');
    setVars(m, { '--oxe-x': st.cols.start(c + 1), '--oxe-y': st.rows.start(r) });
    marks.push(m);
    if (st.noteTexts && st.idBase) {
      const d = document.createElement('div');
      d.id = `${st.idBase}-note-${r}-${c}`;
      d.textContent = describeNote(n);
      texts.push(d);
    }
    // (Open: hovered or Shift+F2'd; a note its shape shows always (x:Visible, Show/Hide Note in Excel); every note
    // with Review > Show All Notes.)
    if (st.noteOpen === key || (n.kind === 'note' && (n.visible || st.allNotes))) marks.push(noteBox(st, n, r, c));
  }
  layer.replaceChildren(...marks);
  if (st.noteTexts && st.idBase) st.noteTexts.replaceChildren(...texts);
}
