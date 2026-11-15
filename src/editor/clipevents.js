// Copy, cut and paste through the browser's clipboard events (Ctrl+C, X, V, and the browser's own menu): the cells go
// on the system clipboard as text, tabs between cells and a line for each row, as Excel puts them.
// Attached by interact.js to the grid's view; returns the keys' handlers and a detach.
import { announce } from './live.js';
import { refName } from '../ooxml/refs.js';

/** The clipboard for the grid (`view`, `grid`), its copy, cut and paste run through `run(name, ...)` (the commands). */
export function attachClipboard({ view, grid, ctx, run }) {
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
    const clip = run(cut ? 'cut' : 'copy');
    if (!clip) return null;
    const text = clipText(clip);
    clip.text = norm(text);
    announce(`${cut ? 'Cut' : 'Copied'} ${refName(clip.rect.r0, clip.rect.c0)}${clip.rect.r1 > clip.rect.r0 || clip.rect.c1 > clip.rect.c0 ? ` to ${refName(clip.rect.r1, clip.rect.c1)}` : ''}`);
    return text;
  };
  const pasteText = (text) => run('paste', text == null ? null : norm(text));
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
  return {
    keyCopy,
    keyPaste,
    takeClip,
    pasteText,
    detach() {
      clearTimeout(pasteTimer);
      document.removeEventListener('copy', onCopy);
      document.removeEventListener('cut', onCopy);
      document.removeEventListener('paste', onPaste);
    },
  };
}
