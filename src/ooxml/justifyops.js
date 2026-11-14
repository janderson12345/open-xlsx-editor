// Home > Editing > Fill > Justify: the text of a selection's first column rewrapped to the selection's width, as
// Excel's Range.Justify does it. Measured over COM in a hidden Excel (docs/FILL.md):
// - Each run of filled cells is a paragraph: their text joined with a space, spaces at its start dropped, then
//   broken at spaces into lines as wide as fit. A word wider than that has a line of its own. Spaces beyond the one
//   a break takes stay at the end of the line ("lead and   inner" in a narrow column: "lead and  ", "inner").
// - The blank rows after each paragraph stay as many, so the paragraphs below move down or up with it; rows the text
//   no longer needs, to the selection's last, are cleared. Text needing more rows than selected goes on below,
//   over whatever is there (Excel asks first: "Text will extend below range.").
// - Only the first column's values move: the other columns of the selection and every cell's format stay.
// - A number, TRUE or FALSE, an error or a formula in the first column: Excel refuses, and nothing changes.

export const JUSTIFY_REFUSED = 'Cannot justify cells containing numbers or formulas.';
export const JUSTIFY_EXTENDS = 'Text will extend below range.';

/** The rows Justify writes for `texts` (one per row of the first column: its text, or null when empty), each line
 *  kept while `fits(line)`: the new rows' text, null for a blank row, as far as the last line. */
export function justifyLines(texts, fits) {
  const out = [];
  let para = [];
  const flush = () => {
    if (!para.length) return;
    const words = para.join(' ').replace(/^ +/, '').split(' ');
    para = [];
    let line = words[0];
    for (const word of words.slice(1)) {
      const next = `${line} ${word}`;
      if (fits(next)) line = next;
      else { out.push(line); line = word; }
    }
    if (line !== '') out.push(line);
  };
  for (const t of texts) {
    if (t == null || t === '') { flush(); out.push(null); } else para.push(t);
  }
  flush();
  while (out.length && out[out.length - 1] == null) out.pop();
  return out;
}

/** Justify over `sel` ({ r0, c0, r1, c1 }) of sheet si: { error } when Excel refuses, else { steps, extends } (the
 *  steps for runSteps; `extends` when the text goes on below the selection, Excel's question first). `fits(line)`
 *  says whether a line fits the selection's width, as Excel measures it. */
export function justifySteps(wb, si, sel, fits) {
  const sheet = wb.sheet(si);
  const c = sel.c0;
  const texts = [];
  for (let r = sel.r0; r <= sel.r1; r++) {
    const cell = sheet.cells.get(r)?.get(c);
    if (!cell || cell.v == null || cell.v === '') { texts.push(null); continue; }
    if (cell.f != null || typeof cell.v !== 'string' || cell.t === 'e' || cell.t === 'b') return { error: JUSTIFY_REFUSED };
    texts.push(cell.v);
  }
  const lines = justifyLines(texts, fits);
  const last = Math.max(sel.r1, sel.r0 + lines.length - 1);
  const contents = [];
  for (let r = sel.r0; r <= last; r++) {
    const value = lines[r - sel.r0] ?? null;
    const had = sheet.cells.get(r)?.get(c);
    const now = had && had.v != null && had.v !== '' ? had : null;
    if (value == null ? now : !now || now.f != null || now.v !== value) contents.push({ r, c, value });
  }
  return { extends: lines.length > sel.r1 - sel.r0 + 1, steps: [{ si, contents, styles: [], area: { r0: sel.r0, c0: sel.c0, r1: last, c1: sel.c1 } }] };
}
