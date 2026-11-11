// Data > Flash Fill (Ctrl+E): the column's examples (its cells already typed) read as a pattern of the cells to their
// left, and the pattern carried down the rest (measured over COM: excel-datatools.ps1, five patterns). A pattern is
// pieces in a row: a whole cell, a word of one (split at spaces and punctuation), a word's first letter, characters at
// fixed places in a cell, or fixed text; each in its own case, upper or lower. The first pattern (fewest pieces, words
// before places) that makes every example is used; none, and Excel's words. Numbers made from an example that was a
// number stay numbers, as Excel's do.

const words = (s) => String(s).match(/[A-Za-z0-9]+/g) || [];
const CASES = [['same', (t) => t], ['upper', (t) => t.toUpperCase()], ['lower', (t) => t.toLowerCase()], ['proper', (t) => t.charAt(0).toUpperCase() + t.slice(1).toLowerCase()]];

/** Every piece that could make the start of `rest` from `sources` (the row's cells, as text). */
function piecesAt(rest, sources) {
  const out = [];
  sources.forEach((src, j) => {
    for (const [name, fn] of CASES) {
      if (src && rest.startsWith(fn(src))) out.push({ kind: 'cell', j, case: name, len: src.length });
      words(src).forEach((w, i) => {
        if (rest.startsWith(fn(w))) out.push({ kind: 'word', j, i, case: name, len: w.length });
        if (rest.startsWith(fn(w.charAt(0)))) out.push({ kind: 'initial', j, i, case: name, len: 1 });
      });
    }
    // (Characters at fixed places: the longest run of this cell the output starts with.)
    for (let a = 0; a < src.length; a++) {
      let n = 0;
      while (a + n < src.length && n < rest.length && src[a + n] === rest[n]) n++;
      if (n) out.push({ kind: 'chars', j, a, n, case: 'same', len: n });
    }
  });
  return out;
}

/** What a piece makes from a row's cells (null when the row hasn't what it takes). */
function run(piece, sources) {
  if (piece.kind === 'text') return piece.text;
  const src = sources[piece.j] ?? '';
  const fn = CASES.find(([n]) => n === piece.case)[1];
  if (piece.kind === 'cell') return src ? fn(src) : null;
  if (piece.kind === 'chars') return src.length >= piece.a + piece.n ? src.slice(piece.a, piece.a + piece.n) : null;
  const w = words(src)[piece.i];
  if (w == null) return null;
  return piece.kind === 'word' ? fn(w) : fn(w.charAt(0));
}
const make = (program, sources) => { let s = ''; for (const p of program) { const t = run(p, sources); if (t == null) return null; s += t; } return s; };
// (Characters at fixed places come last but for fixed text: words fit names of any length, places only alike ones.)
const rank = (p) => ({ cell: 0, initial: 0, word: 1, chars: 20, text: 6 }[p.kind]);

/** Patterns that make `target` from `sources`, fewest pieces first (a bounded search). */
function programs(target, sources, limit = 400) {
  const found = [];
  const walk = (at, sofar) => {
    if (found.length >= limit) return;
    if (at === target.length) { found.push(sofar); return; }
    if (sofar.length > 8) return;
    const rest = target.slice(at);
    const options = piecesAt(rest, sources).sort((a, b) => b.len - a.len || rank(a) - rank(b));
    for (const p of options) walk(at + p.len, [...sofar, p]);
    // (Fixed text: one character at a time, run together after.)
    const last = sofar.at(-1);
    walk(at + 1, last?.kind === 'text' ? [...sofar.slice(0, -1), { kind: 'text', text: last.text + rest[0] }] : [...sofar, { kind: 'text', text: rest[0] }]);
  };
  walk(0, []);
  // (Fixed text costs by its length: a pattern that explains the output from the cells wins over copying the example.)
  const cost = (prog) => prog.length * 10 + prog.reduce((s, p) => s + rank(p) + (p.kind === 'text' ? 12 * p.text.length : 0), 0);
  return found.sort((a, b) => cost(a) - cost(b));
}

/**
 * Flash Fill of column `col` (absolute), rows r0 to r1 of sheet si, from the cells to its left in `sourceCols`. Rows
 * with a value already are the examples; the others are filled. Returns { steps } or { error } with Excel's words.
 */
export function flashFill(wb, si, { col, r0, r1, sourceCols }) {
  const sheet = wb.sheet(si);
  const text = (r, c) => { const x = sheet.cells.get(r)?.get(c); return x?.v == null ? '' : String(x.v); };
  const rows = [];
  for (let r = r0; r <= r1; r++) if (sourceCols.some((c) => text(r, c) !== '')) rows.push(r);
  const examples = rows.filter((r) => text(r, col) !== '');
  if (!examples.length) return { error: "Flash Fill isn't available: type an example in the column first." };
  const sources = (r) => sourceCols.map((c) => text(r, c));
  const numeric = examples.every((r) => typeof sheet.cells.get(r)?.get(col)?.v === 'number');
  // (A number example matched as Excel shows it: 0042 typed is 42.)
  const want = (r) => text(r, col);
  const program = programs(want(examples[0]), sources(examples[0])).find((p) => examples.every((r) => make(p, sources(r)) === want(r)));
  if (!program) return { error: "Flash Fill didn't recognize a pattern to fill." };
  const out = { si, contents: [], styles: [], area: { r0, c0: col, r1, c1: col } };
  for (const r of rows) {
    if (text(r, col) !== '') continue;
    const v = make(program, sources(r));
    if (v == null) continue;
    out.contents.push({ r, c: col, value: numeric && v.trim() !== '' && Number.isFinite(Number(v)) ? Number(v) : v });
  }
  return { steps: [out], filled: out.contents.length };
}
