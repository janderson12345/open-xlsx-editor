// Moving the active cell as Excel does: by a cell, to the edge of a block of data (Ctrl+arrow), to the row's start
// (Home), the sheet's start (Ctrl+Home) or its last used cell (Ctrl+End); over merged ranges as single cells.
// Pure functions of the sheet model, so that they can be tested without a page.

const filled = (sheet, r, c) => sheet.cells.get(r)?.get(c)?.v != null;
const mergeAt = (sheet, r, c) => sheet.merges.find((m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1) || null;
const clamp = (v, max) => Math.max(0, Math.min(v, max));

/** One step from (r, c) in direction (dr, dc), out of a merged range from its far side, onto a merged range's
 *  top-left cell. `limits` = { rows, cols } (the last index of each). */
export function step(sheet, r, c, dr, dc, limits) {
  const from = mergeAt(sheet, r, c);
  const br = from ? (dr > 0 ? from.r1 : dr < 0 ? from.r0 : r) : r;
  const bc = from ? (dc > 0 ? from.c1 : dc < 0 ? from.c0 : c) : c;
  const nr = clamp(br + dr, limits.rows);
  const nc = clamp(bc + dc, limits.cols);
  const to = mergeAt(sheet, nr, nc);
  return to ? { r: to.r0, c: to.c0 } : { r: nr, c: nc };
}

/** Ctrl+arrow: from a filled cell whose neighbour is filled, to the last filled cell of the block; otherwise to the
 *  next filled cell in that direction, or the edge. */
export function jump(sheet, r, c, dr, dc, limits) {
  const inside = (nr, nc) => nr >= 0 && nc >= 0 && nr <= limits.rows && nc <= limits.cols;
  let nr = r + dr;
  let nc = c + dc;
  if (!inside(nr, nc)) return { r, c };
  if (filled(sheet, r, c) && filled(sheet, nr, nc)) {
    while (inside(nr + dr, nc + dc) && filled(sheet, nr + dr, nc + dc)) { nr += dr; nc += dc; }
    return { r: nr, c: nc };
  }
  while (inside(nr, nc) && !filled(sheet, nr, nc)) {
    if (!inside(nr + dr, nc + dc)) return { r: nr, c: nc };
    nr += dr;
    nc += dc;
  }
  return { r: nr, c: nc };
}

/** The last used cell (Ctrl+End): the last row and the last column with anything in them. */
export const lastCell = (sheet) => ({ r: Math.max(0, sheet.maxRow), c: Math.max(0, sheet.maxCol) });
