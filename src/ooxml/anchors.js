// What sits on a sheet at a place of its own, moved when its rows and columns are inserted or deleted, as Excel moves
// it (measured over COM: tools/excel-compare/excel-anchors.ps1, docs/ROWCOL.md):
// - notes and threaded comments go with their cell (also when cells shift), and are deleted with it;
// - pictures, charts and shapes (a drawing's anchors) move and stretch with whole rows and columns: a corner after the
//   change moves with it, an end corner on the boundary where lines are inserted stays, and one whose whole extent is
//   deleted is deleted; shifting cells doesn't move them.
// Each part is edited where it says so; the rest of it stays byte for byte.

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const THREADED = 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment';

/**
 * A drawing's corner (`line`, `off` within it) on the axis that changes, after inserting `n` lines before `at` (n > 0)
 * or deleting -n from `at` (n < 0). `end`: the corner where the object ends (on a boundary it belongs to the line
 * before).
 */
export function shiftPoint(line, off, end, at, n) {
  if (n > 0) return line > at || (line === at && (!end || off > 0)) ? { line: line + n, off } : { line, off };
  const last = at - n - 1;
  if (line < at || (end && line === at && off === 0)) return { line, off };
  if (line > last) return { line: line + n, off };
  return { line: at, off: 0 };
}

/** Whether an object from (line, off) to (line, off) lies wholly in the lines deleted (n < 0) from `at`. */
export function wholeDeleted(from, to, at, n) {
  if (n >= 0) return false;
  const after = at - n; // (the first line after those deleted)
  return from.line >= at && (to.line < after || (to.line === after && to.off === 0));
}

const relsOf = (wb, path, type) => wb.pkg.readRels(path).filter((r) => r.type === type || r.type.endsWith(`/${type}`));

/**
 * Notes and threaded comments of sheet si to their cells' new places: `move(r, c)` gives a cell's place after the
 * change, or null when it is deleted.
 */
export function moveNotes(wb, si, move) {
  const path = wb.sheets[si].path;
  const pkg = wb.pkg;
  const cellAt = (ref) => {
    const m = /^([A-Z]+)(\d+)$/.exec(ref);
    if (!m) return null;
    let c = 0;
    for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64;
    return { r: Number(m[2]) - 1, c: c - 1 };
  };
  const name = (r, c) => { let s = ''; for (let k = c + 1; k > 0; k = Math.floor((k - 1) / 26)) s = String.fromCharCode(65 + ((k - 1) % 26)) + s; return `${s}${r + 1}`; };
  const refTo = (ref) => { const at = cellAt(ref); const to = at && move(at.r, at.c); return at ? (to ? name(to.r, to.c) : null) : ref; };
  // The comments part: each comment's ref.
  for (const rel of relsOf(wb, path, `${REL}/comments`)) {
    const text = pkg.text(rel.target);
    if (!text) continue;
    pkg.setText(rel.target, text.replace(/<comment\b[^>]*?\sref="([^"]+)"[^>]*>[\s\S]*?<\/comment>/g, (el, ref) => {
      const next = refTo(ref);
      return next == null ? '' : el.replace(`ref="${ref}"`, `ref="${next}"`);
    }));
  }
  // Threaded comments: each one and its replies, by ref.
  for (const rel of relsOf(wb, path, THREADED)) {
    const text = pkg.text(rel.target);
    if (!text) continue;
    pkg.setText(rel.target, text.replace(/<threadedComment\b[^>]*?\sref="([^"]+)"[^>]*?(?:\/>|>[\s\S]*?<\/threadedComment>)/g, (el, ref) => {
      const next = refTo(ref);
      return next == null ? '' : el.replace(`ref="${ref}"`, `ref="${next}"`);
    }));
  }
  // The notes' boxes in the VML drawing: their row and column, and the box moved as far as its cell.
  for (const rel of relsOf(wb, path, `${REL}/vmlDrawing`)) {
    const text = pkg.text(rel.target);
    if (!text) continue;
    pkg.setText(rel.target, text.replace(/<v:shape\b[^>]*>[\s\S]*?<\/v:shape>/g, (shape) => {
      if (!/ObjectType="Note"/.test(shape)) return shape;
      const row = /<x:Row>(\d+)<\/x:Row>/.exec(shape);
      const col = /<x:Column>(\d+)<\/x:Column>/.exec(shape);
      if (!row || !col) return shape;
      const r = Number(row[1]);
      const c = Number(col[1]);
      const to = move(r, c);
      if (!to) return '';
      const dr = to.r - r;
      const dc = to.c - c;
      return shape
        .replace(row[0], `<x:Row>${to.r}</x:Row>`)
        .replace(col[0], `<x:Column>${to.c}</x:Column>`)
        .replace(/<x:Anchor>([^<]*)<\/x:Anchor>/, (a, list) => {
          const v = list.split(',').map((s) => Number(s.trim()));
          if (v.length !== 8 || v.some(Number.isNaN)) return a;
          v[0] += dc; v[4] += dc; v[2] += dr; v[6] += dr;
          return `<x:Anchor>${v.join(', ')}</x:Anchor>`;
        });
    }));
  }
}

/**
 * Pictures, charts and shapes of sheet si after whole rows (axis 'row') or columns ('col') are inserted (n > 0) or
 * deleted (n < 0) at `at`. An object wholly deleted goes, with its chart part when it is a chart.
 */
export function moveDrawings(wb, si, axis, at, n) {
  const pkg = wb.pkg;
  const tag = axis === 'row' ? ['row', 'rowOff'] : ['col', 'colOff'];
  for (const rel of relsOf(wb, wb.sheets[si].path, `${REL}/drawing`)) {
    const text = pkg.text(rel.target);
    if (!text) continue;
    const gone = [];
    const point = (el, which) => {
      const re = new RegExp(`<(\\w+:)?${which}>([\\s\\S]*?)</\\1?${which}>`);
      const m = re.exec(el);
      if (!m) return null;
      const p = m[1] || '';
      const line = new RegExp(`<${p}${tag[0]}>(\\d+)</${p}${tag[0]}>`).exec(m[2]);
      const off = new RegExp(`<${p}${tag[1]}>(-?\\d+)</${p}${tag[1]}>`).exec(m[2]);
      if (!line || !off) return null;
      return { m, p, line: Number(line[1]), off: Number(off[1]) };
    };
    const put = (el, pt, next) => el.replace(pt.m[0], pt.m[0]
      .replace(`<${pt.p}${tag[0]}>${pt.line}</${pt.p}${tag[0]}>`, `<${pt.p}${tag[0]}>${next.line}</${pt.p}${tag[0]}>`)
      .replace(`<${pt.p}${tag[1]}>${pt.off}</${pt.p}${tag[1]}>`, `<${pt.p}${tag[1]}>${next.off}</${pt.p}${tag[1]}>`));
    const out = text.replace(/<(\w+:)?(twoCellAnchor|oneCellAnchor)\b([^>]*)>[\s\S]*?<\/\1?\2>/g, (el, _p, kind, attrs) => {
      const edit = /editAs="(\w+)"/.exec(attrs)?.[1] || (kind === 'oneCellAnchor' ? 'oneCell' : 'twoCell');
      if (edit === 'absolute') return el;
      const from = point(el, 'from');
      if (!from) return el;
      const to = kind === 'twoCellAnchor' ? point(el, 'to') : null;
      if (to && edit === 'twoCell' && wholeDeleted(from, to, at, n)) {
        for (const m of el.matchAll(/r:(?:id|embed)="([^"]+)"/g)) gone.push(m[1]);
        return '';
      }
      const f = shiftPoint(from.line, from.off, false, at, n);
      let next = put(el, from, f);
      if (to) {
        // (Moved but not sized: the end goes as far as the start went.)
        const t = edit === 'oneCell' ? { line: to.line + (f.line - from.line), off: to.off } : shiftPoint(to.line, to.off, true, at, n);
        next = put(next, to, t);
      }
      return next;
    });
    pkg.setText(rel.target, out);
    // A deleted chart's part (and what only it uses), and its relationship from the drawing.
    for (const id of gone) {
      const r = pkg.readRels(rel.target).find((x) => x.id === id);
      if (!r || r.external) continue;
      if (out.includes(`"${id}"`)) continue; // (still used by another object)
      dropRel(pkg, rel.target, id);
      if (/\/chart$/.test(r.type)) pkg.removeTree(r.target);
    }
  }
}

/** Remove relationship `id` from a part's relationships. */
function dropRel(pkg, partPath, id) {
  const dir = partPath.slice(0, partPath.lastIndexOf('/') + 1);
  const relsPath = `${dir}_rels/${partPath.slice(dir.length)}.rels`;
  const text = pkg.text(relsPath);
  if (text) pkg.setText(relsPath, text.replace(new RegExp(`<Relationship\\b[^>]*\\sId="${id}"[^>]*/>`), ''));
}
