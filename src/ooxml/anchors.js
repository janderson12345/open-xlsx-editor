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

// --- Tables ----------------------------------------------------------------------------------------------------------
// As Excel does it (tools/excel-compare/excel-anchors.ps1 on anchor-cases.mjs): lines inserted strictly inside a
// table widen it (a new column named Column1, Column2... in its header); at its edges (the line after it, the column
// before it) they move it or leave it; lines deleted inside narrow it, its formulas' references to a deleted column
// become #REF!; all of its data rows deleted leave the header and one empty row; all of it deleted deletes it (its
// references #REF!). Its header row deleted without all of it, and cells shifted through part of it, are refused.

export const TABLE_MESSAGES = {
  shift: "This won't work because it would move cells in a table on your worksheet.",
  header: "A table's header row can't be deleted without the rest of the table.",
  totals: "A table's totals row can't be deleted without the rest of the table yet.",
};

/** The tables of sheet si: [{ path, relId, name, r0, c0, r1, c1, headers, totals, columns }]. */
function tablesOf(wb, si) {
  return relsOf(wb, wb.sheets[si].path, `${REL}/table`).map((rel) => {
    const text = wb.pkg.text(rel.target) || '';
    const ref = /<table\b[^>]*\sref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/.exec(text);
    if (!ref) return null;
    const col = (s) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
    return {
      path: rel.target, relId: rel.id,
      name: /<table\b[^>]*\sdisplayName="([^"]*)"/.exec(text)?.[1] || /<table\b[^>]*\sname="([^"]*)"/.exec(text)?.[1],
      r0: Number(ref[2]) - 1, c0: col(ref[1]), r1: Number(ref[4]) - 1, c1: col(ref[3]),
      headers: /<table\b[^>]*\sheaderRowCount="0"/.test(text) ? 0 : 1,
      totals: /<table\b[^>]*\stotalsRowCount="[1-9]/.test(text) ? 1 : 0,
      columns: [...text.matchAll(/<tableColumn\b[^>]*\sname="([^"]*)"/g)].map((m) => m[1]),
    };
  }).filter(Boolean);
}

/**
 * Rows deleted (n < 0) that hold all of a table's data rows but not its header: Excel keeps the first data row, as
 * it is, and deletes the others (the table left its header and that row). Returns the rows to delete then, [at, n]
 * (n 0: nothing).
 */
export function keepTableRow(wb, si, axis, at, n, band) {
  if (axis !== 'row' || n >= 0) return [at, n];
  for (const t of tablesOf(wb, si)) {
    if (!t.headers || t.totals) continue;
    if (band && (band[0] > t.c0 || band[1] < t.c1)) continue;
    const last = at - n - 1;
    if (at === t.r0 + 1 && last >= t.r1) return [at + 1, n + 1];
  }
  return [at, n];
}

/**
 * What a change does to each table of sheet si: [{ t, shift } (moved), { t, grow: { index, count } }, { t, shift,
 * shrink: [indexes] } or { t, gone }]. A change Excel refuses calls `refuse(message)`, which throws.
 */
export function planTables(wb, si, axis, at, n, band, refuse) {
  const plans = [];
  for (const t of tablesOf(wb, si)) {
    const [l0, l1, p0, p1] = axis === 'row' ? [t.r0, t.r1, t.c0, t.c1] : [t.c0, t.c1, t.r0, t.r1];
    if (band) {
      // (Cells shifted: the table is touched when cells of it move; then only across its whole width, or height.)
      const touched = p1 >= band[0] && p0 <= band[1] && l1 >= at;
      if (!touched) continue;
      if (band[0] > p0 || band[1] < p1) refuse(TABLE_MESSAGES.shift);
    }
    if (n > 0) {
      if (at <= l0) plans.push({ t, shift: n });
      else if (at <= l1) plans.push({ t, grow: { index: at - l0, count: n } });
      continue;
    }
    const a = at;
    const b = at - n - 1;
    if (a > l1) continue;
    if (b < l0) { plans.push({ t, shift: n }); continue; }
    if (a <= l0 && b >= l1) { plans.push({ t, gone: true }); continue; }
    if (axis === 'row' && t.headers && a <= l0) refuse(TABLE_MESSAGES.header);
    if (axis === 'row' && t.totals && b >= l1) refuse(TABLE_MESSAGES.totals);
    const before = Math.max(0, Math.min(b, l0 - 1) - a + 1);
    const inside = [];
    for (let k = Math.max(a, l0); k <= Math.min(b, l1); k++) inside.push(k - l0);
    plans.push({ t, shift: -before, shrink: inside });
  }
  return plans;
}

/**
 * Make the plans (planTables) after the sheet's cells moved: each table part's ref, autoFilter and columns; new columns'
 * names in the header cells (`doc`, the sheet's moved XML, through `setCell(doc, r, c, text)`); deleted tables' parts.
 * Returns the table references to make #REF!: [{ table, columns (names, or null for the whole table) }].
 */
export function applyTables(wb, si, axis, plans, doc, setCell) {
  const pkg = wb.pkg;
  const dead = [];
  const name = (r, c) => { let s = ''; for (let k = c + 1; k > 0; k = Math.floor((k - 1) / 26)) s = String.fromCharCode(65 + ((k - 1) % 26)) + s; return `${s}${r + 1}`; };
  const COLUMN = /<tableColumn\b[^>]*?(?:\/>|>[\s\S]*?<\/tableColumn>)/g;
  for (const plan of plans) {
    const t = plan.t;
    if (plan.gone) {
      dead.push({ table: t.name, columns: null });
      for (const el of [...doc.documentElement.getElementsByTagNameNS('*', 'tablePart')]) {
        if (el.getAttributeNS(REL, 'id') === t.relId || el.getAttribute('r:id') === t.relId) el.parentNode.removeChild(el);
      }
      const wrap = doc.documentElement.getElementsByTagNameNS('*', 'tableParts')[0];
      if (wrap) {
        const left = wrap.getElementsByTagNameNS('*', 'tablePart').length;
        if (left) wrap.setAttribute('count', String(left)); else wrap.parentNode.removeChild(wrap);
      }
      dropRel(pkg, wb.sheets[si].path, t.relId);
      pkg.removeTree(t.path);
      continue;
    }
    let { r0, c0, r1, c1 } = t;
    let text = pkg.text(t.path);
    if (plan.shift) { if (axis === 'row') { r0 += plan.shift; r1 += plan.shift; } else { c0 += plan.shift; c1 += plan.shift; } }
    if (plan.grow) {
      if (axis === 'row') r1 += plan.grow.count;
      else {
        c1 += plan.grow.count;
        // New columns in their place, named as Excel names them: Column1, Column2..., the first names not taken.
        const taken = new Set(t.columns.map((x) => x.toLowerCase()));
        let id = Math.max(0, ...[...text.matchAll(/<tableColumn\b[^>]*\sid="(\d+)"/g)].map((m) => Number(m[1])));
        const added = [];
        for (let k = 1; added.length < plan.grow.count; k++) if (!taken.has(`column${k}`)) added.push(`Column${k}`);
        const cols = [...text.matchAll(COLUMN)].map((m) => m[0]);
        cols.splice(plan.grow.index, 0, ...added.map((nm) => `<tableColumn id="${++id}" name="${nm}"/>`));
        text = text.replace(/<tableColumns\b[^>]*>[\s\S]*?<\/tableColumns>/, `<tableColumns count="${cols.length}">${cols.join('')}</tableColumns>`);
        if (t.headers) added.forEach((nm, k) => setCell(doc, r0, c0 + plan.grow.index + k, nm));
      }
    }
    if (plan.shrink && plan.shrink.length) {
      if (axis === 'row') {
        r1 -= plan.shrink.length;
        // (All of its data rows deleted: the header and one empty row, as Excel leaves it.)
        if (r1 < r0 + t.headers) r1 = r0 + t.headers;
      } else {
        c1 -= plan.shrink.length;
        const kept = [...text.matchAll(COLUMN)].map((m) => m[0]).filter((_, k) => !plan.shrink.includes(k));
        text = text.replace(/<tableColumns\b[^>]*>[\s\S]*?<\/tableColumns>/, `<tableColumns count="${kept.length}">${kept.join('')}</tableColumns>`);
        dead.push({ table: t.name, columns: plan.shrink.map((k) => t.columns[k]) });
      }
    }
    const ref = `${name(r0, c0)}:${name(r1, c1)}`;
    // (A saved sort's levels follow too, as Excel's table sort was seen to (excel-sortstate.ps1, table-then-row-inserted):
    // each over the data rows, its column moved with the table's columns, gone with its column.)
    const levelCol = (c) => {
      const k = c - t.c0;
      if (axis !== 'col') return c;
      if (plan.shift) return c + plan.shift;
      if (plan.grow) return k >= plan.grow.index ? c + plan.grow.count : c;
      if (plan.shrink?.length) return plan.shrink.includes(k) ? null : c - plan.shrink.filter((x) => x < k).length;
      return c;
    };
    text = text.replace(/<sortCondition\b[^>]*?\sref="([A-Z]+)\d+(?::[A-Z]+\d+)?"[^>]*?\/>/g, (el, col) => {
      let c = 0;
      for (const ch of col) c = c * 26 + ch.charCodeAt(0) - 64;
      const to = levelCol(c - 1);
      if (to == null) return '';
      return el.replace(/\sref="[^"]*"/, ` ref="${name(r0 + t.headers, to)}:${name(r1 - t.totals, to)}"`);
    });
    text = text.replace(/(<table\b[^>]*\sref=")[^"]*"/, `$1${ref}"`)
      .replace(/(<autoFilter\b[^>]*\sref=")[^"]*"/, `$1${name(r0, c0)}:${name(r1 - t.totals, c1)}"`)
      .replace(/(<sortState\b[^>]*\sref=")[^"]*"/, `$1${name(r0 + t.headers, c0)}:${name(r1 - t.totals, c1)}"`);
    pkg.setText(t.path, text);
  }
  return dead;
}

/** A formula's table references to deleted columns, or to a deleted table, as #REF! (as Excel writes them). */
export function deadTableRefs(text, dead, tokenize) {
  let toks;
  try { toks = tokenize(text); } catch { return text; }
  let out = '';
  let last = 0;
  for (const tk of toks) {
    if (tk.type !== 'struct' || !tk.table) continue;
    const d = dead.find((x) => x.table && x.table.toLowerCase() === tk.table.toLowerCase());
    if (!d) continue;
    // (The columns it names: [Qty], [[#Data],[Qty]], [@Qty], or Qty alone between the brackets.)
    const named = [...tk.spec.matchAll(/\[([^\]#@][^\]]*)\]|@([^\]]+)/g)].map((m) => (m[1] ?? m[2]).replace(/'(.)/g, '$1').toLowerCase());
    const plain = /^[^[#@]/.test(tk.spec) ? [tk.spec.replace(/'(.)/g, '$1').toLowerCase()] : [];
    if (d.columns != null && ![...named, ...plain].some((nm) => d.columns.some((c) => c.toLowerCase() === nm))) continue;
    out += `${text.slice(last, tk.start)}#REF!`;
    last = tk.end;
  }
  return out + text.slice(last);
}

// --- PivotTables -----------------------------------------------------------------------------------------------------
// As Excel does it (tools/excel-compare/excel-pivot-lines.ps1): lines before a pivot move it, lines after it leave it,
// all of it deleted deletes it (its cache stays, as a workbook's caches may outlive their pivots); lines through it,
// and cells shifted through part of it, are refused in Excel's words. Its cache's source range follows its source
// sheet's lines as a reference does, and stays as it was when all of it is deleted (Excel keeps it).

export const PIVOT_LINE_MESSAGE = "We can't make this change for the selected cells because it will affect a PivotTable. Use the field list to change the report. If you are trying to insert or delete cells, move the PivotTable and try again.";

const areaOf = (ref) => {
  const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(ref || '');
  if (!m) return null;
  const col = (s) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  return { r0: Number(m[2]) - 1, c0: col(m[1]), r1: Number(m[4] ?? m[2]) - 1, c1: col(m[3] ?? m[1]) };
};
const refOf = ({ r0, c0, r1, c1 }) => {
  const name = (r, c) => { let s = ''; for (let k = c + 1; k > 0; k = Math.floor((k - 1) / 26)) s = String.fromCharCode(65 + ((k - 1) % 26)) + s; return `${s}${r + 1}`; };
  return r0 === r1 && c0 === c1 ? name(r0, c0) : `${name(r0, c0)}:${name(r1, c1)}`;
};

/** What a change does to each pivot table on sheet si: [{ p, area, shift } or { p, gone }]; Excel's refusals through
 *  `refuse(message)`, which throws. `pivots`: readPivots(wb). */
export function planPivots(wb, si, axis, at, n, band, refuse, pivots) {
  const plans = [];
  for (const p of pivots) {
    if (p.si !== si) continue;
    const area = areaOf(/<location\b[^>]*\sref="([^"]+)"/.exec(wb.pkg.text(p.path) || '')?.[1]);
    if (!area) continue;
    const [l0, l1, p0, p1] = axis === 'row' ? [area.r0, area.r1, area.c0, area.c1] : [area.c0, area.c1, area.r0, area.r1];
    if (band) {
      const touched = p1 >= band[0] && p0 <= band[1] && l1 >= at;
      if (!touched) continue;
      if (band[0] > p0 || band[1] < p1) refuse(PIVOT_LINE_MESSAGE);
    }
    if (n > 0) {
      if (at <= l0) plans.push({ p, area, shift: n });
      else if (at <= l1) refuse(PIVOT_LINE_MESSAGE);
      continue;
    }
    const last = at - n - 1;
    if (at > l1) continue;
    if (last < l0) plans.push({ p, area, shift: n });
    else if (at <= l0 && last >= l1) plans.push({ p, gone: true });
    else refuse(PIVOT_LINE_MESSAGE);
  }
  return plans;
}

/** Make the plans (planPivots): each pivot's location moved, a deleted one's part and relationship removed. */
export function applyPivots(wb, si, axis, plans) {
  const pkg = wb.pkg;
  for (const plan of plans) {
    const { p } = plan;
    if (plan.gone) {
      const rel = pkg.readRels(wb.sheets[si].path).find((r) => r.target === p.path);
      if (rel) dropRel(pkg, wb.sheets[si].path, rel.id);
      pkg.removeTree(p.path);
      continue;
    }
    const a = { ...plan.area };
    if (axis === 'row') { a.r0 += plan.shift; a.r1 += plan.shift; } else { a.c0 += plan.shift; a.c1 += plan.shift; }
    pkg.setText(p.path, pkg.text(p.path).replace(/(<location\b[^>]*\sref=")[^"]*"/, `$1${refOf(a)}"`));
  }
}

/** Every pivot cache whose source is a range of the sheet named `sheet`: its range as `move(area)` gives it (null:
 *  all of it deleted, left as it was). */
export function movePivotSources(wb, sheet, move) {
  const pkg = wb.pkg;
  const unescape = (s) => s.replace(/&apos;/g, "'").replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
  for (const path of Object.keys(pkg.files).filter((k) => /^xl\/pivotCache\/pivotCacheDefinition[^/]*\.xml$/i.test(k))) {
    const text = pkg.text(path);
    const out = text.replace(/<worksheetSource\b[^>]*\/>/, (el) => {
      const name = /\ssheet="([^"]*)"/.exec(el)?.[1];
      const ref = /\sref="([^"]*)"/.exec(el)?.[1];
      if (name == null || ref == null || unescape(name).toLowerCase() !== sheet.toLowerCase()) return el;
      const area = areaOf(ref);
      const next = area && move(area);
      return next ? el.replace(` ref="${ref}"`, ` ref="${refOf(next)}"`) : el;
    });
    if (out !== text) pkg.setText(path, out);
  }
}
