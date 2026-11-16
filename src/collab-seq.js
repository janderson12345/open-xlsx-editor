// Which row (or column) is which, in a co-edited sheet (collab.js): each line has an identity that stays with it when
// rows are inserted or deleted above it, in any editor, so a cell's contents, its format and a row's height are kept
// in the room under the line's identity, not under its number. A line of the room's workbook is known by its number
// in that workbook ('12'); a line inserted since by the insert that made it and its place in it ('<client>.<n>:3').
//
// The order of the lines is a replicated sequence (RGA, Roh et al. 2011): an insert goes after a line (its anchor),
// past any lines a later insert put there; a delete marks lines dead without removing them, so a concurrent insert
// next to a deleted line still finds its place. Two editors applying the same inserts and deletes, in any order,
// end with the same lines in the same order. Lines are kept as runs (one per insert, split where needed): an insert
// of 1,000 rows is one run, and the workbook's own lines are a single run without end.

const BASE = 'b';

/** A line's identity: a line of the room's workbook by its number, an inserted one by its insert and its place. */
export const tokenOf = (id, k) => (id === BASE ? String(k) : `${id}:${k}`);
const parse = (tok) => {
  const i = tok.lastIndexOf(':');
  return i < 0 ? { id: BASE, k: Number(tok) } : { id: tok.slice(0, i), k: Number(tok.slice(i + 1)) };
};

/** (ts, cid) ordering of inserts: the later (or, at the same time, the higher client) first after their anchor. */
const greater = (a, b) => a.ts > b.ts || (a.ts === b.ts && a.cid > b.cid);

export class LineSeq {
  constructor() {
    /** Runs of lines: { id (BASE or the insert's), start (its first line's place in that insert), len, ts, cid, dead }. */
    this.runs = [{ id: BASE, start: 0, len: Infinity, ts: 0, cid: 0, dead: false }];
  }

  /** The identity of the line shown at `pos` (0-based). */
  tokAt(pos) {
    let at = 0;
    for (const s of this.runs) {
      if (s.dead) continue;
      if (pos < at + s.len) return tokenOf(s.id, s.start + (pos - at));
      at += s.len;
    }
    return null;
  }

  /** Where the line `tok` is shown ({ pos, alive }), or null for one this sequence has never had. A dead line's pos is
   *  where it would be: the place of the next line alive. */
  posOf(tok) {
    const { id, k } = parse(tok);
    let at = 0;
    for (const s of this.runs) {
      if (s.id === id && k >= s.start && k < s.start + s.len) return { pos: s.dead ? at : at + (k - s.start), alive: !s.dead };
      if (!s.dead) at += s.len;
    }
    return null;
  }

  /** Split runs so that one starts at line `tok` (`after` false) or right after it (true); the index of that run. */
  _cut(tok, after) {
    const { id, k } = parse(tok);
    const i = this.runs.findIndex((s) => s.id === id && k >= s.start && k < s.start + s.len);
    if (i < 0) return -1;
    const s = this.runs[i];
    const at = after ? k + 1 : k;
    if (at === s.start) return i;
    if (at >= s.start + s.len) return i + 1;
    const rest = { ...s, start: at, len: s.len - (at - s.start) };
    s.len = at - s.start;
    this.runs.splice(i + 1, 0, rest);
    return i + 1;
  }

  /** The runs of `count` lines shown from `pos` on ([{ id, start, len }]): what a delete names. */
  span(pos, count) {
    const out = [];
    let at = 0;
    for (const s of this.runs) {
      if (s.dead) continue;
      const a = Math.max(pos, at);
      const b = Math.min(pos + count, at + s.len);
      if (a < b) out.push({ id: s.id, start: s.start + (a - at), len: b - a });
      at += s.len;
      if (at >= pos + count) break;
    }
    return out;
  }

  /** An insert { id, ts, cid, after (the anchor's identity, null for the first line), n }: where its first line is
   *  shown, or -1 when it was integrated before (an editor's own, applied when made). */
  insert(op) {
    if (this.runs.some((s) => s.id === op.id)) return -1;
    let i = op.after == null ? 0 : this._cut(op.after, true);
    if (i < 0) i = 0;
    // (Past the lines a later insert put after the same anchor: RGA's rule, so the order is the same everywhere.)
    while (i < this.runs.length && this.runs[i].id !== BASE && greater(this.runs[i], op)) i++;
    this.runs.splice(i, 0, { id: op.id, start: 0, len: op.n, ts: op.ts, cid: op.cid, dead: false });
    let pos = 0;
    for (let k = 0; k < i; k++) if (!this.runs[k].dead) pos += this.runs[k].len;
    return pos;
  }

  /** A delete of lines `runs` ([{ id, start, len }]): the places of those still shown, as [{ pos, n }] from the last
   *  to the first (the order to delete them in), and the lines marked dead. */
  remove(runs) {
    const hit = [];
    for (const r of runs) {
      const a = this._cut(tokenOf(r.id, r.start), false);
      if (a < 0) continue;
      this._cut(tokenOf(r.id, r.start + r.len - 1), true);
      for (let i = a; i < this.runs.length && this.runs[i].id === r.id && this.runs[i].start < r.start + r.len; i++) if (!this.runs[i].dead) hit.push(this.runs[i]);
    }
    const out = [];
    let at = 0;
    for (const s of this.runs) {
      if (s.dead) continue;
      if (hit.includes(s)) {
        const last = out[out.length - 1];
        if (last && last.pos + last.n === at) last.n += s.len; else out.push({ pos: at, n: s.len });
      }
      at += s.len;
    }
    for (const s of hit) s.dead = true;
    // (Positions as shown before the delete: taken from the last back, each leaves the ones before it in place.)
    return out.reverse();
  }
}
