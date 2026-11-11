// Other people's cursors in a co-edited workbook (collab.js): each editor puts its selection (the sheet, the active
// cell and the ranges) in the provider's awareness as it moves, at most every `wait` milliseconds, and draws the others'
// on the sheet shown. Drawn as Excel for the web shows its co-authors (Microsoft's "Collaborate on Excel workbooks at
// the same time with co-authoring"): a coloured outline round the other person's cell and their name in that colour.
// Nothing here reaches the workbook: a cursor is view state, carried by awareness only, gone when its editor leaves.
import { areasOf } from './editor/selection.js';

/** The colours given to people whose `user` has none: each at least 4.5:1 against white paper (the sheet stays white
 *  in dark mode too), so the outline and the flag read on the grid, and none of them the selection's own green. */
export const PEER_COLORS = ['#2b6cdf', '#c0392b', '#7b3fb5', '#b25e00', '#0f7c8c', '#c2185b', '#4d5bd6', '#6d6d00'];

/** At most this many ranges of a selection are sent (a long Ctrl+click selection is still shown by its last ones). */
const MAX_AREAS = 32;

const hex = (s) => (typeof s === 'string' && /^#(?:[0-9a-f]{3}|[0-9a-f]{6})$/i.test(s.trim()) ? s.trim().toLowerCase() : null);

/** A colour for awareness client `id` when its user brings none. */
export const peerColor = (id) => PEER_COLORS[Math.abs(Number(id) || 0) % PEER_COLORS.length];

/** The WCAG relative luminance of '#rgb' or '#rrggbb'. */
export function luminance(color) {
  let h = hex(color).slice(1);
  if (h.length === 3) h = [...h].map((x) => x + x).join('');
  const [r, g, b] = [0, 2, 4].map((i) => {
    const v = parseInt(h.slice(i, i + 2), 16) / 255;
    return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

/** The WCAG contrast ratio of two colours. */
export function contrast(a, b) {
  const [x, y] = [luminance(a), luminance(b)].sort((p, q) => q - p);
  return (x + 0.05) / (y + 0.05);
}

/** The text colour for a name on a flag of `color`: black or white, whichever contrasts more. */
export const inkOn = (color) => (contrast(color, '#000000') >= contrast(color, '#ffffff') ? '#000000' : '#ffffff');

/** What this editor sends: { sheet (its sheetId), ar, ac, areas: [{ r0, c0, r1, c1 }] }, or null with nothing open. */
export function cursorOf(api) {
  const sel = api.selection;
  const sheet = api.workbook?.sheets[api.sheetIndex];
  if (!sel || !sheet) return null;
  return { sheet: sheet.sheetId, ar: sel.ar, ac: sel.ac, areas: areasOf(sel).slice(-MAX_AREAS) };
}

const whole = (n) => Number.isInteger(n) && n >= 0;
const fair = (a) => a && whole(a.r0) && whole(a.c0) && whole(a.r1) && whole(a.c1) && a.r1 >= a.r0 && a.c1 >= a.c0;

/** The others' cursors on sheet `sheetId`, from awareness states (a Map of client id to state): [{ id, name, color,
 *  ink, ar, ac, areas }], in client order. States that aren't well formed (another program's, an older editor's) are
 *  passed over. */
export function peersOn(states, self, sheetId) {
  const out = [];
  for (const [id, s] of states) {
    const cur = s?.cursor;
    if (id === self || !cur || cur.sheet !== sheetId || !whole(cur.ar) || !whole(cur.ac)) continue;
    const color = hex(s.user?.color) || peerColor(id);
    const name = String(s.user?.name || 'Guest').slice(0, 40);
    out.push({ id, name, color, ink: inkOn(color), ar: cur.ar, ac: cur.ac, areas: (Array.isArray(cur.areas) ? cur.areas : []).filter(fair).slice(-MAX_AREAS) });
  }
  return out.sort((a, b) => a.id - b.id);
}

/** `fn` called at once, then at most once every `wait` milliseconds, the last call never lost. `.cancel()` drops a
 *  call still to come. (`clock`: { now, later, clear } for the tests.) */
export function throttle(fn, wait, clock = { now: () => Date.now(), later: (f, ms) => setTimeout(f, ms), clear: (t) => clearTimeout(t) }) {
  let last = -Infinity;
  let timer = null;
  const run = () => { timer = null; last = clock.now(); fn(); };
  const call = () => {
    if (timer != null) return;
    const left = last + wait - clock.now();
    if (left <= 0) run();
    else timer = clock.later(run, left);
  };
  call.cancel = () => { if (timer != null) clock.clear(timer); timer = null; };
  return call;
}

/** Share this editor's cursor through `awareness` and draw the others' in `api`'s grid. Returns { redraw(), stop() };
 *  stop() takes this editor's cursor away from the others at once. */
export function shareCursors(api, awareness, { wait = 100 } = {}) {
  const send = throttle(() => awareness.setLocalStateField('cursor', cursorOf(api)), wait);
  const redraw = () => {
    if (!api.grid || !api.workbook) return;
    api.grid.setPeers(peersOn(awareness.getStates(), awareness.clientID, api.workbook.sheets[api.sheetIndex]?.sheetId));
  };
  const moved = () => { send(); redraw(); };
  const stopSelect = api.onSelect(moved);
  awareness.on('change', redraw);
  moved();
  return {
    redraw,
    stop() {
      stopSelect();
      send.cancel();
      awareness.off('change', redraw);
      awareness.setLocalStateField('cursor', null);
      api.grid?.setPeers([]);
    },
  };
}
