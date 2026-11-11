import { test } from 'node:test';
import assert from 'node:assert/strict';
import { PEER_COLORS, peerColor, contrast, inkOn, cursorOf, peersOn, throttle } from '../src/collab-peers.js';
import { addArea, extendSelection, cellSelection } from '../src/editor/selection.js';

// Other people's cursors (collab-peers.js): what is sent, what is drawn from the others' states, and the throttle.
test('the colours given out read on white paper and the flag text contrasts with each', () => {
  for (const c of PEER_COLORS) {
    assert.ok(contrast(c, '#ffffff') >= 4.5, `${c} on white: ${contrast(c, '#ffffff').toFixed(2)}`);
    assert.ok(contrast(c, inkOn(c)) >= 4.5, `${inkOn(c)} on ${c}`);
  }
  assert.equal(contrast('#000000', '#ffffff').toFixed(1), '21.0');
  assert.equal(inkOn('#ffd700'), '#000000', 'black on a light yellow');
  assert.equal(inkOn('#c0392b'), '#ffffff', 'white on dark red');
  assert.equal(inkOn('#fff'), '#000000', 'short form');
  assert.equal(peerColor(9), PEER_COLORS[9 % PEER_COLORS.length]);
});

test('the cursor sent: the sheet by its id, the active cell and every range', () => {
  const sel = addArea(extendSelection(cellSelection(4, 2), 5, 3), 9, 0);
  const api = { selection: sel, sheetIndex: 1, workbook: { sheets: [{ sheetId: 1 }, { sheetId: 7 }] } };
  assert.deepEqual(cursorOf(api), { sheet: 7, ar: 9, ac: 0, areas: [{ r0: 4, c0: 2, r1: 5, c1: 3 }, { r0: 9, c0: 0, r1: 9, c1: 0 }] });
  assert.equal(cursorOf({ selection: null, workbook: null, sheetIndex: 0 }), null);
});

test('the others on the sheet shown, well-formed states only, never this editor itself', () => {
  const cur = (sheet, ar, ac, areas = [{ r0: ar, c0: ac, r1: ar, c1: ac }]) => ({ sheet, ar, ac, areas });
  const states = new Map([
    [5, { user: { name: 'Avery Lin', color: '#C0392B' }, cursor: cur(1, 2, 3) }],
    [2, { user: { name: 'Rowan Hale' }, cursor: cur(1, 0, 0, [{ r0: 0, c0: 0, r1: 3, c1: 1 }, { r0: 2, c0: 2, r1: 1, c1: 2 }]) }],
    [3, { user: { name: 'Sam Ito' }, cursor: cur(2, 1, 1) }],
    [4, { user: { name: 'Me' }, cursor: cur(1, 1, 1) }],
    [6, { user: { name: 'Odd' }, cursor: { sheet: 1, ar: -1, ac: 'x' } }],
    [7, { user: { name: 'Left' }, cursor: null }],
    [8, { user: { name: 'Bad colour', color: 'red; background: url(x)' }, cursor: cur(1, 4, 4) }],
  ]);
  const out = peersOn(states, 4, 1);
  assert.deepEqual(out.map((p) => p.name), ['Rowan Hale', 'Avery Lin', 'Bad colour'], 'in client order; other sheets, this editor, bad and empty cursors passed over');
  assert.equal(out[0].color, peerColor(2), 'no colour of their own: one from the list');
  assert.deepEqual(out[0].areas, [{ r0: 0, c0: 0, r1: 3, c1: 1 }], 'a range with its corners the wrong way round passed over');
  assert.equal(out[1].color, '#c0392b');
  assert.equal(out[1].ink, '#ffffff');
  assert.equal(out[2].color, peerColor(8), 'anything but a hex colour replaced');
  assert.equal(peersOn(new Map([[1, { cursor: cur(1, 0, 0) }]]), 9, 1)[0].name, 'Guest');
});

test('the throttle: at once, then at most once a period, the last call kept', () => {
  let now = 0;
  const timers = [];
  const clock = { now: () => now, later: (f, ms) => { timers.push({ f, at: now + ms }); return timers.length; }, clear: (i) => { timers[i - 1].f = null; } };
  let calls = 0;
  const go = throttle(() => calls++, 100, clock);
  go();
  assert.equal(calls, 1, 'the first at once');
  now = 30; go(); now = 60; go();
  assert.equal(calls, 1);
  assert.equal(timers.length, 1, 'one call waiting for the rest of the period');
  assert.equal(timers[0].at, 100);
  now = 100; timers[0].f();
  assert.equal(calls, 2, 'the last one made at the end of the period');
  now = 250; go();
  assert.equal(calls, 3, 'a period later: at once again');
  now = 260; go(); go.cancel();
  assert.equal(timers[1].f, null, 'cancel drops the call to come');
});
