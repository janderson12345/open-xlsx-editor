// Co-editing (src/collab.js): two editors, each in its own page, their Y.Docs relayed to each other as a provider
// would (and held back, to make changes at the same moment). The second opens the first's workbook; contents typed
// in either reach the other, also when typed at once; rows inserted, formats, sizes and sheets reach the other piece
// by piece (the room's epoch unchanged), also when made at once in both; a merge goes as the whole workbook. Each
// sees the other's cursor (outline, shaded range, name on a flag) on the sheet it is on, and loses it when it leaves.
// Run after `npm run build`: node test/collab.mjs
import { createRequire } from 'node:module';
import { openDemo, openFixture, checker } from './browser.mjs';

const t = checker();
const a = await openDemo();
const b = await openDemo();
await openFixture(a.page, 'basic.xlsx');

// (A Y.Doc in each page; each update made in one is applied in the other, marked so it isn't sent back.)
const setup = (page) => page.evaluate(async () => {
  const m = await import('/dist/collab.js');
  const doc = new m.Y.Doc();
  window.ydoc = doc;
  doc.on('update', (u, origin) => { if (origin !== 'relay') window.relay(Array.from(u)); });
  window.applyRelay = (u) => m.Y.applyUpdate(doc, new Uint8Array(u), 'relay');
  // (A stand-in for y-protocols' Awareness, the part connectWorkbook uses: each client's state, relayed as the
  // provider would; a state set to null is that client gone.)
  const listeners = new Set();
  const states = new Map();
  const emit = () => { for (const f of [...listeners]) f({}, 'local'); };
  const aware = {
    clientID: doc.clientID,
    getStates: () => states,
    on: (_, f) => listeners.add(f),
    off: (_, f) => listeners.delete(f),
    setLocalStateField(k, v) {
      const s = { ...(states.get(doc.clientID) || {}), [k]: v };
      states.set(doc.clientID, s);
      window.relayAware(doc.clientID, s);
      emit();
    },
  };
  window.aware = aware;
  window.applyAware = (id, s) => { if (s === null) states.delete(id); else states.set(id, s); emit(); };
  window.connect = async (user) => { window.room = await m.connectWorkbook(window.oxe, doc, { awareness: aware, user }); };
});
// (`held`: updates kept back while set, sent in order by release().)
let held = null;
const send = (to, u) => (held ? held.push([to, u]) : to.page.evaluate((x) => window.applyRelay(x), u));
const hold = () => { held = []; };
const release = async () => { const list = held; held = null; for (const [to, u] of list) await to.page.evaluate((x) => window.applyRelay(x), u); };
await a.page.exposeFunction('relay', (u) => send(b, u));
await b.page.exposeFunction('relay', (u) => send(a, u));
await a.page.exposeFunction('relayAware', (id, s) => b.page.evaluate(([i, x]) => window.applyAware(i, x), [id, s]));
await b.page.exposeFunction('relayAware', (id, s) => a.page.evaluate(([i, x]) => window.applyAware(i, x), [id, s]));
await setup(a.page);
await setup(b.page);
await a.page.evaluate(() => window.connect({ name: 'Avery Lin', color: '#c0392b' }));
await b.page.evaluate(() => window.connect({ name: 'Rowan Hale' }));

const shown = (page, r, c) => page.evaluate(([r, c]) => document.querySelector(`.oxe-row[aria-rowindex="${r + 2}"] .oxe-cell[aria-colindex="${c + 2}"]`)?.textContent ?? null, [r, c]);
const waitShown = async (page, r, c, text) => {
  try { await page.waitForFunction(([r, c, x]) => document.querySelector(`.oxe-row[aria-rowindex="${r + 2}"] .oxe-cell[aria-colindex="${c + 2}"]`)?.textContent === x, [r, c, text], { timeout: 5000 }); return true; } catch { return false; }
};
const sheets = (page) => page.evaluate(() => window.oxe.workbook?.sheets.map((s) => s.name).join() ?? null);

t.check(await sheets(b.page) === await sheets(a.page) && await sheets(a.page) !== null, 'the second editor opens the room’s workbook', `${await sheets(a.page)} / ${await sheets(b.page)}`);
const epochOf = (page) => page.evaluate(() => window.ydoc.getMap('oxe-files').get('epoch'));
const epoch0 = await epochOf(a.page);

await a.page.evaluate(() => window.oxe.enter('H3', '42'));
t.check(await waitShown(b.page, 2, 7, '42'), 'typed in the first (H3 42): shown in the second', await shown(b.page, 2, 7));

await b.page.evaluate(() => window.oxe.enter('H4', '=H3*2'));
t.check(await waitShown(a.page, 3, 7, '84'), 'a formula typed in the second (H4 =H3*2): worked out in the first', await shown(a.page, 3, 7));

// (Rows 8 and 9: basic.xlsx hides one of the rows above.)
await Promise.all([a.page.evaluate(() => window.oxe.enter('H8', 'from one')), b.page.evaluate(() => window.oxe.enter('H9', 'from two'))]);
const both = async (page) => (await waitShown(page, 7, 7, 'from one')) && (await waitShown(page, 8, 7, 'from two'));
t.check(await both(a.page) && await both(b.page), 'typed in both at once (H8, H9): both in each');

await a.page.evaluate(() => window.oxe.command('select', 2, 0));
await a.page.evaluate(() => window.oxe.command('lines', 'row', 1));
t.check(await waitShown(b.page, 3, 7, '42'), 'a row inserted above row 3 in the first: the second’s 42 now on row 4', await shown(b.page, 3, 7));

await b.page.evaluate(() => window.oxe.enter('H1', 'after'));
t.check(await waitShown(a.page, 0, 7, 'after'), 'typing goes on in step after the row (H1 in the second, shown in the first)', await shown(a.page, 0, 7));

await a.page.evaluate(() => window.oxe.undo());
t.check(await waitShown(b.page, 2, 7, '42') && await waitShown(a.page, 2, 7, '42'), 'Ctrl+Z of the row in the first: 42 back on row 3 in both');

await a.page.evaluate(() => { window.oxe.command('select', 2, 7); window.oxe.command('toggleFont', 'b'); });
const bold = (page) => page.waitForFunction(() => Number(getComputedStyle(document.querySelector('.oxe-row[aria-rowindex="4"] .oxe-cell[aria-colindex="9"]')).fontWeight) >= 600, null, { timeout: 5000 }).then(() => true, () => false);
t.check(await bold(a.page) && await bold(b.page) && await shown(b.page, 2, 7) === '42', 'H3 made bold in the first: bold in the second, its 42 kept');
t.check(await epochOf(a.page) === epoch0 && await epochOf(b.page) === epoch0, 'the row, its undo and the bold went across piece by piece: the room’s epoch unchanged');

// Sizes and hiding: row 6 30pt high and column J hidden in the first, in the second; Undo of both in both.
const model = (page, fn, arg = null) => page.evaluate(([f, x]) => new Function('wb', 'x', `return (${f})(wb, x)`)(window.oxe.workbook, x), [fn.toString(), arg]);
const until = async (page, fn, arg, want) => {
  for (let k = 0; k < 50; k++) { if (JSON.stringify(await model(page, fn, arg)) === JSON.stringify(want)) return true; await page.waitForTimeout(100); }
  return false;
};
const height = (wb, r) => wb.sheet(0).rows.get(r)?.ht ?? null;
const hiddenJ = (wb) => !!wb.sheet(0).cols.find((x) => x.min <= 9 && x.max >= 9)?.hidden;
await a.page.evaluate(() => { window.oxe.command('select', 5, 0); window.oxe.command('selectRows'); window.oxe.command('rowHeight', 30); });
await a.page.evaluate(() => { window.oxe.command('select', 0, 9); window.oxe.command('selectColumns'); window.oxe.command('hideLines', 'col', true); });
t.check(await until(b.page, height, 5, 30) && await until(b.page, hiddenJ, null, true), 'row 6 made 30pt and column J hidden in the first: so in the second', JSON.stringify([await model(b.page, height, 5), await model(b.page, hiddenJ)]));
await a.page.evaluate(() => { window.oxe.undo(); window.oxe.undo(); });
const h6 = await model(a.page, height, 5);
t.check(h6 !== 30 && await until(b.page, height, 5, h6) && await until(b.page, hiddenJ, null, false), 'both undone in the first: undone in the second', JSON.stringify([h6, await model(b.page, height, 5)]));

// At the same moment: each inserts a row above row 11 and types in it; then the first makes “mark” italic while the
// second inserts a row at the top. Both end with the same cells, the italic on the same contents.
const column = (wb, c) => Array.from({ length: 16 }, (_, r) => wb.sheet(0).cells.get(r)?.get(c)?.v ?? null);
await a.page.evaluate(() => window.oxe.enter('B12', 'mark'));
await waitShown(b.page, 11, 1, 'mark');
hold();
await a.page.evaluate(() => { window.oxe.command('select', 10, 0); window.oxe.command('lines', 'row', 1); window.oxe.enter('A11', 'one'); });
await b.page.evaluate(() => { window.oxe.command('select', 10, 0); window.oxe.command('lines', 'row', 1); window.oxe.enter('A11', 'two'); });
await release();
await a.page.waitForTimeout(200);
const colA = await model(a.page, column, 0);
t.check(await until(b.page, column, 0, colA) && colA.includes('one') && colA.includes('two'), 'a row inserted above row 11 in both at once, typed in: both rows in both, in one order', JSON.stringify(colA));
const markAt = await model(a.page, (wb) => Array.from({ length: 30 }, (_, r) => wb.sheet(0).cells.get(r)?.get(1)?.v ?? null).indexOf('mark'));
hold();
await a.page.evaluate((r) => { window.oxe.command('select', r, 1); window.oxe.command('toggleFont', 'i'); }, markAt);
await b.page.evaluate(() => { window.oxe.command('select', 0, 0); window.oxe.command('lines', 'row', 1); });
await release();
const italic = (wb) => { const r = [...wb.sheet(0).cells].find(([, row]) => row.get(1)?.v === 'mark')?.[0]; const s = wb.sheet(0).cells.get(r).get(1).s; return [r, !!wb.styles.fonts[wb.styles.cellXfs[s].fontId].i]; };
t.check(await until(a.page, italic, null, [markAt + 1, true]) && await until(b.page, italic, null, [markAt + 1, true]), '“mark” made italic in the first while the second inserts a row at the top: italic, a row down, in both', JSON.stringify([await model(a.page, italic), await model(b.page, italic)]));
t.check(JSON.stringify(await model(a.page, column, 0)) === JSON.stringify(await model(b.page, column, 0)), 'and column A the same in both');

// Sheets: the first adds one after the sheet shown, renames it, colours it and moves it first; the second hides
// another. Each sees the other's; the first's Undo of the move moves it back in both.
const tabs = (wb) => wb.sheets.map((s) => `${s.name}:${s.state}`).join();
await a.page.evaluate(() => {
  window.oxe.command('insertSheet', { after: true });
  const k = window.oxe.sheetIndex;
  window.oxe.command('renameSheet', k, 'Shared');
  window.oxe.command('tabColor', k, { rgb: 'FF00B050' });
  window.oxe.command('moveSheet', k, 0);
});
await b.page.evaluate(() => window.oxe.command('hideSheet', window.oxe.workbook.sheets.findIndex((s) => s.name === 'Notes & more')));
await a.page.waitForTimeout(200);
const tabsA = await model(a.page, tabs);
t.check(await until(b.page, tabs, null, tabsA) && tabsA.startsWith('Shared:visible') && tabsA.includes('Notes & more:hidden'), 'a sheet added, renamed, coloured and moved first in the first, another hidden in the second: the same tabs in both', `${tabsA} / ${await model(b.page, tabs)}`);
const green = (wb) => wb.tabColor(wb.sheets.findIndex((s) => s.name === 'Shared'));
t.check(await until(b.page, green, null, '#00B050'), 'its tab colour in the second', await model(b.page, green));
await a.page.evaluate(() => window.oxe.undo());
const back = await model(a.page, tabs);
t.check(!back.startsWith('Shared') && await until(b.page, tabs, null, back), 'the move undone in the first: moved back in both', `${back} / ${await model(b.page, tabs)}`);
await a.page.evaluate(() => window.oxe.redo());
t.check(await until(b.page, tabs, null, tabsA) && await model(a.page, tabs) === tabsA, 'and redone: first again in both', await model(b.page, tabs));
await a.page.evaluate(() => window.oxe.undo());
await until(b.page, tabs, null, back);
await a.page.evaluate(() => window.oxe.showSheet(window.oxe.workbook.sheets.findIndex((s) => s.name === 'Data')));
await b.page.evaluate(() => window.oxe.showSheet(window.oxe.workbook.sheets.findIndex((s) => s.name === 'Data')));
// A row deleted in the second and its Undo: the row's cells come back in both, while the first types further down.
const rowText = (wb, r) => Array.from({ length: 6 }, (_, c) => wb.sheet(0).cells.get(r)?.get(c)?.v ?? null).join('|');
const row4 = await model(b.page, rowText, 3);
const row5 = await model(b.page, rowText, 4);
await b.page.evaluate(() => { window.oxe.command('select', 3, 0); window.oxe.command('selectRows'); window.oxe.command('lines', 'row', -1); });
t.check(await until(a.page, rowText, 3, row5), 'row 4 deleted in the second: row 5 moved up in the first', await model(a.page, rowText, 3));
await a.page.evaluate(() => window.oxe.enter('K30', 'below'));
await waitShown(b.page, 29, 10, 'below');
await b.page.evaluate(() => window.oxe.undo());
t.check(await until(a.page, rowText, 3, row4) && await until(b.page, rowText, 3, row4) && await waitShown(a.page, 30, 10, 'below'), 'its Undo in the second: row 4’s cells back in both, the first’s K30 moved down with its row', `${await model(a.page, rowText, 3)} / ${row4}`);
t.check(await epochOf(a.page) === epoch0 && await epochOf(b.page) === epoch0, 'sizes, rows inserted at once, the italic, the sheets and the deleted row went across piece by piece: the epoch unchanged');

// A merge isn't held piece by piece: the whole workbook goes, as a new epoch, and typing goes on after it.
await a.page.evaluate(() => { window.oxe.command('select', 20, 0); window.oxe.command('extendTo', 20, 2); window.oxe.command('merge', 'cells'); });
const merged = (wb) => wb.sheet(wb.sheets.findIndex((s) => s.name === 'Data')).merges.some((m) => m.r0 === 20 && m.c0 === 0 && m.c1 === 2);
t.check(await until(b.page, merged, null, true) && await epochOf(b.page) !== epoch0, 'A21:C21 merged in the first: merged in the second, shared as the whole workbook');
await b.page.evaluate(() => window.oxe.enter('J2', 'still'));
t.check(await waitShown(a.page, 1, 9, 'still'), 'typing goes on after it (J2 in the second, shown in the first)', await shown(a.page, 1, 9));
// (Both back at the sheet's top-left corner for the cursors' places below.)
for (const p of [a.page, b.page]) await p.evaluate(() => { window.oxe.command('select', 0, 0); window.oxe.grid.el.scrollTop = 0; window.oxe.grid.el.scrollLeft = 0; });

// Cursors: the first selects C5:D6 (C5 active); the second sees that range and C5 outlined in the first's colour,
// with the first's name on a flag, at the cells' own place; neither draws its own cursor as someone else's.
const peer = (page) => page.evaluate(() => {
  const el = document.querySelector('.oxe-grid > .oxe-sheet > .oxe-overlay .oxe-peer-active');
  if (!el) return null;
  const flag = el.querySelector('.oxe-peer-flag');
  const box = el.getBoundingClientRect();
  const cell = window.oxe.grid.cellRect(4, 2);
  const g = window.oxe.grid.geometry();
  const sheet = document.querySelector('.oxe-grid > .oxe-sheet').getBoundingClientRect();
  return {
    name: el.dataset.peer, flag: flag?.textContent, flagShown: !!flag && flag.getBoundingClientRect().width > 0,
    border: getComputedStyle(el).borderTopColor, ink: getComputedStyle(flag).color, back: getComputedStyle(flag).backgroundColor,
    // (Where the outline is against where cell C5 is, in pixels from the sheet's corner.)
    dx: Math.round(box.left - sheet.left - (g.rowHeadW + cell.x - 1)), dy: Math.round(box.top - sheet.top - (g.colHeadH + cell.y - 1)),
    ranges: document.querySelectorAll('.oxe-grid > .oxe-sheet > .oxe-overlay .oxe-peer-range').length,
  };
});
const waitPeer = (page, gone = false) => page.waitForFunction((g) => !document.querySelector('.oxe-grid > .oxe-sheet > .oxe-overlay .oxe-peer-active') === g, gone, { timeout: 5000 }).then(() => true, () => false);
await a.page.evaluate(() => { window.oxe.command('select', 4, 2); window.oxe.command('extendTo', 5, 3); });
// (Waited for until the second has the first's C5:D6 in its awareness and has drawn it: not just for any shaded range,
// which the first's A21:C21 from the merge above can still be when the select is sent, the moves coalesced.)
await b.page.waitForFunction(() => {
  const me = window.aware.clientID;
  const cur = [...window.aware.getStates()].find(([id]) => id !== me)?.[1]?.cursor;
  const range = document.querySelector('.oxe-grid > .oxe-sheet > .oxe-overlay .oxe-peer-range');
  return cur?.ar === 4 && cur.ac === 2 && cur.areas?.length === 1 && cur.areas[0].r1 === 5 && cur.areas[0].c1 === 3 && range;
}, null, { timeout: 5000 }).catch(() => {});
const seen = await peer(b.page);
t.check(seen?.name === 'Avery Lin' && seen.flag === 'Avery Lin' && seen.flagShown && seen.border === 'rgb(192, 57, 43)' && seen.dx === 0 && seen.dy === 0 && seen.ranges === 1,
  'C5:D6 selected in the first: the second outlines C5 and shades the range in the first’s colour, its name on a flag', JSON.stringify(seen));
t.check(seen?.ink === 'rgb(255, 255, 255)' && seen.back === 'rgb(192, 57, 43)', 'the flag’s name in white on the dark red, the colour that contrasts more', JSON.stringify(seen));
const own = await a.page.evaluate(() => [...document.querySelectorAll('.oxe-peer-active')].map((e) => e.dataset.peer));
t.check(own.length === 1 && own[0] === 'Rowan Hale', 'the first sees the second’s cursor and not its own', JSON.stringify(own));

// Contrast: axe on the second page with the flag shown, light and dark.
await b.page.addScriptTag({ path: createRequire(import.meta.url).resolve('axe-core/axe.min.js') });
const axe = () => b.page.evaluate(async () => (await window.axe.run(document, { runOnly: { type: 'tag', values: ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa'] } })).violations.map((v) => `${v.id}: ${v.nodes.slice(0, 2).map((n) => n.target.join(' ')).join(', ')}`));
for (const theme of ['light', 'dark']) {
  await b.page.evaluate((x) => window.oxe.setTheme(x), theme);
  const issues = await axe();
  t.check(!issues.length, `${theme}: a cursor and its flag shown, no axe violations (the flag's own contrast is checked in the unit tests)`, issues.join('\n     '));
}
await b.page.evaluate(() => window.oxe.setTheme(null));

// The first moves to another sheet: its cursor leaves the second's grid, which shows the first sheet still; back, it
// returns. Then the first leaves the room: its cursor goes at once.
await a.page.evaluate(() => window.oxe.showSheet(1));
t.check(await waitPeer(b.page, true), 'the first on another sheet: its cursor not drawn on the second’s sheet');
await a.page.evaluate(() => { window.oxe.showSheet(0); window.oxe.command('select', 6, 1); });
t.check(await waitPeer(b.page) && (await peer(b.page))?.name === 'Avery Lin', 'back on the sheet (B7): drawn again');
await a.page.evaluate(() => window.room.stop());
t.check(await waitPeer(b.page, true), 'the first leaves the room: its cursor gone from the second');
t.check(await waitPeer(a.page, true), 'and the second’s cursor gone from the first that left');

const bytes = await a.page.evaluate(() => window.oxe.save().length);
t.check(bytes > 0 && !a.errors.length && !b.errors.length, 'no page errors', [...a.errors, ...b.errors].join('\n     '));
await a.close();
await b.close();
console.log(t.failed ? `${t.failed} failed` : 'all passed');
process.exit(t.failed ? 1 : 0);
