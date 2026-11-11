// Co-editing (src/collab.js): two editors, each in its own page, their Y.Docs relayed to each other as a provider
// would. The second opens the first's workbook; contents typed in either reach the other, also when typed at once; a
// row inserted (shared as the whole workbook) reaches the other, and typing goes on in step after it. Each sees the
// other's cursor (outline, shaded range, name on a flag) on the sheet it is on, and loses it when the other leaves.
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
  window.applyAware = (id, s) => { if (s === null) states.delete(id); else states.set(id, s); emit(); };
  window.connect = async (user) => { window.room = await m.connectWorkbook(window.oxe, doc, { awareness: aware, user }); };
});
await a.page.exposeFunction('relay', (u) => b.page.evaluate((x) => window.applyRelay(x), u));
await b.page.exposeFunction('relay', (u) => a.page.evaluate((x) => window.applyRelay(x), u));
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
t.check(await bold(a.page) && await bold(b.page) && await shown(b.page, 2, 7) === '42', 'H3 made bold in the first (shared as the workbook): bold in the second, its 42 kept');

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
await b.page.waitForFunction(() => document.querySelector('.oxe-overlay .oxe-peer-range'), null, { timeout: 5000 }).catch(() => {});
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
