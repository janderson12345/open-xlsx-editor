// Co-editing (src/collab.js): two editors, each in its own page, their Y.Docs relayed to each other as a provider
// would. The second opens the first's workbook; contents typed in either reach the other, also when typed at once; a
// row inserted (shared as the whole workbook) reaches the other, and typing goes on in step after it.
// Run after `npm run build`: node test/collab.mjs
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
  window.connect = async () => { window.room = await m.connectWorkbook(window.oxe, doc); };
});
await a.page.exposeFunction('relay', (u) => b.page.evaluate((x) => window.applyRelay(x), u));
await b.page.exposeFunction('relay', (u) => a.page.evaluate((x) => window.applyRelay(x), u));
await setup(a.page);
await setup(b.page);
await a.page.evaluate(() => window.connect());
await b.page.evaluate(() => window.connect());

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

const bytes = await a.page.evaluate(() => window.oxe.save().length);
t.check(bytes > 0 && !a.errors.length && !b.errors.length, 'no page errors', [...a.errors, ...b.errors].join('\n     '));
await a.close();
await b.close();
console.log(t.failed ? `${t.failed} failed` : 'all passed');
process.exit(t.failed ? 1 : 0);
