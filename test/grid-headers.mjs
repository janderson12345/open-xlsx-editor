// Browser check of the headings a screen reader gives each cell, from Chromium's own accessibility tree (the tree NVDA
// reads in Edge and Chrome, through the DevTools protocol): every cell's place in its row is its column heading's place
// in the heading row, so the heading at that place names the cell's own column, and each row's first cell is its
// number. NVDA read column B's cells under C's heading before (WBS POC's UI-A11Y-033): the heading row had no first
// cell where every other row had its number, and empty cells weren't there to be counted. Checked on a plain sheet,
// scrolled right, with a merged range, and with panes frozen.
// Run after `npm run build`: node test/grid-headers.mjs [outDir]
import { editingPage } from './editing-kit.mjs';
import { openFixture } from './browser.mjs';

const { page, check, finish } = await editingPage();
const cdp = await page.context().newCDPSession(page);
await cdp.send('Accessibility.enable');
await cdp.send('DOM.enable');

const CELLS = new Set(['columnheader', 'rowheader', 'gridcell', 'cell']);
/** The main grid's rows as Chromium has them: for each row, its cells in order, each { role, name, col } (col: the
 *  column letter its aria-colindex says, or null). Ignored wrappers are looked through, as a screen reader does. */
async function treeRows() {
  const { nodes } = await cdp.send('Accessibility.getFullAXTree');
  const byId = new Map(nodes.map((n) => [n.nodeId, n]));
  const role = (n) => n.role?.value;
  const grids = nodes.filter((n) => role(n) === 'grid' && !n.ignored);
  const grid = grids[0];
  const rows = [];
  const walk = (n, into) => {
    for (const id of n.childIds || []) {
      const k = byId.get(id);
      if (!k) continue;
      if (!k.ignored && role(k) === 'row') { const cells = []; walkCells(k, cells); rows.push(cells); continue; }
      walk(k, into);
    }
  };
  const walkCells = (n, out) => {
    for (const id of n.childIds || []) {
      const k = byId.get(id);
      if (!k) continue;
      if (!k.ignored && CELLS.has(role(k))) out.push(k);
      else if (k.ignored || !CELLS.has(role(k))) walkCells(k, out);
    }
  };
  if (grid) walk(grid, rows);
  // (Each cell's column as its element says it, read from the DOM node behind it.)
  const out = [];
  for (const cells of rows) {
    const list = [];
    for (const k of cells) {
      let col = null;
      if (k.backendDOMNodeId) {
        const { node } = await cdp.send('DOM.describeNode', { backendNodeId: k.backendDOMNodeId });
        const attrs = node.attributes || [];
        const i = attrs.indexOf('aria-colindex');
        const n = i >= 0 ? Number(attrs[i + 1]) : null;
        if (n && n >= 2) { let c = n - 2; let s = ''; c += 1; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } col = s; }
      }
      list.push({ role: role(k), name: k.name?.value ?? '', col });
    }
    out.push(list);
  }
  return { grids: grids.length, rows: out };
}

/** Every cell's heading by its place: the first row's cells are the headings; for each other row, its first cell a
 *  row heading, and each gridcell's heading at the same place naming its own column. Returns the mismatches. */
function mismatches({ rows }) {
  const heads = rows.find((r) => r.some((c) => c.role === 'columnheader'));
  if (!heads) return ['no heading row'];
  const wrong = [];
  for (const row of rows) {
    if (row === heads) continue;
    if (row[0]?.role !== 'rowheader') { wrong.push(`row starts with ${row[0]?.role}`); continue; }
    row.forEach((cell, p) => {
      if (cell.role !== 'gridcell') return;
      const head = heads[p];
      if (!head || head.role !== 'columnheader' || head.name !== cell.col) wrong.push(`row ${row[0].name}: ${cell.col} at place ${p} under ${head ? `${head.role} "${head.name}"` : 'nothing'}`);
    });
  }
  return wrong;
}

await openFixture(page, 'basic.xlsx');
await page.evaluate(() => window.oxe.select('B1'));
await page.waitForTimeout(200);
let tree = await treeRows();
const heads = tree.rows.find((r) => r.some((c) => c.role === 'columnheader'));
const row1 = tree.rows.find((r) => r[0]?.role === 'rowheader' && r[0].name === '1');
const b1 = row1 ? row1.findIndex((c) => c.col === 'B') : -1;
check(b1 > 0 && heads[b1]?.name === 'B' && row1[0].name === '1', 'B1: its heading is "B", and its row\'s heading "1"', JSON.stringify({ heading: heads?.[b1]?.name, row: row1?.[0]?.name, place: b1 }));
let wrong = mismatches(tree);
check(tree.grids >= 1 && tree.rows.length > 5 && !wrong.length, `every cell of the ${tree.rows.length - 1} rows in view under its own column's heading`, wrong.slice(0, 5).join('; '));

// Scrolled right: the view starts past column A, so a cell's place no longer starts from A.
await page.evaluate(() => { const g = window.oxe.grid.el; g.scrollLeft = 400; });
await page.waitForTimeout(300);
tree = await treeRows();
wrong = mismatches(tree);
const firstHead = tree.rows.find((r) => r.some((c) => c.role === 'columnheader'))?.find((c) => c.role === 'columnheader' && c.name !== 'Row')?.name;
check(firstHead !== 'A' && !wrong.length, `scrolled right (first heading ${firstHead}): every cell under its own column's heading`, wrong.slice(0, 5).join('; '));
await page.evaluate(() => { window.oxe.grid.el.scrollLeft = 0; });

// A merged range (B10:D11): its covered columns still have a cell each, so E10 stays under E.
await page.evaluate(() => { window.oxe.select('B10'); window.oxe.command('extendTo', 10, 3); window.oxe.command('merge', 'cells'); window.oxe.select('E10'); });
await page.waitForTimeout(300);
wrong = mismatches(await treeRows());
check(!wrong.length, 'with B10:D11 merged: every cell under its own column\'s heading', wrong.slice(0, 5).join('; '));

// Panes frozen at C3: the frozen columns' headings drawn again for the eye only.
await page.evaluate(() => { window.oxe.select('C3'); window.oxe.command('freeze', 'panes'); });
await page.waitForTimeout(300);
tree = await treeRows();
wrong = mismatches(tree);
check(!wrong.length, 'with panes frozen at C3: every cell under its own column\'s heading', wrong.slice(0, 5).join('; '));

await finish();
