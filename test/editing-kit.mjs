// The page the editing suites (editing.mjs, editing-data.mjs, editing-more.mjs) share: the demo opened in a browser,
// a checker, and the helpers each uses to read cells and click the grid and the ribbon.
import fs from 'node:fs';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { setXmlImpl } from '../src/ooxml/xml.js';
import { openDemo, checker } from './browser.mjs';

setXmlImpl({ DOMParser, XMLSerializer });

/** The demo page at 1100 by 760, and its helpers. */
export async function editingPage() {
  const outDir = process.argv[2] || 'test/out';
  fs.mkdirSync(outDir, { recursive: true });
  const { page, errors, close } = await openDemo({ width: 1100, height: 760 });
  const t = checker();
  const { check } = t;
  /** Open a workbook from a path through the file input, waiting until it is shown. */
  const openFixtureFrom = async (pg, file, name) => {
    await pg.setInputFiles('input[type=file]', file);
    await pg.waitForFunction((n) => window.oxe.fileName === n, name);
  };
  const shown = (r, c) => page.evaluate(([r, c]) => document.querySelector(`.oxe-row[aria-rowindex="${r + 2}"] .oxe-cell[aria-colindex="${c + 2}"]`)?.textContent ?? null, [r, c]);
  const sel = () => page.evaluate(() => window.oxe.selection);
  const value = (ref) => page.evaluate((ref) => {
    const { r, c } = { r: Number(ref.slice(1)) - 1, c: ref.charCodeAt(0) - 65 };
    const cell = window.oxe.workbook.sheet(window.oxe.sheetIndex).cells.get(r)?.get(c);
    return cell ? cell.v : null;
  }, ref);
  const live = () => page.evaluate(() => document.querySelector('.oxe-live')?.textContent || '');
  /** Where cell (r, c) is on the page, from the grid's geometry (empty cells have no element to click). */
  const pointAt = (r, c) => page.evaluate(([r, c]) => {
    const g = window.oxe.grid;
    const v = g.el.getBoundingClientRect();
    const { rows, cols, rowHeadW, colHeadH } = g.geometry();
    return { x: v.left + rowHeadW + cols.start(c) - g.el.scrollLeft + 6, y: v.top + colHeadH + rows.start(r) - g.el.scrollTop + 6 };
  }, [r, c]);
  const clickCell = async (r, c) => { const p = await pointAt(r, c); await page.mouse.click(p.x, p.y); };
  // A ribbon button, by its title: groups that don't fit at this width (it depends on the fonts) are behind the ribbon's
  // More button, opened for the click and closed after `then` (a menu choice, a box answered).
  const ribbon = async (title, then = async () => {}) => {
    const btn = `.oxe-toolbar button[title="${title}"]`;
    const more = !(await page.isVisible(btn)) && await page.isVisible('.oxe-more');
    if (more) await page.click('.oxe-more');
    await page.click(btn);
    await then();
    // (Not while the choice opened a box: the box holds the page until it's answered.)
    if (more && await page.isVisible('.oxe-more') && !(await page.isVisible('dialog[open]'))) await page.click('.oxe-more');
  };
  const menu = (title, item) => ribbon(title, () => page.click(`.oxe-pop >> text=${item}`));

  /** The suite's end: no page errors, the browser closed, the count of failures. */
  const finish = async () => {
    check(!errors.length, 'no page errors', errors.join(String.fromCharCode(10) + '     '));
    await close();
    console.log(t.failed ? `${t.failed} failed` : 'all passed');
    process.exit(t.failed ? 1 : 0);
  };
  return { page, errors, close, t, check, outDir, finish, openFixtureFrom, shown, sel, value, live, pointAt, clickCell, ribbon, menu };
}
