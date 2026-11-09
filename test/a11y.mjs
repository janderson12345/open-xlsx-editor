// Accessibility: axe-core (WCAG 2.1 A and AA rules) on the demo with a workbook open, in light and dark, at a desktop
// width and at 320px (reflow: no page-level horizontal scrolling, the grid scrolls inside itself); the grid and the
// sheet tabs as ARIA describes them.
// Run after `npm run build`: node test/a11y.mjs
import { createRequire } from 'node:module';
import { openDemo, openFixture, checker } from './browser.mjs';

const axePath = createRequire(import.meta.url).resolve('axe-core/axe.min.js');
const t = checker();

for (const colorScheme of ['light', 'dark']) {
  for (const [width, height] of [[1280, 900], [320, 640]]) {
    const { page, errors, close } = await openDemo({ width, height, colorScheme });
    await openFixture(page, 'basic.xlsx');
    await page.addScriptTag({ path: axePath });
    const result = await page.evaluate(() => window.axe.run(document, { runOnly: { type: 'tag', values: ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa'] } }));
    const issues = result.violations.map((v) => `${v.id} (${v.nodes.length}): ${v.help} - ${v.nodes.slice(0, 2).map((n) => n.target.join(' ')).join(', ')}`);
    const passed = result.passes.map((p) => p.id);
    t.check(!issues.length && passed.includes('color-contrast') && passed.includes('aria-required-children'), `${colorScheme}, ${width}px: no axe violations (${passed.length} rules passed, contrast and grid structure among them)`, issues.join('\n     '));
    const reflow = await page.evaluate(() => ({ page: document.documentElement.scrollWidth, view: window.innerWidth }));
    t.check(reflow.page <= reflow.view, `${colorScheme}, ${width}px: no horizontal page scroll`, JSON.stringify(reflow));
    const grid = await page.evaluate(() => {
      const g = document.querySelector('.oxe-grid');
      return { role: g.getAttribute('role'), focusable: g.tabIndex >= 0, rows: g.getAttribute('aria-rowcount'), label: g.getAttribute('aria-label') };
    });
    t.check(grid.role === 'grid' && grid.focusable && Number(grid.rows) > 0 && grid.label === 'Data', `${colorScheme}, ${width}px: the grid is a labelled, focusable ARIA grid`, JSON.stringify(grid));
    const tabs = await page.evaluate(() => {
      const list = document.querySelector('.oxe-tabs');
      const id = document.querySelector('.oxe-grid').id;
      return { role: list.getAttribute('role'), label: list.getAttribute('aria-label'), tabs: [...list.children].map((el) => `${el.getAttribute('role')}:${el.getAttribute('aria-controls') === id}:${el.getAttribute('aria-selected')}:${el.tabIndex}`) };
    });
    t.check(tabs.role === 'tablist' && tabs.label === 'Sheets' && tabs.tabs.join() === 'tab:true:true:0,tab:true:false:-1', `${colorScheme}, ${width}px: the sheet tabs are a labelled tablist controlling the grid, one tab in the tab order`, JSON.stringify(tabs));
    if (width > 320) {
      // The tab menu with its colours open, then the Unhide box: axe again on each.
      const axe = async () => (await page.evaluate(() => window.axe.run(document, { runOnly: { type: 'tag', values: ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa'] } }))).violations.map((v) => `${v.id}: ${v.nodes.slice(0, 2).map((n) => n.target.join(' ')).join(', ')}`);
      await page.click('.oxe-tab >> text=Data', { button: 'right' });
      await page.click('.oxe-menuitem >> text=Tab Color');
      const menu = await axe();
      t.check(!menu.length, `${colorScheme}: the tab menu and its colours, no axe violations`, menu.join('\n     '));
      await page.keyboard.press('Escape');
      await page.keyboard.press('Escape');
      await page.click('.oxe-tab >> text=Data', { button: 'right' });
      await page.click('.oxe-menuitem >> text=Unhide...');
      const box = await axe();
      t.check(!box.length && await page.evaluate(() => document.activeElement.tagName) === 'SELECT', `${colorScheme}: the Unhide box, no axe violations, the list focused`, box.join('\n     '));
      await page.keyboard.press('Escape');
      // The ribbon's colour palette and Borders menu open: axe on each; the ribbon's tabs a labelled tablist.
      await page.click('.oxe-split-more[aria-label="Fill Color options"]');
      const palette = await axe();
      t.check(!palette.length && await page.isVisible('.oxe-pop .oxe-swatches-theme'), `${colorScheme}: the Fill Color palette, no axe violations`, palette.join('\n     '));
      await page.keyboard.press('Escape');
      await page.click('.oxe-split-more[aria-label="Bottom Border options"]');
      const borders = await axe();
      t.check(!borders.length, `${colorScheme}: the Borders menu, no axe violations`, borders.join('\n     '));
      await page.keyboard.press('Escape');
      await page.click('.oxe-toolbar button[title="Format"]');
      await page.click('.oxe-pop >> text=Row Height...');
      const sizeBox = await axe();
      t.check(!sizeBox.length && await page.evaluate(() => document.activeElement.classList.contains('oxe-dialog-field')), `${colorScheme}: the Row Height box, no axe violations, its field focused`, sizeBox.join('\n     '));
      await page.keyboard.press('Escape');
      const ribbon = await page.evaluate(() => {
        const list = document.querySelector('.oxe-rtablist');
        const tab = list.querySelector('[role=tab]');
        return `${list.getAttribute('role')}:${list.getAttribute('aria-label')}:${tab.getAttribute('aria-selected')}:${document.getElementById(tab.getAttribute('aria-controls'))?.getAttribute('role')}`;
      });
      t.check(ribbon === 'tablist:Ribbon tabs:true:tabpanel', `${colorScheme}: the ribbon's tabs are a labelled tablist, each controlling its panel`, ribbon);
    }
    t.check(!errors.length, `${colorScheme}, ${width}px: no page errors`, errors.join('\n     '));
    await close();
  }
}
console.log(t.failed ? `${t.failed} failed` : 'all passed');
process.exit(t.failed ? 1 : 0);
