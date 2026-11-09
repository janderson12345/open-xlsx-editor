// Accessibility: axe-core (WCAG 2.1 A and AA rules) on the demo with a workbook open, in light and dark, at a desktop
// width and at 320px (reflow: no page-level horizontal scrolling, the grid scrolls inside itself).
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
    t.check(!errors.length, `${colorScheme}, ${width}px: no page errors`, errors.join('\n     '));
    await close();
  }
}
console.log(t.failed ? `${t.failed} failed` : 'all passed');
process.exit(t.failed ? 1 : 0);
