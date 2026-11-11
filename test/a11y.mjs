// Accessibility: axe-core (WCAG 2.1 A and AA rules) on the demo with a workbook open, in light and dark, at a desktop
// width and at 320px (reflow: no page-level horizontal scrolling, the grid scrolls inside itself); the grid and the
// sheet tabs as ARIA describes them.
// Run after `npm run build`: node test/a11y.mjs
import { createRequire } from 'node:module';
import fs from 'node:fs';
import { openDemo, openFixture, checker } from './browser.mjs';
import { encryptWorkbook } from '../src/ooxml/encryption.js';

const axePath = createRequire(import.meta.url).resolve('axe-core/axe.min.js');
const t = checker();
// (A workbook with a test password, for the Password box.)
fs.mkdirSync('test/out', { recursive: true });
fs.writeFileSync('test/out/a11y-encrypted.xlsx', encryptWorkbook(new Uint8Array(fs.readFileSync('test/fixtures/basic.xlsx')), 'Test-pass 1', null, { spins: 1000 }));

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
      // (Format is in the Cells group, which may be behind the ribbon's More button at this width.)
      const more = !(await page.isVisible('.oxe-toolbar button[title="Format"]')) && await page.isVisible('.oxe-more');
      if (more) await page.click('.oxe-more');
      await page.click('.oxe-toolbar button[title="Format"]');
      await page.click('.oxe-pop >> text=Row Height...');
      const sizeBox = await axe();
      t.check(!sizeBox.length && await page.evaluate(() => document.activeElement.classList.contains('oxe-dialog-field')), `${colorScheme}: the Row Height box, no axe violations, its field focused`, sizeBox.join('\n     '));
      await page.keyboard.press('Escape');
      if (more) await page.click('.oxe-more');
      await page.focus('.oxe-grid');
      await page.keyboard.press('Control+Shift+Equal');
      await page.waitForSelector('dialog.oxe-dialog');
      const insertBox = await axe();
      t.check(!insertBox.length && await page.evaluate(() => document.activeElement.type === 'radio' && document.activeElement.checked), `${colorScheme}: the Insert box, no axe violations, its chosen option focused`, insertBox.join('\n     '));
      await page.keyboard.press('Escape');
      await page.focus('.oxe-grid');
      await page.keyboard.press('Control+c');
      await page.keyboard.press('Control+Alt+v');
      await page.waitForSelector('dialog.oxe-dialog');
      const specialBox = await axe();
      t.check(!specialBox.length && await page.evaluate(() => document.activeElement.type === 'radio' && document.activeElement.checked), `${colorScheme}: the Paste Special box, no axe violations, All focused`, specialBox.join('\n     '));
      await page.keyboard.press('Escape');
      // Sort & Filter: on E2:E3 Sort A to Z asks with Excel's Sort Warning; on one cell Custom Sort... opens the Sort box.
      const sortMenu = async (item) => {
        const hidden = !(await page.isVisible('.oxe-toolbar button[title="Sort & Filter"]')) && await page.isVisible('.oxe-more');
        if (hidden) await page.click('.oxe-more');
        await page.click('.oxe-toolbar button[title="Sort & Filter"]');
        await page.click(`.oxe-pop >> text=${item}`);
        return hidden;
      };
      await page.evaluate(() => window.oxe.select('E2'));
      await page.focus('.oxe-grid');
      await page.keyboard.press('Shift+ArrowDown');
      let opened = await sortMenu('Sort A to Z');
      await page.waitForSelector('dialog.oxe-dialog');
      const warningBox = await axe();
      t.check(!warningBox.length && await page.evaluate(() => document.activeElement.type === 'radio' && document.activeElement.checked), `${colorScheme}: the Sort Warning, no axe violations, Expand the selection focused`, warningBox.join('\n     '));
      await page.keyboard.press('Escape');
      if (opened) await page.click('.oxe-more');
      await page.evaluate(() => window.oxe.select('B3'));
      opened = await sortMenu('Custom Sort...');
      await page.waitForSelector('dialog.oxe-sort-box');
      const sortDialog = await axe();
      t.check(!sortDialog.length && await page.evaluate(() => document.activeElement.classList.contains('oxe-sort-select')), `${colorScheme}: the Sort box, no axe violations, its first column list focused`, sortDialog.join('\n     '));
      await page.keyboard.press('Escape');
      if (opened) await page.click('.oxe-more');
      // The filter: Ctrl+Shift+L on B3, Alt+Down on B1: the column's list; axe on it; then the filter off again.
      await page.evaluate(() => window.oxe.select('B3'));
      await page.focus('.oxe-grid');
      await page.keyboard.press('Control+Shift+KeyL');
      await page.evaluate(() => window.oxe.select('B1'));
      await page.focus('.oxe-grid');
      await page.keyboard.press('Alt+ArrowDown');
      await page.waitForSelector('dialog.oxe-filter-box');
      const filterList = await axe();
      t.check(!filterList.length && await page.evaluate(() => document.activeElement.type === 'search'), `${colorScheme}: a filter's list, no axe violations, its search box focused`, filterList.join('\n     '));
      await page.keyboard.press('Escape');
      await page.focus('.oxe-grid');
      await page.keyboard.press('Control+Shift+KeyL');
      await page.setInputFiles('input[type=file]', 'test/out/a11y-encrypted.xlsx');
      await page.waitForSelector('dialog.oxe-dialog input[type=password]');
      const passwordDialog = await axe();
      t.check(!passwordDialog.length && await page.evaluate(() => document.activeElement.type === 'password'), `${colorScheme}: the Password box, no axe violations, its field focused`, passwordDialog.join('\n     '));
      await page.keyboard.press('Escape');
      await page.focus('.oxe-grid');
      await page.evaluate(() => window.oxe.select('C7'));
      await page.focus('.oxe-grid');
      await page.keyboard.press('Shift+F2');
      await page.waitForSelector('dialog.oxe-note-edit');
      const noteEdit = await axe();
      t.check(!noteEdit.length && await page.evaluate(() => document.activeElement.tagName === 'TEXTAREA'), `${colorScheme}: the New Note box, no axe violations, its text focused`, noteEdit.join(String.fromCharCode(10) + '     '));
      await page.click('dialog.oxe-note-edit >> text=Cancel');
      await page.focus('.oxe-grid');
      await page.keyboard.press('F5');
      await page.waitForSelector('dialog.oxe-dialog');
      const goTo = await axe();
      t.check(!goTo.length && await page.evaluate(() => document.activeElement.classList.contains('oxe-dialog-field')), `${colorScheme}: the Go To box, no axe violations, its Reference field focused`, goTo.join('\n     '));
      await page.keyboard.press('Escape');
      await page.focus('.oxe-grid');
      const cell = await page.evaluate(() => { const r = document.querySelector('.oxe-row[aria-rowindex="3"] .oxe-cell[aria-colindex="3"]').getBoundingClientRect(); return { x: r.x + r.width / 2, y: r.y + r.height / 2 }; });
      await page.mouse.click(cell.x, cell.y, { button: 'right' });
      await page.waitForSelector('.oxe-menu >> text=Paste Special...');
      const cellMenu = await axe();
      t.check(!cellMenu.length, `${colorScheme}: a cell's menu, no axe violations`, cellMenu.join('\n     '));
      await page.keyboard.press('Escape');
      await page.keyboard.press('Escape');
      await page.focus('.oxe-grid');
      await page.keyboard.press('Control+h');
      const findBar = await axe();
      t.check(!findBar.length && await page.evaluate(() => document.activeElement.getAttribute('aria-label') === 'Find what'), `${colorScheme}: the Find and Replace bar, no axe violations, its box focused`, findBar.join('\n     '));
      await page.keyboard.press('Escape');
      await page.click('.oxe-rtab >> text=Data');
      await page.click('.oxe-toolbar button.oxe-split-more[aria-label^="Refresh All"]');
      const dataTab = await axe();
      t.check(!dataTab.length, `${colorScheme}: the Data tab with its Refresh All menu, no axe violations`, dataTab.join('\n     '));
      await page.keyboard.press('Escape');
      await page.click('.oxe-rtab >> text=Home');
      const cfHidden = !(await page.isVisible('.oxe-toolbar button[title="Conditional Formatting"]')) && await page.isVisible('.oxe-more');
      if (cfHidden) await page.click('.oxe-more');
      await page.click('.oxe-toolbar button[title="Conditional Formatting"]');
      const cfMenu = await axe();
      await page.click('.oxe-pop >> text=Between...');
      await page.waitForSelector('dialog.oxe-cf-box');
      const cfBox = await axe();
      t.check(!cfMenu.length && !cfBox.length && await page.evaluate(() => document.activeElement.tagName === 'INPUT'), `${colorScheme}: the Conditional Formatting menu and its Between box, no axe violations, From focused`, [...cfMenu, ...cfBox].join(String.fromCharCode(10) + '     '));
      await page.keyboard.press('Escape');
      if (cfHidden && await page.isVisible('.oxe-more')) await page.click('.oxe-more');
      await page.click('.oxe-rtab >> text=Data');
      await page.click('.oxe-toolbar button[title="Data Validation..."]');
      await page.waitForSelector('dialog.oxe-dv-box');
      const dvBox = await axe();
      t.check(!dvBox.length && await page.evaluate(() => document.activeElement.tagName === 'SELECT'), `${colorScheme}: the Data Validation box, no axe violations, Allow focused`, dvBox.join(String.fromCharCode(10) + '     '));
      await page.keyboard.press('Escape');
      // (Go To Special from the Go To box's Special... button, as F5 then Special... reaches it.)
      await page.focus('.oxe-grid');
      await page.keyboard.press('F5');
      await page.click('dialog.oxe-dialog >> button:text-is("Special...")');
      await page.waitForSelector('dialog.oxe-special-box');
      const gotoSpecialAxe = await axe();
      t.check(!gotoSpecialAxe.length && await page.evaluate(() => document.activeElement.type === 'radio'), `${colorScheme}: the Go To Special box, no axe violations, a choice focused`, gotoSpecialAxe.join(String.fromCharCode(10) + '     '));
      await page.keyboard.press('Escape');
      // (Insert > PivotTable: its source box, then its fields box, from A1:B3 typed in.)
      await page.click('.oxe-rtab >> text=Insert');
      await page.click('.oxe-toolbar button[title="PivotTable"]');
      await page.waitForSelector('dialog.oxe-pivot-box');
      const pivotSourceAxe = await axe();
      await page.fill('dialog.oxe-pivot-box input[type="text"]', 'A1:B3');
      await page.click('dialog.oxe-pivot-box >> button:text-is("OK")');
      await page.waitForSelector('dialog.oxe-pivot-fields');
      const pivotFieldsAxe = await axe();
      t.check(!pivotSourceAxe.length && !pivotFieldsAxe.length, `${colorScheme}: Insert > PivotTable's two boxes, no axe violations`, [...pivotSourceAxe, ...pivotFieldsAxe].join(String.fromCharCode(10) + '     '));
      await page.keyboard.press('Escape');
      // (Formulas > Name Manager, and its New... box.)
      await page.click('.oxe-rtab >> text=Formulas');
      await page.click('.oxe-toolbar button[title="Name Manager (Ctrl+F3)"]');
      await page.waitForSelector('dialog.oxe-name-manager');
      const managerAxe = await axe();
      await page.click('dialog.oxe-name-manager >> button:text-is("New...")');
      await page.waitForSelector('dialog.oxe-name-box');
      const newNameAxe = await axe();
      t.check(!managerAxe.length && !newNameAxe.length, `${colorScheme}: the Name Manager and its New Name box, no axe violations`, [...managerAxe, ...newNameAxe].join(String.fromCharCode(10) + '     '));
      await page.keyboard.press('Escape');
      await page.keyboard.press('Escape');
      await page.click('.oxe-rtab >> text="View"');
      await page.click('.oxe-toolbar button[title="Freeze Panes"]');
      const viewTab = await axe();
      t.check(!viewTab.length, `${colorScheme}: the View tab with its Freeze Panes menu, no axe violations`, viewTab.join('\n     '));
      await page.keyboard.press('Escape');
      await page.click('.oxe-rtab >> text=Home');
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
