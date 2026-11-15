// Formulas > Insert Function (Shift+F3, the fx button): Excel's box. "Search for a function" with Go, "Or select a
// category" (Most Recently Used, All and Excel's categories), the functions of the one chosen, and the syntax and
// description of the one picked under the list (the functions and their syntax: calc/catalog.js; the descriptions:
// calc/describe.js). OK or a double click resolves to the function's name; Cancel or Esc to null.
import { modal, button } from './dialogs.js';
import { CATEGORIES, catalog, inCategory, functionInfo, syntaxOf, searchFunctions } from '../calc/catalog.js';
import { describeFunction } from '../calc/describe.js';

const make = (tag, props = {}, ...kids) => { const e = Object.assign(document.createElement(tag), props); e.append(...kids); return e; };
const RECENT = 'Most Recently Used';
const ALL = 'All';

/** The functions of a category as the box lists them: the recent ones in their order, the others by name. */
export function functionsOf(category, recent) {
  if (category === RECENT) return recent.map(functionInfo).filter(Boolean);
  if (category === ALL) return catalog();
  return inCategory(category);
}

/** The Insert Function box: `recent` the Most Recently Used names, newest first. Resolves to a name or null. */
export function insertFunctionBox(root, { recent }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-function-box');
    const search = make('input', { id: `${id}-search`, className: 'oxe-dialog-field', type: 'text', spellcheck: false, value: 'Type a brief description of what you want to do and then click Go' });
    search.dataset.first = '';
    search.addEventListener('focus', () => search.select());
    const go = button('Go', () => find());
    const category = make('select', { id: `${id}-category`, className: 'oxe-dialog-select' },
      ...[RECENT, ALL, ...CATEGORIES].map((c) => make('option', { value: c, textContent: c })));
    const list = make('select', { id: `${id}-list`, className: 'oxe-dialog-list oxe-function-list', size: 9 });
    const syntax = make('p', { id: `${id}-syntax`, className: 'oxe-dialog-text oxe-function-syntax' });
    const about = make('p', { id: `${id}-about`, className: 'oxe-dialog-text oxe-function-about' });
    syntax.setAttribute('aria-live', 'polite');
    list.setAttribute('aria-describedby', `${syntax.id} ${about.id}`);
    const pick = (name) => { syntax.textContent = name ? syntaxOf(name) : 'No functions found.'; about.textContent = name ? describeFunction(name) : ''; };
    const show = (fns) => {
      list.replaceChildren(...fns.map((f) => make('option', { value: f.name, textContent: f.name })));
      if (fns.length) list.selectedIndex = 0;
      pick(fns.length ? list.value : null);
    };
    // (Go: the functions whose names hold the words typed; Excel's "Recommended" category for its results.)
    const find = () => {
      const found = searchFunctions(search.value);
      if (!found.length) { show([]); return; }
      let rec = category.querySelector('option[value="Recommended"]');
      if (!rec) category.prepend(rec = make('option', { value: 'Recommended', textContent: 'Recommended' }));
      category.value = 'Recommended';
      show(found);
    };
    category.addEventListener('change', () => { if (category.value !== 'Recommended') show(functionsOf(category.value, recent)); });
    list.addEventListener('change', () => pick(list.value));
    list.addEventListener('dblclick', () => { if (list.value) done(list.value); });
    const ok = () => done(list.value || null);
    dlg.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter' || e.target.tagName === 'BUTTON') return;
      e.preventDefault();
      if (e.target === search) find(); else ok();
    });
    show(functionsOf(RECENT, recent));
    const label = (forId, text) => make('label', { className: 'oxe-dialog-text oxe-dialog-label', htmlFor: forId, textContent: text });
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Insert Function' }),
      label(search.id, 'Search for a function:'), make('div', { className: 'oxe-function-search' }, search, go),
      label(category.id, 'Or select a category:'), category, label(list.id, 'Select a function:'), list, syntax, about,
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))));
  });
}
