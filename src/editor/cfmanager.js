// Home > Conditional Formatting > Manage Rules... and New Rule...: Excel's Conditional Formatting Rules Manager (the
// rules for the selection or the sheet, in priority order: each one's description, the cells it applies to and Stop If
// True; New Rule..., Delete Rule, Move Up, Move Down) and its New Formatting Rule box (Excel's six rule types). Each
// change is one step of the history (cfmanage.js, cfedit.js through acts-more.js).
import { modal, button } from './dialogs.js';
import { make, select, labelled, check } from './formatbox-kit.js';
import { ruleFormatBox } from './cfformatbox.js';
import { CF_PRESETS, CF_PERIODS, CF_SCALES } from '../ooxml/cfedit.js';
import { BAR_COLORS } from '../ooxml/cfbarsedit.js';
import { rangeName, parseRange } from '../ooxml/refs.js';

const OPS = [['between', 'between'], ['notBetween', 'not between'], ['equal', 'equal to'], ['notEqual', 'not equal to'], ['greaterThan', 'greater than'],
  ['lessThan', 'less than'], ['greaterThanOrEqual', 'greater than or equal to'], ['lessThanOrEqual', 'less than or equal to']];
const SIGN = { between: 'between', notBetween: 'not between', equal: '=', notEqual: '<>', greaterThan: '>', lessThan: '<', greaterThanOrEqual: '>=', lessThanOrEqual: '<=' };
const ICON_SETS = ['3Arrows', '3ArrowsGray', '3Triangles', '4Arrows', '4ArrowsGray', '5Arrows', '5ArrowsGray', '3TrafficLights1', '3TrafficLights2', '3Signs', '4RedToBlack',
  '4TrafficLights', '3Symbols', '3Symbols2', '3Flags', '3Stars', '4Rating', '5Quarters', '5Rating', '5Boxes'];

/** A rule as the Rules Manager names it (sheet.js's reading of it). */
export function describeRule(rule) {
  const f = (i) => (rule.formulas[i] != null ? `=${rule.formulas[i]}` : '');
  switch (rule.type) {
    case 'cellIs': return rule.operator === 'between' || rule.operator === 'notBetween' ? `Cell Value ${SIGN[rule.operator]} ${f(0)} and ${f(1)}` : `Cell Value ${SIGN[rule.operator] || rule.operator} ${f(0)}`;
    case 'expression': return `Formula: ${f(0)}`;
    case 'containsText': return `Cell value contains "${rule.text ?? ''}"`;
    case 'top10': return `${rule.bottom ? 'Bottom' : 'Top'} ${rule.rank}${rule.percent ? '%' : ''}`;
    case 'aboveAverage': return rule.aboveAverage ? 'Above Average' : 'Below Average';
    case 'timePeriod': return CF_PERIODS[rule.timePeriod]?.[0] || 'A Date Occurring';
    case 'duplicateValues': return 'Duplicate Values';
    case 'uniqueValues': return 'Unique Values';
    case 'colorScale': return 'Graded Color Scale';
    case 'dataBar': return 'Data Bar';
    case 'iconSet': return `Icon Set (${rule.iconSet?.set || ''})`;
    default: return rule.type.replace(/([A-Z])/g, ' $1').replace(/^./, (c) => c.toUpperCase());
  }
}

/** A rule's formula as the box's Value takes it: a number as it is, a quoted text without its quotes, else =formula. */
const typedOf = (f) => (f == null ? '' : /^-?[\d.]+(E[+-]?\d+)?$/i.test(f) ? f : /^"(.*)"$/s.test(f) ? f.slice(1, -1).replace(/""/g, '"') : `=${f}`);
/** The box's rule type for a rule read from the sheet. */
const TYPE_OF = { cellIs: 'contain', containsText: 'contain', timePeriod: 'contain', top10: 'rank', aboveAverage: 'average', duplicateValues: 'unique', uniqueValues: 'unique',
  expression: 'formula', colorScale: 'values', dataBar: 'values', iconSet: 'values' };

/** The New Formatting Rule box, or (`rule`, a rule read from the sheet) Edit Formatting Rule, set to that rule, its
 *  Format the rule's own unless another is picked. Resolves to a rule for cfedit.js (or cfbarsedit.js; `preset` null
 *  to keep the rule's format), or null. */
export function newRuleBox(root, { rule: old = null } = {}) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-cf-new');
    const type = select([['values', 'Format all cells based on their values'], ['contain', 'Format only cells that contain'], ['rank', 'Format only top or bottom ranked values'],
      ['average', 'Format only values that are above or below average'], ['unique', 'Format only unique or duplicate values'], ['formula', 'Use a formula to determine which cells to format']], old ? TYPE_OF[old.type] || 'formula' : 'values');
    type.size = 6;
    type.className = 'oxe-dialog-list';
    type.dataset.first = '';
    const panel = make('div', { className: 'oxe-cf-panel' });
    const preset = select([...(old ? [['keep', "(the rule's own format)"]] : []), ...Object.entries(CF_PRESETS).map(([v, p]) => [v, p.label]), ['custom', 'Custom Format...']], old ? 'keep' : 'lightRed');
    // (Custom Format...: Excel's Format Cells for the rule, cfformatbox.js; chosen again, the box opens again.)
    let customFormat = null;
    let lastPreset = preset.value;
    preset.addEventListener('change', async () => {
      if (preset.value !== 'custom') { lastPreset = preset.value; return; }
      const f = await ruleFormatBox(root);
      if (f) { customFormat = f; lastPreset = 'custom'; } else preset.value = lastPreset;
    });
    const fields = {};
    const field = (k, el) => { fields[k] = el; return el; };
    const input = (k, label) => labelled(label, field(k, make('input', { type: 'text', className: 'oxe-dialog-field', spellcheck: false })));
    const build = () => {
      const kids = [];
      if (type.value === 'values') {
        const style = field('style', select([['scale2', '2-Color Scale'], ['scale3', '3-Color Scale'], ['dataBar', 'Data Bar'], ['iconSet', 'Icon Set']], 'scale2'));
        const extra = make('div');
        const more = () => {
          const k = [];
          if (style.value === 'dataBar') k.push(labelled('Bar color:', field('barColor', select(Object.keys(BAR_COLORS).map((c) => [c, c.replace(/([A-Z])/g, ' $1').replace(/^./, (x) => x.toUpperCase())]), 'blue'))),
            labelled('Fill:', field('barFill', select([['gradient', 'Gradient Fill'], ['solid', 'Solid Fill']], 'gradient'))));
          if (style.value === 'iconSet') k.push(labelled('Icon Style:', field('icons', select(ICON_SETS.map((s) => [s, s]), '3TrafficLights1'))));
          // (A colour scale's Minimum, Midpoint and Maximum: a type, a value, a colour, as Excel's New Rule has them.)
          if (style.value === 'scale2' || style.value === 'scale3') {
            const three = style.value === 'scale3';
            const starts = three ? [['min', '', CF_SCALES[3][0]], ['percentile', '50', CF_SCALES[3][1]], ['max', '', CF_SCALES[3][2]]] : [['min', '', CF_SCALES[2][0]], ['max', '', CF_SCALES[2][1]]];
            fields.stops = starts.map(([type, val, color], i) => {
              const ends = i === 0 ? [['min', 'Lowest Value']] : i === starts.length - 1 ? [['max', 'Highest Value']] : [];
              const t = select([...ends, ['num', 'Number'], ['percent', 'Percent'], ['formula', 'Formula'], ['percentile', 'Percentile']], type);
              const v = make('input', { type: 'text', className: 'oxe-dialog-field', value: val, spellcheck: false });
              const c = make('input', { type: 'color', value: `#${color.slice(2)}` });
              const name = i === 0 ? 'Minimum' : i === starts.length - 1 ? 'Maximum' : 'Midpoint';
              v.setAttribute('aria-label', `${name} value`); c.setAttribute('aria-label', `${name} colour`);
              const sync = () => { v.disabled = t.value === 'min' || t.value === 'max'; };
              t.addEventListener('change', sync); sync();
              k.push(labelled(`${name}:`, t), v, c);
              return { t, v, c };
            });
          } else fields.stops = null;
          extra.replaceChildren(...k);
        };
        style.addEventListener('change', more);
        more();
        kids.push(labelled('Format Style:', style), extra);
      } else if (type.value === 'contain') {
        // (Cell Value and its comparison, Specific Text containing, or Dates Occurring and the period.)
        const what = field('what', select([['value', 'Cell Value'], ['text', 'Specific Text'], ['date', 'Dates Occurring']], 'value'));
        const rest = make('div');
        const syncWhat = () => {
          if (what.value === 'text') { rest.replaceChildren(input('text', 'containing')); return; }
          if (what.value === 'date') { rest.replaceChildren(labelled('Dates in the period', field('period', select(Object.entries(CF_PERIODS).map(([v, [l]]) => [v, l]), 'yesterday')))); return; }
          const op = field('op', select(OPS, 'between'));
          const second = make('div');
          const sync = () => second.replaceChildren(...(op.value === 'between' || op.value === 'notBetween' ? [input('v2', 'and')] : []));
          op.addEventListener('change', sync);
          rest.replaceChildren(labelled('Comparison', op), input('v1', 'Value'), second);
          sync();
        };
        what.addEventListener('change', syncWhat);
        kids.push(labelled('Format only cells with', what), rest);
        syncWhat();
      } else if (type.value === 'rank') {
        kids.push(labelled('Format values that rank in the', field('which', select([['top', 'Top'], ['bottom', 'Bottom']], 'top'))), input('rank', 'Number'));
        fields.rank.value = '10';
        const pc = check('% of the selected range', false);
        fields.percent = pc.input;
        kids.push(pc.el);
      } else if (type.value === 'average') kids.push(labelled('Format values that are', field('avg', select([['above', 'above'], ['below', 'below']], 'above'))));
      else if (type.value === 'unique') kids.push(labelled('Format all', field('dup', select([['duplicate', 'duplicate'], ['unique', 'unique']], 'duplicate'))));
      else kids.push(input('formula', 'Format values where this formula is true:'));
      if (type.value !== 'values') kids.push(labelled('Format:', preset));
      panel.replaceChildren(...kids);
    };
    type.addEventListener('change', build);
    build();
    // (Edit Rule: the box set to the rule's settings.)
    const put = (k, v) => { if (fields[k] && v != null) { fields[k].value = String(v); fields[k].dispatchEvent(new Event('change')); } };
    if (old) {
      if (old.type === 'colorScale') put('style', old.colorScale?.cfvo.length === 2 ? 'scale2' : 'scale3');
      if (old.type === 'dataBar') put('style', 'dataBar');
      if (old.type === 'iconSet') { put('style', 'iconSet'); put('icons', old.iconSet?.set); }
      if (old.type === 'containsText') { put('what', 'text'); put('text', old.text); }
      if (old.type === 'timePeriod') { put('what', 'date'); put('period', old.timePeriod); }
      if (old.type === 'cellIs') { put('op', old.operator); put('v1', typedOf(old.formulas[0])); put('v2', typedOf(old.formulas[1])); }
      if (old.type === 'top10') { put('which', old.bottom ? 'bottom' : 'top'); put('rank', old.rank); fields.percent.checked = !!old.percent; }
      if (old.type === 'aboveAverage') put('avg', old.aboveAverage ? 'above' : 'below');
      if (old.type === 'uniqueValues') put('dup', 'unique');
      if (old.type === 'expression') put('formula', typedOf(old.formulas[0]).replace(/^(?!=)/, '='));
    }
    const format = () => (preset.value === 'keep' ? null : preset.value === 'custom' ? customFormat || 'lightRed' : preset.value);
    const why = make('p', { className: 'oxe-dialog-text', role: 'alert' });
    const ok = () => {
      const v = (k) => fields[k]?.value ?? '';
      let rule = null;
      switch (type.value) {
        case 'values': {
          const s = v('style');
          const stops = fields.stops?.map(({ t, v: val, c }) => ({ type: t.value, val: val.value.trim(), color: `FF${c.value.slice(1).toUpperCase()}` }));
          if (stops?.some((x) => x.type !== 'min' && x.type !== 'max' && !x.val)) { why.textContent = 'Enter a value.'; return; }
          rule = s === 'dataBar' ? { kind: 'dataBar', color: BAR_COLORS[v('barColor')], solid: v('barFill') === 'solid' } : s === 'iconSet' ? { kind: 'iconSet', set: v('icons') }
            : { kind: 'colorScale', points: s === 'scale2' ? 2 : 3, stops };
          break;
        }
        case 'contain':
          if (v('what') === 'text') { rule = { kind: 'containsText', text: v('text'), preset: format() }; break; }
          if (v('what') === 'date') { rule = { kind: 'timePeriod', period: v('period'), preset: format() }; break; }
          if (!v('v1').trim() || ((v('op') === 'between' || v('op') === 'notBetween') && !v('v2').trim())) { why.textContent = 'Enter a value.'; return; }
          rule = { kind: 'cellIs', operator: v('op'), f1: v('v1'), f2: v('v2'), preset: format() };
          break;
        case 'rank': rule = { kind: 'top10', rank: Math.max(1, Math.round(Number(v('rank')) || 10)), percent: fields.percent.checked, bottom: v('which') === 'bottom', preset: format() }; break;
        case 'average': rule = { kind: 'average', below: v('avg') === 'below', preset: format() }; break;
        case 'unique': rule = { kind: 'duplicate', unique: v('dup') === 'unique', preset: format() }; break;
        default:
          if (!v('formula').trim()) { why.textContent = 'Enter a formula.'; return; }
          rule = { kind: 'expression', formula: v('formula'), preset: format() };
      }
      done(rule);
    };
    dlg.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target.tagName === 'INPUT') { e.preventDefault(); ok(); } });
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: old ? 'Edit Formatting Rule' : 'New Formatting Rule' }), labelled('Select a Rule Type:', type), panel, why,
      make('div', { className: 'oxe-dialog-buttons' }, button('OK', ok), button('Cancel', () => done(null))));
  });
}

/** The Rules Manager. `rules()` gives the sheet's rules ([{ rule, ranges }]), `selection` the selected cells;
 *  `run(change)` makes a change (cfmanage.js's, by priority) as one step and returns Excel's words when refused;
 *  `newRule()` opens New Rule for the selection. Edit Rule (or a double-click on the rule) opens Edit Formatting Rule.
 *  Resolves when closed. */
export function rulesManager(root, { rules, selection, run, newRule }) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-cf-manager');
    const scope = select([['selection', 'Current Selection'], ['sheet', 'This Worksheet']], 'selection');
    const body = make('tbody');
    const table = make('table', { className: 'oxe-pivot-table oxe-names-table' },
      make('thead', {}, make('tr', {}, ...['Rule (applied in order shown)', 'Applies to', 'Stop If True'].map((h) => make('th', { scope: 'col', textContent: h })))), body);
    const why = make('p', { className: 'oxe-dialog-text', role: 'alert' });
    let chosen = null;
    const meets = (a, b) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;
    const say = (refused) => { why.textContent = refused || ''; };
    const fill = () => {
      const list = rules().filter((x) => scope.value === 'sheet' || x.ranges.some((r) => selection.some((s) => meets(r, s)))).sort((a, b) => a.rule.priority - b.rule.priority);
      if (chosen != null && !list.some((x) => x.rule.priority === chosen)) chosen = list[0]?.rule.priority ?? null;
      if (chosen == null) chosen = list[0]?.rule.priority ?? null;
      body.replaceChildren(...list.map((x, k) => {
        const p = x.rule.priority;
        const applies = make('input', { type: 'text', className: 'oxe-dialog-field', value: `=${x.ranges.map(rangeName).join(',')}`, spellcheck: false });
        applies.setAttribute('aria-label', `Applies to, ${describeRule(x.rule)}`);
        applies.addEventListener('change', () => {
          const ranges = applies.value.replace(/^=/, '').split(/[,\s]+/).map((t) => parseRange(t.replace(/\$/g, ''))).filter(Boolean);
          if (!ranges.length) { say('Enter the cells as a reference, like =A1:B10.'); return; }
          say(run({ ranges: { p, ranges } }));
          fill();
        });
        const stop = make('input', { type: 'checkbox', checked: !!x.rule.stopIfTrue, disabled: ['dataBar', 'iconSet', 'colorScale'].includes(x.rule.type) });
        stop.setAttribute('aria-label', `Stop If True, ${describeRule(x.rule)}`);
        stop.addEventListener('change', () => { say(run({ stop: { p, on: stop.checked } })); fill(); });
        const row = make('tr', { tabIndex: k === 0 ? 0 : -1, className: 'oxe-names-row' }, make('td', { textContent: describeRule(x.rule) }), make('td', {}, applies), make('td', {}, stop));
        row.setAttribute('aria-selected', String(p === chosen));
        row.addEventListener('click', () => { chosen = p; for (const r of body.children) r.setAttribute('aria-selected', String(r === row)); });
        row.addEventListener('focus', () => { chosen = p; for (const r of body.children) r.setAttribute('aria-selected', String(r === row)); });
        row.addEventListener('dblclick', (e) => { if (e.target === row || e.target.tagName === 'TD') editChosen(); });
        row.addEventListener('keydown', (e) => { if (e.target === row && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) { e.preventDefault(); (e.key === 'ArrowDown' ? row.nextSibling : row.previousSibling)?.focus(); } });
        return row;
      }));
      table.setAttribute('aria-label', list.length ? `${list.length} rules` : 'No rules');
    };
    const order = () => rules().map((x) => x.rule.priority).sort((a, b) => a - b);
    const up = () => { const ps = order(); const k = ps.indexOf(chosen); if (k > 0) { say(run({ move: { p: chosen, to: ps[k - 1] } })); chosen = ps[k - 1]; fill(); } };
    // (Down: the next rule moved above this one, as Excel's priorities go.)
    const down = () => { const ps = order(); const k = ps.indexOf(chosen); if (k >= 0 && k < ps.length - 1) { const next = ps[k + 1]; say(run({ move: { p: next, to: chosen } })); chosen += 1; fill(); } };
    scope.addEventListener('change', fill);
    async function editChosen() {
      const x = rules().find((y) => y.rule.priority === chosen);
      if (!x) return;
      const rule = await newRuleBox(root, { rule: x.rule });
      if (rule) { say(run({ edit: { p: chosen, rule } })); fill(); }
    }
    const top = make('div', { className: 'oxe-dialog-buttons oxe-names-actions' },
      button('New Rule...', async () => { await newRule(); fill(); }),
      button('Edit Rule...', editChosen),
      button('Delete Rule', () => { if (chosen != null) { say(run({ remove: chosen })); chosen = null; fill(); } }),
      button('Move Up', up), button('Move Down', down));
    fill();
    scope.dataset.first = '';
    dlg.append(make('h2', { id: `${id}-title`, className: 'oxe-dialog-title', textContent: 'Conditional Formatting Rules Manager' }), labelled('Show formatting rules for:', scope), top,
      make('div', { className: 'oxe-names-list' }, table), why, make('div', { className: 'oxe-dialog-buttons' }, button('Close', () => done(true))));
  });
}
