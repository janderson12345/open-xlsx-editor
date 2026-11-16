// Data > Queries & Connections: Excel's Queries & Connections pane, docked under the ribbon as the Watch Window is.
// Its Queries tab lists the workbook's Power Query queries (ooxml/queries.js), each with what Excel says under it
// ("4 rows loaded.", "Connection only.") and where it loads; its Connections tab the workbook's connections. Refresh
// and Refresh All get each query's rows from the host page (the onRefreshQuery option: a page can't reach a database
// or run the M engine) and write them into the query's table as Excel does, as one step of the history; without the
// option, Excel for the web's words that the connections can't be refreshed. Edit opens the query's M in an Advanced
// Editor (read-only when the workbook is), written back into the DataMashup on Done. docs/QUERIES.md.
import { el } from './ribbonkit.js';
import { modal, button } from './dialogs.js';
import { readQueries, queryStatus, QUERY_MESSAGES } from '../ooxml/queries.js';

const LOADS = { table: 'Table', model: 'Data model', connection: 'Connection only' };

/** What the host is given for query or connection `q`: { name, kind ('query', a Power Query; 'connection', another
 *  connection), connection ({ name, description, type, connection, command } or null), m (the query's M, or null),
 *  target ({ sheet, ref, table } of its table, or null) }. */
export function hostRequest(q, kind) {
  const c = kind === 'query' ? q.connection : q;
  return {
    name: q.name, kind,
    connection: c ? { name: c.name, description: c.description, type: c.typeName, connection: c.connection, command: c.command } : null,
    m: kind === 'query' ? q.m : null,
    target: q.table ? { sheet: q.table.sheet, ref: q.table.ref, table: q.table.name } : null,
  };
}

/** The connection names Excel lists when they couldn't be refreshed. */
const connNames = (items) => items.map(({ q, kind }) => (kind === 'query' ? q.connection?.name || `Query - ${q.name}` : q.name));

/**
 * Refresh queries and connections through the host: `items` [{ q (readQueries' query or connection), kind }]; `env`
 * { host() (the onRefreshQuery option), run(name, ...args) (a command), alert(text), announce(text), readOnly() }.
 * Resolves true when their rows were written.
 */
export async function refreshThrough(env, items) {
  if (!items.length) return false;
  if (env.readOnly()) { env.announce('The workbook is read-only'); return false; }
  const host = env.host();
  if (typeof host !== 'function') { env.alert(QUERY_MESSAGES.cant(connNames(items))); return false; }
  const results = [];
  for (const it of items) {
    env.announce(`Refreshing ${it.q.name}`);
    try {
      results.push({ name: it.q.name, kind: it.kind, data: await host(hostRequest(it.q, it.kind)) });
    } catch (e) {
      env.alert(`${QUERY_MESSAGES.cant(connNames([it]))}\n\n${e?.message || e}`);
      return false;
    }
  }
  return env.run('queryRows', results);
}

/** Excel's Advanced Editor over query `q`'s M: resolves to the new text on Done, null on Cancel (or when read-only,
 *  where it only shows it). */
export function advancedEditor(root, q, { readOnly = false } = {}) {
  return modal(root, (dlg, done, id) => {
    dlg.setAttribute('aria-labelledby', `${id}-title`);
    dlg.classList.add('oxe-query-editor');
    const text = el('textarea', { class: 'oxe-query-m', spellcheck: 'false', 'aria-label': `${q.name}'s formula`, rows: '16' });
    text.value = q.m;
    text.readOnly = readOnly;
    text.dataset.first = '';
    // (Tab types a tab, as in Excel's editor, the box's buttons then reached with Shift+Tab out of it... or the mouse.)
    text.addEventListener('keydown', (e) => {
      if (e.key !== 'Tab' || e.shiftKey || readOnly) return;
      e.preventDefault();
      text.setRangeText('    ', text.selectionStart, text.selectionEnd, 'end');
    });
    const buttons = readOnly ? [button('Close', () => done(null))] : [button('Done', () => done(text.value === q.m ? null : text.value)), button('Cancel', () => done(null))];
    dlg.append(el('h2', { id: `${id}-title`, class: 'oxe-dialog-title' }, q.name),
      el('p', { class: 'oxe-dialog-text' }, readOnly ? 'The workbook is read-only: the query\'s formula is shown, not changed.' : 'Advanced Editor'),
      text, el('div', { class: 'oxe-dialog-buttons' }, ...buttons));
  });
}

/** The pane. `ctx` = { root, workbook(), readOnly(), host(), run(name, ...args), alert(text), announce(text),
 *  go({ si, r, c }), home() }. */
export function createQueriesPane(ctx) {
  let tab = 'queries';
  const tabQ = el('button', { type: 'button', role: 'tab', class: 'oxe-query-tab' }, 'Queries');
  const tabC = el('button', { type: 'button', role: 'tab', class: 'oxe-query-tab' }, 'Connections');
  const all = el('button', { type: 'button', class: 'oxe-watch-btn' }, 'Refresh All');
  const close = el('button', { type: 'button', class: 'oxe-watch-btn oxe-find-close', 'aria-label': 'Close Queries & Connections', title: 'Close' }, '✕');
  const list = el('ul', { class: 'oxe-query-list', role: 'tabpanel' });
  const bar = el('section', { class: 'oxe-watchbar oxe-queriesbar', 'aria-label': 'Queries & Connections', hidden: '' },
    el('div', { class: 'oxe-watch-head' }, el('strong', {}, 'Queries & Connections'), el('div', { role: 'tablist', class: 'oxe-query-tabs' }, tabQ, tabC), all, close),
    el('div', { class: 'oxe-watch-scroll oxe-query-scroll' }, list));
  const env = { host: ctx.host, run: ctx.run, alert: ctx.alert, announce: ctx.announce, readOnly: ctx.readOnly };
  const refreshing = async (items) => { const done = await refreshThrough(env, items); refresh(); return done; };
  const edit = async (q) => {
    const m = await advancedEditor(ctx.root, q, { readOnly: ctx.readOnly() });
    if (m != null && ctx.run('queryFormula', q.name, m)) ctx.announce(`${q.name}'s formula changed`);
    refresh();
  };
  const small = (label, fn, disabled = false) => {
    const b = el('button', { type: 'button', class: 'oxe-watch-btn oxe-query-btn' }, label);
    b.disabled = disabled;
    b.addEventListener('click', fn);
    return b;
  };
  const item = (title, lines, buttons) => el('li', { class: 'oxe-query-item' },
    el('div', { class: 'oxe-query-name' }, title), ...lines.filter(Boolean).map((t) => el('div', { class: 'oxe-query-line' }, t)), el('div', { class: 'oxe-query-btns' }, ...buttons));
  const goTo = (t) => small('Go to', () => ctx.go({ si: t.si, r: t.r0, c: t.c0 }));
  function refresh() {
    if (bar.hidden) return;
    const wb = ctx.workbook();
    if (!wb) return;
    const { queries, connections } = readQueries(wb);
    const ro = ctx.readOnly();
    tabQ.textContent = `Queries (${queries.length})`;
    tabC.textContent = `Connections (${connections.length})`;
    tabQ.setAttribute('aria-selected', String(tab === 'queries'));
    tabC.setAttribute('aria-selected', String(tab === 'connections'));
    const tableQueries = queries.filter((q) => q.table);
    const tableConns = connections.filter((c) => c.table && !c.query);
    all.disabled = ro || !(tableQueries.length + tableConns.length);
    const rows = tab === 'queries'
      ? queries.map((q) => item(q.name, [queryStatus(q), `Load to: ${LOADS[q.load]}${q.table ? ` ${q.table.name} (${q.table.sheet}!${q.table.ref})` : ''}`],
        [small('Refresh', () => refreshing([{ q, kind: 'query' }]), ro || !q.table), small(ro ? 'View' : 'Edit', () => edit(q)), ...(q.table ? [goTo(q.table)] : [])]))
      : connections.map((c) => item(c.name, [c.description, `${c.typeName}${c.query ? `: the query ${c.query}` : ''}${c.table ? `, loaded to ${c.table.name} (${c.table.sheet}!${c.table.ref})` : c.model ? ', in the data model' : ''}`],
        [small('Refresh', () => (c.query ? refreshing(queries.filter((q) => q.name === c.query && q.table).map((q) => ({ q, kind: 'query' }))) : refreshing([{ q: c, kind: 'connection' }])), ro || !c.table), ...(c.table ? [goTo(c.table)] : [])]));
    list.replaceChildren(...(rows.length ? rows : [el('li', { class: 'oxe-query-none' }, tab === 'queries' ? 'This workbook has no queries.' : 'This workbook has no connections.')]));
  }
  tabQ.addEventListener('click', () => { tab = 'queries'; refresh(); });
  tabC.addEventListener('click', () => { tab = 'connections'; refresh(); });
  all.addEventListener('click', () => {
    const { queries, connections } = readQueries(ctx.workbook());
    refreshing([...queries.filter((q) => q.table).map((q) => ({ q, kind: 'query' })), ...connections.filter((c) => c.table && !c.query).map((c) => ({ q: c, kind: 'connection' }))]);
  });
  close.addEventListener('click', () => { bar.hidden = true; ctx.home(); });
  return {
    el: bar,
    /** Shown, or hidden again (Data > Queries & Connections turns it over). */
    toggle() { bar.hidden = !bar.hidden; refresh(); if (!bar.hidden) tabQ.focus(); else ctx.home(); },
    refresh,
    /** Refresh and Refresh All of the Data tab (Alt+F5, Ctrl+Alt+F5) for the queries and connections loaded to
     *  tables: every one (`all`), or the one whose table holds cell (r, c) of sheet si. False when there is none. */
    refreshTables(all, si, r, c) {
      const wb = ctx.workbook();
      const { queries, connections } = readQueries(wb);
      const inIt = (t) => t && t.si === si && r >= t.r0 && r <= t.r1 && c >= t.c0 && c <= t.c1;
      const items = [...queries.filter((q) => q.table && (all || inIt(q.table))).map((q) => ({ q, kind: 'query' })),
        ...connections.filter((x) => x.table && !x.query && (all || inIt(x.table))).map((x) => ({ q: x, kind: 'connection' }))];
      return items.length ? refreshing(items) : false;
    },
    get isOpen() { return !bar.hidden; },
  };
}
