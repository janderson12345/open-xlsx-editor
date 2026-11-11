// Review of an agent's edits (agent.js changes() and revert()). Excel has no tracked changes for cells, so the review is
// the workbook now beside the workbook as it was opened: the sheets added, deleted or renamed (sheets are matched by
// their ids, so a rename is one change), the rows and columns inserted or deleted (from the agent's own log), and each
// cell whose content changed (`edited`: what was entered, a formula or a value) or whose formula's result changed with
// it (`result`: worked out again, nothing entered there; only where the file came with a result to compare with).

/** The changes from workbook `before` to workbook `after` (each an AgentWorkbook), with the agent's `log` of rows,
 *  columns and sheets: { sheets, structure, cells: [{ ref, kind, before, after }], more } (cells past `limit` are
 *  counted in `more`). */
export function workbookChanges(before, after, log = [], { limit = 500 } = {}) {
  const sheets = [];
  const cells = [];
  let more = 0;
  const byId = (wb) => new Map(wb.wb.sheets.map((s, si) => [s.sheetId, { s, si }]));
  const was = byId(before);
  const now = byId(after);
  for (const [id, { s }] of was) if (!now.has(id)) sheets.push({ change: 'deleted', name: s.name });
  for (const [id, { s, si }] of now) {
    const old = was.get(id);
    if (!old) { sheets.push({ change: 'added', name: s.name }); continue; }
    if (old.s.name !== s.name) sheets.push({ change: 'renamed', from: old.s.name, name: s.name });
    if (s.kind !== 'worksheet' || old.s.kind !== 'worksheet') continue;
    const a = before.wb.sheet(old.si);
    const b = after.wb.sheet(si);
    const keys = new Set();
    for (const sh of [a, b]) for (const [r, line] of sh.cells) for (const c of line.keys()) keys.add(r * 16384 + c);
    for (const k of [...keys].sort((x, y) => x - y)) {
      const r = Math.floor(k / 16384);
      const c = k % 16384;
      const x = before.cell(old.si, a, r, c);
      const y = after.cell(si, b, r, c);
      const entered = (x.formula ?? x.value) !== (y.formula ?? y.value) || (x.formula == null) !== (y.formula == null);
      // (A result the file came without, as some programs save formulas, isn't a change: nothing was known before.)
      const kind = entered ? 'edited' : x.value !== y.value && x.value != null ? 'result' : null;
      if (!kind) continue;
      if (cells.length >= limit) { more++; continue; }
      cells.push({ ref: after.qualified(si, { r0: r, c0: c, r1: r, c1: c }), kind, before: { value: x.value, formula: x.formula }, after: { value: y.value, formula: y.formula } });
    }
  }
  return { sheets, structure: log.map((e) => ({ ...e })), cells, more };
}

/** The changes as a Markdown summary for a person (or an assistant) to read. */
export function changesMarkdown(ch) {
  const show = (v) => (v.formula ?? (v.value == null ? '(empty)' : String(v.value)));
  const lines = [];
  for (const s of ch.sheets) lines.push(`- Sheet ${s.change === 'renamed' ? `${s.from} renamed ${s.name}` : `${s.name} ${s.change}`}`);
  for (const e of ch.structure) lines.push(`- ${e.text}`);
  const edited = ch.cells.filter((x) => x.kind === 'edited');
  const results = ch.cells.filter((x) => x.kind === 'result');
  if (edited.length) {
    lines.push('', '| Cell | Before | After |', '|---|---|---|');
    for (const x of edited) lines.push(`| ${x.ref} | ${show(x.before)} | ${show(x.after)} |`);
  }
  if (results.length) lines.push('', `Results worked out again: ${results.map((x) => `${x.ref} ${x.before.value ?? '(empty)'} → ${x.after.value ?? '(empty)'}`).join('; ')}`);
  if (ch.more) lines.push('', `…and ${ch.more} more cells.`);
  return lines.join('\n').trim() || 'No changes.';
}
