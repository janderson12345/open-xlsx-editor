// The Review tab, Excel's groups as far as they are here: Comments (New Comment, which opens the active cell's thread to
// reply to when it has one, Delete Comment: threadedit.js, threadbox.js) and Notes (New Note, which edits the active
// cell's note when it has one, Delete Note, Show/Hide Note: noteedit.js, notemarks.js).

/** The Review tab. `act(name, ...args)` runs a command. */
export function reviewTab(kit, { act }) {
  const { grp, item, menu, btn } = kit;
  const f = (name, ...args) => () => act(name, ...args);
  const groups = [
    grp('Comments', btn('💬', 'New Comment (Ctrl+Alt+M)', f('commentBox'), { label: 'New Comment', big: true }), btn('🗑', 'Delete Comment', f('deleteNote'), { label: 'Delete' })),
    grp('Notes', menu('🗒', 'Notes', () => [item('New Note', f('editNote'), { hint: 'Shift+F2' }), item('Edit Note', f('editNote')), item('Delete Note', f('deleteNote')),
      item('Show/Hide Note', f('showNote'))], { label: 'Notes', big: true })),
  ];
  return { groups, update() {} };
}
