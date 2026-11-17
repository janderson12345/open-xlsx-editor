// The Review tab, Excel's groups as far as they are here: Comments (New Comment, which opens the active cell's thread to
// reply to when it has one, Delete Comment: threadedit.js, threadbox.js) and Notes (New Note, which edits the active
// cell's note when it has one, Delete Note, Show/Hide Note, Show All Notes, Convert to Comments: noteedit.js,
// noteconvert.js, notemarks.js) and Protect (Protect Sheet, or Unprotect Sheet once it is; Protect Workbook:
// protectbox.js, ooxml/protection.js; Allow Edit Ranges: rangesbox.js, ooxml/editranges.js).

/** The Review tab. `act(name, ...args)` runs a command. */
export function reviewTab(kit, { act }) {
  const { grp, item, menu, btn } = kit;
  const f = (name, ...args) => () => act(name, ...args);
  const protect = btn('🔒', 'Protect Sheet', f('protectBox'), { label: 'Protect Sheet', big: true });
  // (Excel's Protect Workbook stays so named, pressed while the workbook is protected.)
  const book = kit.toggle('📕', 'Protect Workbook', f('protectWorkbookBox'), { label: 'Protect Workbook', big: true });
  const groups = [
    grp('Comments', btn('💬', 'New Comment (Ctrl+Alt+M)', f('commentBox'), { label: 'New Comment', big: true }), btn('🗑', 'Delete Comment', f('deleteNote'), { label: 'Delete' })),
    grp('Notes', menu('🗒', 'Notes', () => [item('New Note', f('editNote'), { hint: 'Shift+F2' }), item('Edit Note', f('editNote')), item('Delete Note', f('deleteNote')),
      item('Show/Hide Note', f('showNote')), item('Show All Notes', f('showAllNotes')), item('Convert to Comments', f('convertNotes'))], { label: 'Notes', big: true })),
    grp('Protect', protect, book, btn('🔓', 'Allow Edit Ranges', f('editRangesBox'), { label: 'Allow Edit Ranges' })),
  ];
  return {
    groups,
    /** Protect Sheet becomes Unprotect Sheet on a protected sheet, as Excel's button does. */
    update(look) {
      const label = look.sheetProtected ? 'Unprotect Sheet' : 'Protect Sheet';
      protect.title = label;
      protect.setAttribute('aria-label', label);
      protect.querySelector('.oxe-lbl').textContent = label;
      book.setOn(look.bookProtected);
    },
  };
}
