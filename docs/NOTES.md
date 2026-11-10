# Notes and comments

How notes and threaded comments are shown (`src/ooxml/notes.js` reads them, `src/editor/notemarks.js` shows them).
They are read only: their parts are saved byte for byte.

## How Excel 365 stores them

Read from a workbook Excel saved with `Range.AddComment` and `Range.AddCommentThreaded` over COM:

- A **note** is a `comment` in the sheet's comments part (`xl/comments1.xml`): its cell, its author, and its text in
  runs. Excel's Insert Note starts with the author's name in bold, then a line break. Its box is a shape in a VML
  drawing (`xl/drawings/vmlDrawing1.vml`), hidden until shown.
- A **threaded comment** is a `threadedComment` in `xl/threadedComments/threadedComment1.xml`, its replies pointing to
  it with `parentId`. Its people are in the workbook's `xl/persons/person.xml`. Excel also writes a placeholder note
  for the same cell ("[Threaded comment] Your version of Excel allows you to read this threaded comment..."), with
  the author `tc={id}`, for older versions. Excel 365 shows the thread, not the placeholder.

`test/fixtures/notes.xlsx` is written that way by `make-fixtures.mjs`, with made-up people. Opened in Excel over COM,
it reads two notes (B2 and D4, by Avery Chen, "Avery Chen:" then the text) and one thread on C3 (Avery Chen, with a
reply from Sam Rivera).

## In the editor

- A small red triangle in the top-right corner of a cell with a note; a purple one for a comment thread, as Excel
  marks them.
- Resting the pointer on the cell shows the note beside it (the author in bold, on Excel's yellow), or the thread:
  each person's name, the date and the text. Shift+F2 shows the active cell's note, and hides it again.
- A screen reader hears the note or thread as the cell's description (`aria-describedby`).

## Not yet

New, edit and delete notes and comments, replies, resolving threads, Show All Notes, a note's own size and position
(from its VML shape), notes following their cells when rows or columns are inserted, sorted or deleted, and marks in
frozen panes' cells drawn twice.
