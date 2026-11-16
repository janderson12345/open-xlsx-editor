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
  each person's name, the date and the text. Show/Hide Note (the cell's menu, Review > Notes) keeps it open.
- A screen reader hears the note or thread as the cell's description (`aria-describedby`).

## Writing notes

There are four ways to work with a note: Shift+F2, the cell's menu (New Note on a cell without one; Edit Note, Delete
Note and Show/Hide Note on one with), and Review > Notes. New Note and Edit Note open a box beside the cell, in a
note's colours, with the author's name above the text. OK or Escape keeps what was typed, as Escape does in Excel's
note; Cancel leaves the note as it was. Each change is one step of the undo history.

A note is written as Excel's New Note writes it (`src/ooxml/noteedit.js`):

- **The comment:** in the sheet's comments part, its text the author's name in bold ("Avery Chen:") and the note's own
  text after a line break, both in Tahoma 9.
- **Its box:** a shape in the sheet's VML drawing, hidden until the pointer rests on the cell.
- **A sheet with no notes yet:** the comments part, the VML drawing (with a block of shape ids no other drawing
  uses), their relationships and content types, and the sheet's `legacyDrawing` are made.

Editing replaces the text and keeps the note's author. Deleting removes the comment and its box. A threaded comment is
deleted with its replies, but threads are not edited here.

Measured: `tools/excel-compare/notes-ours.mjs` writes two new notes on `widths.xlsx` (which has none). On `notes.xlsx`
it edits a note, adds one, deletes another and deletes the thread. Excel opened both (`excel-notes-check.ps1`) and
read every note back with the cell, author and text written (`Comment.Text()`, the bold name and the text), with no
thread left. `test/noteedit.test.js` covers the same in the unit tests.

## A note's box, and Show All Notes

A note's box is drawn where its VML shape puts it, at its size: `x:Anchor`'s corners, each a column or row and a pixel
offset into it (Excel's default for a new note: from the column right of the cell, 15px in, the row above, 10px down,
to two columns and four rows on). A note whose shape has `x:Visible` (Excel's Show/Hide Note) is always shown. Review >
Notes > Show All Notes shows every note's box, or hides them again; as Excel's, it is a view of the notes, not saved.

## Not yet

Show/Hide Note written to the shape (here it shows or hides the box for now, unsaved), and marks in frozen panes'
cells drawn twice. Notes and threaded comments follow their cells when rows and columns are inserted or deleted, and
when they are sorted, as Excel moves them (docs/ROWCOL.md).

## Writing threaded comments (`src/ooxml/threadedit.js`, `src/editor/threadbox.js`)

Review > New Comment (Ctrl+Alt+M, or New Comment on a cell's menu) opens the active cell's thread: post the first
comment, reply, edit any comment, delete a reply or the whole thread. Each is one step of Undo. A cell has a note or a
thread, not both. The author is the editor's `author` option.

Resolve thread (in the thread's box) marks it resolved: measured over COM (`CommentThreaded.Resolved = True`, the
file saved), Excel writes `done="1"` on the thread's first comment, after its id, and nothing on the replies; a reply
added over COM to a resolved thread wasn't saved. Here a resolved thread's box has Reopen and no reply box (nor Edit),
its mark is grey and a screen reader hears "Comment thread, resolved"; Reopen takes `done` away. Each is one step of
Undo. A thread resolved here opens in Excel resolved, its reply kept (`CommentThreaded.Resolved`). The grey mark and
the box's words weren't measured (COM can't show Excel's).

Measured over COM (`tools/excel-compare/excel-threads.ps1`: `AddCommentThreaded`, `AddReply`, `CommentThreaded.Delete`
on a new workbook, saved; the signed-in person Excel signs them with replaced by a made-up one before anything is
written out; `test/threads-excel.json`). The same steps here write each part as Excel does, its made-up ids and times
aside:

| Part | As Excel writes it |
|---|---|
| `xl/threadedComments/threadedCommentN.xml` | A `threadedComment` per comment: `ref`, `dT` (local time to hundredths), `personId`, `id`, and for a reply `parentId`, the thread's first. Line breaks CR LF. |
| `xl/persons/person.xml` | The workbook's people: `displayName`, `id`; for someone not signed in here, `userId` the name and `providerId` None (Excel's signed-in person has their account's). |
| The comments part | A placeholder note for older versions of Excel: "[Threaded comment]", Excel's sentence about it, then each comment as "Comment:" or "Reply:" and its text indented four spaces; its author `tc={id}`; `xr:uid` the thread's id. A deleted thread's author leaves the list and the others are numbered again. |
| The VML drawing | A hidden box as a note has, at Excel's place for it (B2's at 107.25pt, 7.5pt; 48pt a column, 15pt a row). |

Replies made within the same hundredth of a second are ordered by id in Excel, so each reply after the first takes the
one before's id with its third group one more, as Excel's own do. Excel opened the workbook written here
(`excel-threads-check.ps1`, `test/threads-readback-excel.json`) and read both threads, their authors, texts and replies
in order, and no placeholder notes.

## @mentions (`src/editor/mentions.js`, `src/ooxml/threadedit.js`)

Typing @ in a comment or reply (or while editing one) offers people: the host's list (the `people` option, a name
with a directory's `userId` such as an email) and the workbook's own (its persons part), each name once, narrowed by
what follows the @ (a word of the name starting so first, then the name or user id containing it; eight at most). Up
and Down move through them, Enter or Tab (or a click) puts "@Name " in the text, Esc closes the list and not the box.
The box is a combobox for screen readers, the person under the arrow its active option. A mention shows highlighted in
the thread's box and in the box beside the cell.

Excel makes a mention only from a signed-in organisation's directory, which COM can't reach, so a mention can't be made
in Excel to measure. It is written as [MS-XLSX] gives it: the text keeps "@Name", and after it a `mentions` element has
a `mention` for each "@Name" of the people picked still in the text, in the order they come (`mentionpersonId` the
person, `mentionId` an id of its own, `startIndex` where it starts and `length` with the "@"). A name inside a longer
one ("@Ann" in "@Anna Bell") isn't a mention. Editing a comment writes its mentions again.

Excel opened a workbook with three mentions written here (`tools/excel-compare/excel-mentions-check.ps1` on
`test/out/mentions/mentions.xlsx`, `test/mentions-excel.json`): it read the threads, authors and texts, and the copy
it saved kept all three mentions and the persons part as they were, ids and all. Its first reading of an earlier
version showed how it keeps people, which is what is written now:

| Rule | Seen |
|---|---|
| Someone mentioned is a person of their own, `providerId="PeoplePicker"`, with the directory's `userId` (or the name). | Sam Lee, written with provider AD, came back PeoplePicker with `sam.lee@example.com`. |
| The same name as an author and as mentioned: two people. | Avery Chen mentioned and replying: Excel added a None person for the reply and kept the PeoplePicker one for the mention. |
| The persons part is in order of name, then provider (None before PeoplePicker). | Avery Chen None, Avery Chen PeoplePicker, Robin Park None, Robin Park PeoplePicker, Sam Lee. |

Not yet: the e-mail Excel's service sends to someone mentioned (a host can find mentions in the threadedComments part).
