# Notes and comments

How notes and threaded comments are shown (`src/ooxml/notes.js` reads them, `src/editor/notemarks.js` shows them).
What isn't changed here is saved byte for byte; how each change is written is below.

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
  each person's name, the date and the text. Show/Hide Note (the cell's menu, Review > Notes) keeps a note's box
  shown, saved in the file (below); a thread's, until the page closes.
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
to two columns and four rows on). A note whose shape has `x:Visible` is always shown, so a note shown in Excel or here
opens shown.

Show/Hide Note (`noteedit.js setNoteVisible`) writes it there, as Excel does. Measured over COM
(`tools/excel-compare/excel-notes-visible.ps1`, `test/notes-visible-excel.json`): `Comment.Visible = True` on a new
note, the workbook saved, changes two things in its VML shape, and nothing else:

| | Hidden (Excel's new note) | Shown |
|---|---|---|
| The shape's style | `...;z-index:1;visibility:hidden` | `...;z-index:1;visibility:visible` |
| `x:ClientData` | ends with `<x:Column>1</x:Column>` | `<x:Visible/>` after `x:Column`, on its own line |

A note shown and hidden again is saved as it was before it was shown, byte for byte. Here the same: shown, the two
changes (`<x:Visible/>` with x:Column's indent, so Excel's own shape comes out as Excel's); hidden, both taken back,
which gives Excel's hidden shape. Each is one step of Undo. Excel set `Comment.Visible` on a protected sheet too, so it
isn't refused here. A thread's box, and any box in a read-only workbook, is shown or hidden until the page closes.

Review > Notes > Show All Notes is Excel's `Application.DisplayCommentIndicator = xlCommentAndIndicator`, measured the
same way: while it is on, every note reads `Visible` and is saved shown (both changes above); turned off, every note is
hidden, one shown before it was turned on too. So here Show All Notes shows every note in the workbook (on every sheet,
as Excel's setting is for every workbook open), or when every one is shown already, hides them all, saved as Show/Hide
Note writes them, one step of Undo. A thread's placeholder box isn't touched. (Excel's own button is that setting as far
as COM shows; the button itself can't be pressed over COM.)

Excel read back the workbooks written here (`notes-visible-ours.mjs`, `excel-notes-readback.ps1`,
`test/notes-readback-excel.json`): B2 shown here read `Visible` True, D4 shown and hidden again False, and with Show
All Notes both True. `test/notesvisible.test.js` shows and hides Excel's own shapes and gets Excel's files byte for
byte; `test/notes-visible.mjs` checks it in the browser (a shown note saved and opened again opens shown).

## Convert to Comments (`src/ooxml/noteconvert.js`)

Review > Notes > Convert to Comments turns every note into a threaded comment, as Excel's does. Measured over COM with
`Workbook.ConvertComments()` (`excel-notes-visible.ps1`), on a workbook Excel made (notes on two sheets, one shown, a
thread already there) and on one written here (`notes-convert-ours.mjs`: nine notes by made-up people on two sheets,
starting with the author's name in bold, the name not bold, another person's name, a bold "Important:", no name, or
with no author at all, and a thread with a reply):

| What | As Excel converts it |
|---|---|
| Which notes | Every note in the workbook, on every sheet, not only the active one. Threads there already stay as they are. A protected sheet's notes stay (ConvertComments left them, with no error). |
| The text | The note's text, its first line dropped when it is the note's author's name and a colon ("Avery Chen:"), bold or not. "Sam Rivera:" on Avery Chen's note stays, and so does a bold "Important:". Line breaks CR LF. |
| The author | A person by the note's author's name: `userId` the name, `providerId` None, one for each name and used again for the next note of that name. Not the person signed in, and not a person already there with that display name but another `userId`. New people come first in the persons part. A note with no author gets `personId` `{00000000-0000-0000-0000-000000000000}` and no person; Excel then reads its author as "Author". |
| The date | None: the converted comment has no `dT` (Excel's `CommentThreaded.Date` reads empty). |
| Order | The threads in the threadedComments part, and their placeholders in the comments part, by cell row by row, with the threads already there among them. The comments part's authors are just the `tc={id}` of its comments, in order. |
| The boxes | Each note's VML shape removed and a thread placeholder's box added after the shapes there: a new thread's place and size (B2's at 107.25pt, 7.5pt), hidden even if the note was shown. Ids follow on from the drawing's highest (the removed notes' included), and every shape's z-index is numbered again 1, 2, ... in order. |

Here it is done the same way, one step of Undo, a protected sheet left (unless Edit objects is allowed).
`test/notesvisible.test.js` compares each part with Excel's, its new ids aside: the threads (cell, no date, person,
text, reply), the placeholders, the authors, the persons part, and the VML drawing's shapes, ids and z-indexes
included. Excel read the workbook
converted here (`converted.xlsx`) exactly as it read its own conversion of the same workbook: every thread, author,
text and reply, with no dates. `test/notes-visible.mjs` converts notes.xlsx in the browser and undoes it.

Differences from Excel, all in parts the conversion doesn't need to touch: Excel saves every part again, so a thread
placeholder already there gets CR LF and a box sized to its text, and one without an `xr:uid` gets one; here they stay
byte for byte. Excel also reorders the people already in the persons part, by no rule found. COM can't show the
button's own behaviour (whether Excel asks first, or what it does with no notes), so here it converts without asking,
and with no notes it says "No notes to convert" and changes nothing. The text a note keeps when its first line is
"Name: more text", on one line, wasn't measured; here it is kept whole.

## Not yet

Marks in frozen panes' cells drawn twice; Previous and Next Note. Notes and threaded comments follow their cells when
rows and columns are inserted or deleted, and when they are sorted, as Excel moves them (docs/ROWCOL.md).

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
