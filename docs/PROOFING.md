# Spelling and AutoCorrect

What this editor does for spelling, against Excel 365's Review > Spelling and File > Options > Proofing (AutoCorrect),
and how it was measured.

## In the editor

- **AutoCorrect while a cell is typed in** (in the cell or the formula bar; `src/editor/autocorrect.js`), with the
  options Excel ships turned on:
  - Replace text as you type: entries of Excel's own replacement list (`src/editor/autocorrect-list.js`), each as Excel
    has it: teh to the, (c) to ©, ... to …, i to I, dont to don’t. A word typed with a capital first letter keeps it
    (Teh to The).
  - Capitalize names of days (monday to Monday).
  - Correct TWo INitial CApitals (THe to The).
  - Capitalize the first letter of sentences: after a full stop, question or exclamation mark and a space. The cell's
    first word is left as typed.
  - A word is corrected when it ends (a space or punctuation typed after it) and when the cell is entered.
  - Never in a formula (text starting with =, +, or - followed by something other than a number).
  - Ctrl+Z while still typing takes a correction back, giving what was typed.
- **Spelling underlines**: the browser's own, in the cell editor and the formula bar while they hold text (not a
  formula), in the browser's language.

## Measured

| What | How | Seen |
|---|---|---|
| The options on in a new Excel | `Application.AutoCorrect` over COM | ReplaceText, CorrectSentenceCap, CapitalizeNamesOfDays, TwoInitialCapitals, CorrectCapsLock all True. |
| The replacement list | `AutoCorrect.ReplacementList()` | 913 entries (en-US); 39 are carried here, each with Excel's replacement. |

When a correction happens while typing can't be measured over COM (typing into Excel needs keystrokes); it follows
Microsoft's documentation of AutoCorrect.

## Not yet

- **Review > Spelling (F7)**: Excel's box going through the sheet's cells with suggestions, Ignore, Change and Add to
  Dictionary. Browsers don't let a page ask their spell checker about a word, so this needs a dictionary of its own.
- The rest of Excel's replacement list (874 entries), its AutoCorrect exceptions (e.g. and other abbreviations whose
  full stop doesn't end a sentence), Correct accidental use of cAPS LOCK, and the AutoCorrect Options box to turn
  any of it off or add entries.
- Excel underlines nothing in a cell while it is typed; the underlines here are the browser's, an addition.
