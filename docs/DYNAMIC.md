# Dynamic arrays

Excel 365's dynamic array formulas: a formula whose answer is an array spills it into the cells below and to the right
(`src/calc/spill.js`), with the functions that make arrays (`src/calc/functions-dynamic.js`), LET and LAMBDA
(`src/calc/evaluate.js`), `D1#` and `@` (`src/calc/tokens.js`, `parse.js`), and the file's form of all of them
(`src/ooxml/spillxml.js`, `dynamicmeta.js`).

## In the editor

- A formula typed (or pasted, filled, entered by the agent) is Excel 365's kind: `=SORT(A1:A6)` or `=A1:A3*2` spills
  down from its cell, and every cell of the spill shows its value.
- A value typed in the way of a spill turns the formula into `#SPILL!`; Delete on it, or Ctrl+Z, lets it spill again.
  Typing over a spilled cell's value is how a person blocks it; Delete on one leaves the spill as it was.
- When the answer grows, shrinks or goes, the cells it no longer fills are emptied, and formulas reading them work out
  again. `=SUM(D1#)` reads whatever D1 spills into.
- Functions: SEQUENCE, SORT, SORTBY, UNIQUE, FILTER (with if_empty, else `#CALC!`), TRANSPOSE; LET (names bound in turn)
  and LAMBDA (called where written, `=LAMBDA(x,x*2)(A1)`, through a LET name, or kept in a defined name and called by
  it, `=Twice(21)`); `@` for one value by implicit intersection.
- A LAMBDA kept in a defined name (Excel saves it as `_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)`, measured), as Excel 365 calls
  it: =Twice(21) is 42, =WithRate(100) reads the names its LAMBDA reads, =Twice({1,2,3}) spills 2, 4, 6,
  =MAP({1,2},Twice) passes it on, the wrong number of arguments is #VALUE!, =Twice alone #CALC!; a function of the same
  name comes first (=Fact(5) is FACT's). Ten such formulas in the probe, all as Excel's.
- Formulas in workbooks saved by older versions of Excel keep their old meaning: `=A1:A3` there is one value, A1:A3
  cut to the formula's row (Excel 365 shows it as `=@A1:A3`, and so does the formula bar here).

## The file (measured)

| What | As Excel 365 writes it |
|---|---|
| A formula whose answer is an array (even of one value, `SEQUENCE(1)`, or a function of arrays answering one value, FILTER's if_empty) | `<c r="D1" cm="1"><f t="array" ref="D1:D4">_xlfn.SEQUENCE(4)</f><v>1</v></c>`, the other cells values only (`t="str"` for text) |
| `cm="1"` | the cell metadata block of `xl/metadata.xml`'s XLDAPR entry (`xda:dynamicArrayProperties fDynamic="1"`), the part related from the workbook as sheetMetadata |
| A formula answering one value (LET, `SUM(SEQUENCE(4))`) | a plain `<f>` |
| `D1#` | `_xlfn.ANCHORARRAY(D1)` |
| `@x` inside a formula | `_xlfn.SINGLE(x)`; a formula that is all `@` (`=@A1:A3`) is written plain, `A1:A3`, the old kind |
| LET and LAMBDA names | `_xlpm.x` |
| SORT and FILTER | `_xlfn._xlws.SORT`, `_xlfn._xlws.FILTER`; the others `_xlfn.` |
| `#SPILL!`, `#CALC!` | rich values Excel keeps in other parts. Written here as `#VALUE!` with the workbook marked to be worked out again on opening (`fullCalcOnLoad`), where Excel shows them as its own. |

## Measured against Excel (11 November 2026)

`tools/excel-compare/dynamic-cases.mjs` makes 20 cases: each formula entered in D1 with `Formula2` in a hidden Excel
365 (`excel-dynamic.ps1`), numbers in A1:A6 and words in B1:B6: SEQUENCE (two ways), SORT (numbers, words descending,
two columns), SORTBY, UNIQUE (and exactly once), FILTER (some, none with if_empty, none), a range times 2, TRANSPOSE, a
spill blocked, `SUM(D1#)`, LET, LAMBDA called, `@` on a range, `SUM(SEQUENCE(4))`, `SEQUENCE(1)`.

| Check | Result |
|---|---|
| Typed here (`test/dynamic.test.js`): the spill range, every cell's value or text, the formula as shown, the anchor's form in the file | 20 of 20 Excel's |
| Saved here and opened in Excel (`dynamic-ours.mjs`, `excel-dynamic-check.ps1`): Excel's spill range, values, texts and formulas | 20 of 20 the same as Excel's own |

## Not yet

- The blue outline Excel draws round a spill when its cell is selected.
- RANDARRAY, GROUPBY, PIVOTBY, TRIMRANGE and PERCENTOF. (TAKE, DROP, VSTACK, HSTACK, CHOOSEROWS and the other newer
  array functions, MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY and ISOMITTED are here, in the formula probe.)
- A defined LAMBDA called with a name typed in another case (=twice(5)) is saved as TWICE(5); Excel writes it as the
  name is (Twice(5)). Both work out the same, in both programs.
- `@` shown by the formula bar only where an old formula is a whole range; Excel 365 also shows it before a range an
  old formula's operator cuts to one value (`=@A1:A3*2`).
- A LAMBDA called where it is written is saved as a plain formula when its answer is one value; Excel saves it as a
  dynamic one (both work out the same, in both programs).
- Rows inserted or sorted through a spill: the spill is worked out again from its cell, as Excel does, but sorting rows
  that hold spilled values sorts the values, where Excel refuses to sort part of a spill.
