# Dynamic arrays

Excel 365's dynamic array formulas: a formula whose answer is an array spills it into the cells below and to the right
(`src/calc/spill.js`), with the functions that make arrays (`src/calc/functions-dynamic.js`), LET and LAMBDA
(`src/calc/evaluate.js`), `D1#` and `@` (`src/calc/tokens.js`, `parse.js`, `fileform.js`), and the file's form of all of them
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
  name comes first (=Fact(5) is FACT's); typed in another case it is written as the name is (=twice(5) is
  Twice(5), measured). Ten such formulas in the probe, all as Excel's.
- Rows inserted through a spill: the spill is worked out again from its cell, as Excel does. A sort whose range takes in
  any part of a spill, its formula's cell in it or not, is refused with Excel's "You can't change part of an array."
  (measured: five ranges through C1's =SEQUENCE(6,1,10,-1), refused; A1:B6 beside it sorts), as an old array
  formula's is.
- Formulas in workbooks saved by older versions of Excel keep their old meaning: `=A1:A3` there is one value, A1:A3
  cut to the formula's row. The formula bar (and F2) shows Excel 365's `@` wherever such a formula cuts a range or an
  array to one value (`src/calc/legacyat.js`), so that typed back it keeps that meaning. Measured on 58 formulas (set by
  Range.Formula in a hidden Excel, read back by Formula2), 57 the same here: a range where one value is wanted,
  `=@A1:A3*2`, `=SUM(@A1:A3*2)`, `=VLOOKUP(@A1:A3,A1:B3,2,FALSE)`, `=COUNTIF(A1:A3,@A1:A3)`; a function giving
  several cells there, `=@OFFSET(A1,0,0,2)`, `=@CHOOSE(1,A1:A3)`, `=@INDEX(A1:A3,0,1)`, `=@ROW(A1:A3)`; none where the
  range is taken whole, `=SUMPRODUCT(A1:A3*2)`, `=MATCH(2,A1:A3,0)`, `=SUMIF(A1:A3,">1")`, `=N(A1:A3)`.

## The file (measured)

| What | As Excel 365 writes it |
|---|---|
| A formula whose answer is an array (even of one value, `SEQUENCE(1)`, or a function of arrays answering one value, FILTER's if_empty) | `<c r="D1" cm="1"><f t="array" ref="D1:D4">_xlfn.SEQUENCE(4)</f><v>1</v></c>`, the other cells values only (`t="str"` for text) |
| `cm="1"` | the cell metadata block of `xl/metadata.xml`'s XLDAPR entry (`xda:dynamicArrayProperties fDynamic="1"`), the part related from the workbook as sheetMetadata |
| A formula answering one value (LET, `SUM(SEQUENCE(4))`) | a plain `<f>` |
| A formula with a LAMBDA call giving a value anywhere in it (`LAMBDA(x,x*2)(3)+1`, `Twice(3)`, a LET name called, `IF(TRUE,LAMBDA(x,x)(5))`), or MAP, REDUCE and the others, even inside `SUM` | a dynamic one, `t="array"` and `cm="1"`, though its answer is one value; a call giving a reference (`SUM(LAMBDA(x,x)(A1:A3))`) leaves it plain (14 measured, all as Excel's) |
| A LAMBDA kept in a LET name, called (`f(3)`) | `_xlpm.f(3)` |
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

GROUPBY, PIVOTBY, PERCENTOF and RANDARRAY (`src/calc/functions-groupby.js`), measured on 55 formulas in a hidden Excel
(each spilled answer read cell by cell), then 24 in the formula probe, all as Excel's:

- **Functions as arguments**: a function named on its own where a LAMBDA goes is one (an eta LAMBDA, =MAP(A1:A3,SQRT));
  the file has it as `_xleta.SQRT`, and HSTACK(SUM,AVERAGE) for several. Typed here it is written so; shown without.
  On its own in a cell, =SQRT is #CALC!.
- **Headers**: left out, the first row is a header when the first value is text and the one below a number (dropped,
  not shown); 0 none, 1 there and dropped, 2 made up ("Row Field 1", "Value 1"), 3 there and shown. Several functions
  add a row of their names (SUM, COUNT) above.
- **Totals**: "Total" at the end; with subtotals (total_depth 2) a row for each group after its rows, and "Grand Total";
  negative depths put them before. field_relationship 1 with subtotals is #VALUE!.
- **Sorting**: by the fields; sort_order a field (negative: descending) or a value column (by the groups' answers);
  an array of fields, {1,-2} (in any order; a field twice, or a field with a value column, #VALUE!). Field values go
  numbers, text, TRUE and FALSE, then a blank, which is a group of its own and stays last sorted descending too.
  Sorted by a value column with total_depth left out, every level gets its subtotal.
- **PIVOTBY**: the column fields' values across, one header row for each field, "Total" after (or before); a missing
  group "" (an empty cell of the answer is "", not a blank, in both); with several value columns the total and
  subtotal columns are left "", as Excel leaves them. Several row or column fields (`src/calc/functions-pivotby.js`):
  subtotal rows and columns for the levels the depths reach; the column fields' header in one cell ("Item, Size",
  made up "Column Field 1, Column Field 2"). Sorted by a value column, rows go by their totals at every level,
  columns only at the last level and where they have subtotals (the others by the field), as Excel sorts them.
  PERCENTOF (and a LAMBDA of two arguments) of relative_to's total: 0 the column's, 1 the row's, 2 the grand total,
  3 the parent column's, 4 the parent row's; 5 #VALUE!.
  Measured on 104 formulas in a hidden Excel (two row fields, two column fields, three, with subtotals, headers 1 to
  3, sort arrays, several functions and value columns, filters, every relative_to), all as Excel's, then 12 more in
  the formula probe.
- **RANDARRAY**: 0 rows or columns #CALC!, fewer #VALUE!, min above max #VALUE!, whole numbers between bounds that
  aren't whole #VALUE!.

TRIMRANGE and the trim references (`src/calc/functions-groupby.js`, `tokens.js`, `fileform.js`), measured on 27 formulas in a hidden
Excel on a sheet with B3 1, C4 2, B6 3 and D3 `=""` (every spilled cell read), all 27 as Excel's here, then 15 in the
formula probe:

- **TRIMRANGE(range,[trim_rows],[trim_cols])** cuts the empty rows and columns off a range's edges: 0 none, 1 leading,
  2 trailing, 3 both (left out). A formula answering "" is not empty. A1:E8 is B3:D6; with trim_cols 1, B3:E6. Above
  3, #VALUE!; a range with nothing in it, #REF!; an array or one cell is as it was. A:E is B3:D6 as well.
- **Trim references**: `A1:.E8` trims trailing empty rows and columns (A1:D6), `A1.:E8` leading (B3:E8), `A1.:.E8`
  both. Whole columns and rows too: `B:.B` is 6 rows, `3:.6` 4. A reference with nothing in it is #REF!.
- **The file**: `_xlfn._TRO_TRAILING(A1:E8)`, `_xlfn._TRO_LEADING`, `_xlfn._TRO_ALL`, and `_xlfn.TRIMRANGE`. Read
  and written so here, shown as typed.

## Not yet

- The blue outline Excel draws round a spill when its cell is selected.
- (TAKE, DROP, VSTACK, HSTACK, CHOOSEROWS and the other newer array functions, MAP, REDUCE, SCAN, BYROW, BYCOL,
  MAKEARRAY, ISOMITTED, GROUPBY, PIVOTBY, PERCENTOF, RANDARRAY and TRIMRANGE are here, in the formula probe.)
- An old formula's `@` is found by working the formula out at its cell, so where that gives an error before the array
  comes out it isn't shown: `=SEQUENCE(A1:A3)` below row 3 is `=SEQUENCE(@A1:A3)` here, `=@SEQUENCE(@A1:A3)` in Excel.
