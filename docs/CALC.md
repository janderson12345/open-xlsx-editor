# Calculation

How formulas are worked out (`src/calc/`), and what Excel was seen to do. As with layout, Excel's own results are
the only measure. `test/fixtures/formulas.xlsx` holds 541 formula cells with no saved values, and a table (Fruit) for table references. Excel opened it,
worked them out (`tools/excel-compare/excel-values.ps1`, COM, read-only, `CalculateFull`) and its results are in
`test/formulas-excel.json`. `npm test` checks every one: **541 of 541 as Excel** (to 15 digits; RATE and IRR, searches Excel promises to 0.0000001, to 1e-9).

## How it works

- **Reading formulas.**
  - `tokens.js` reads a formula into tokens.
  - `parse.js` builds a tree by Excel's precedence. Negation binds tightest, so `=-2^2` is 4, and operators of one
    level work left to right, so `=2^3^2` is 64.
- **Working them out.**
  - `evaluate.js` handles the operators, references and names.
  - About 130 functions in three files, by Excel's categories: `functions.js` (maths, statistics, logic, information
    and the IF/IFS family, plus the helpers the others share), `functions-text.js` (text, and dates and times) and
    `functions-lookup.js` (lookup: VLOOKUP, INDEX, MATCH, XLOOKUP, OFFSET, INDIRECT; and SUMIF, COUNTIFS and the other
    conditional sums and counts).
  - `functions-more.js` holds 54 more: SUBTOTAL (hidden rows left out by 101 to 111, other subtotals always), RANK,
    PERCENTILE, QUARTILE and MODE; NETWORKDAYS, WORKDAY, DATEDIF, WEEKNUM, ISOWEEKNUM and YEARFRAC; PMT, PV, FV, NPER,
    RATE, NPV and IRR; LOOKUP, XMATCH and ADDRESS; TEXTBEFORE, TEXTAFTER, NUMBERVALUE, FIXED and DOLLAR; trigonometry,
    GCD, LCM, CEILING.MATH and FLOOR.MATH; TYPE and ERROR.TYPE. Excel raises 1 + rate to a whole number of periods by
    repeated squaring in PV and FV but through logarithms in PMT, and each is computed so, to Excel's last digit. RATE
    and IRR are found by Newton's method (IRR on the discount factor, so IRR({-100,10}) is -0.9).
  - Table references (`ooxml/tables.js`): Fruit[Qty] is the column's data rows; [#All], [#Data], [#Headers], [#Totals]
    and two of them together; [[Qty]:[Price]] the columns between; [#This Row] (shown [@Qty]) the formula's own row,
    from any sheet by the row's number, #VALUE! past the table's rows (Excel's answers in test/formulas-excel.json).
    Names compare case aside. Typing a table reference into a cell isn't read yet.
  - `values.js` defines the values and how each is coerced into another.
- **Recalculating** (`engine.js`):
  - **What a change reaches.** On the first change every worksheet's formulas are read, with what each one reads,
    indexed by column. A change then works out again the formulas it reaches, along with volatile ones (NOW, TODAY,
    RAND, OFFSET, INDIRECT).
  - **Order.** Each formula is worked out after the ones it reads. A depth-first walk with its own stack does this,
    so a column of 10,000 chained formulas doesn't overflow the call stack.
  - **Results.** Each result is written into the cell's `<v>`, and its `t` set to `str`, `b` or `e`. The formula
    itself is kept as written.
- **Formulas Excel saved are trusted until something they read changes,** as Excel trusts them when it opens a file.
- **Opening.** Formulas the file came without a result for are worked out (`Workbook.calculateMissing`), with the
  formulas that read them, as Excel does when it opens such a file; the results saved with the others are kept. They
  are shown at once, and saved with a sheet once it is edited, so a workbook saved unchanged is the same bytes.
- **Shared formulas.** A shared formula's later cells carry only `si`. They read the group's first formula, moved to
  their own position (`shiftFormula`). When one cell of the group changes, every cell is first written out with its
  own formula.
- **Array formulas** (`t="array"`) fill their range. One cell of an array can't be changed, as in Excel. Data
  tables keep the values Excel saved.
- **calcChain.xml** is dropped when a formula is added or removed. A chain that names a cell which no longer holds a
  formula makes Excel offer a repair; without a chain, Excel works out the order again.
- **What the engine can't work out:** an unknown function, a link to another workbook, R1C1 in INDIRECT, a data
  table. These cells keep the value Excel saved (a new entry shows `#NAME?`), and the saved workbook gets
  `<calcPr fullCalcOnLoad="1"/>` so that Excel works everything out again when it opens the file.
- **Typing a formula** (`normalizeFormula`):
  - function names, references and TRUE/FALSE go into capitals
  - newer functions get the `_xlfn.` prefix the file needs, and lose it again on screen (`displayFormula`)
  - missing closing brackets are added
  - a formula that doesn't parse is refused with Excel's words, "There's a problem with this formula", and stays in
    the editor
- **Pointing** (`src/editor/pointing.js`), as Excel's Point mode works:
  - **Where.** A reference can go after `=`, an operator, `(`, `,` or `:`, before the end, `)` or another operator,
    and not inside quotes.
  - **Arrow keys.** In Enter mode (typing over a cell) they point: from the active cell, Shift extends, Ctrl jumps
    over data. In Edit mode (F2, which switches between the two) they move the caret.
  - **Mouse.** A click or a drag on the grid points too, and Shift+click extends.
  - **Ending.** Typing anything else ends the reference, and the next arrow starts a new one.
  - **F4** cycles the reference at the caret: A1, $A$1, A$1, $A1.
  - **Range finder.** The formula's ranges are outlined in colours, and the one being pointed at is dashed. Excel's
    own colours aren't in any file to measure, so these are ours. The reference text isn't coloured in the box.
  - **Other sheets.** A sheet tab clicked where a reference can go shows that sheet. The sheet's name goes into the
    formula at once (`'Notes & more'!`), and the formula goes on in the formula bar. The arrows start from that
    sheet's active cell. Enter puts the formula into its own cell, on its own sheet; Esc goes back without it.
  - **Merged cells.** A merged cell is pointed at by its top-left cell. Excel's behaviour here isn't measured.

## What Excel does (measured)

| Rule | Seen |
|---|---|
| A number becomes text with 15 significant digits. It is written out in full when that takes at most 20 characters, and in scientific notation otherwise. | `=1/3&""` is "0.333333333333333"; "0.000000000000000001", but "1E-19"; "0.000123456789012345", but "1.23456789012345E-05"; "12345678901234500000", but "1E+20". |
| A number written in a formula keeps 15 significant digits; the rest are cut off, not rounded. | `=1234567890123456` is 1234567890123450. |
| When a formula's last operation is + or -, and its two operands agree to 15 significant digits, the result is 0. Brackets, or any later operation, stop this. SUM does the same at each step. | `=0.1+0.2-0.3` is 0, but `=(0.1+0.2-0.3)` and `=1*(0.1+0.2-0.3)` are 5.55E-17. `=1+1E-15-1` is 0, but `=1-1E-15-1` is -9.99E-16. `=SUM(0.1,0.2,-0.3)` is 0. |
| Comparisons treat numbers that agree to 15 significant digits as equal. | `=0.1+0.2=0.3` is TRUE. |
| Outside an array formula, a range given to an operator is cut to the formula's own row or column (implicit intersection). An array constant stays whole. | In row 5, `=SUM(Data!B2:B9*1)` is Data!B5. `=SUM({1,2,3}*2)` is 12, and `={1,2,3}*2` shows 2. |
| SUMPRODUCT's arguments are worked out as arrays. | `=SUMPRODUCT((B2:B9>3)*C2:C9)`, `=SUMPRODUCT(--ISNUMBER(SEARCH("an",A2:A9)))`. |
| A negative number to a fractional power is #NUM!, unless the power is an odd root. | `=(-8)^(1/3)` is -2; `=0^0` is #NUM!. |
| An argument left out counts in COUNT and COUNTA, and counts as 0 in AVERAGE. | `=COUNT(1,)` is 2; `=AVERAGE(1,)` is 0.5. |
| An equal-to criterion matches a number stored as text too; the other operators don't. | `COUNTIF(B2:B9,12)` counts the text "12"; `">=4"` doesn't. |
| A blank criterion with `=` matches blank cells only; with no operator, blank cells and "". A whole-column criterion counts every blank cell. | `COUNTIF(F:F,"")` is 1048576. |
| MOD works on the binary values, without the 15-digit rounding. | `=MOD(0.3,0.1)` is 0.09999999999999998. |
| TIME is the fraction of a day with whole days taken off. | `=TIME(25,0,0)` is 0.04166666666666674. |
| ROUND, ROUNDUP and ROUNDDOWN work on the value's 15 significant decimal digits. | `=ROUND(2.675,2)` is 2.68. |
| Text compares without regard to case, in Excel's order. Numbers come before text, and text before FALSE and TRUE. A blank counts as 0, "" or FALSE, whichever the other value is. | `="a"="A"`, `=1<"a"`, `="a"<TRUE`, `=Data!F2=""` (blank) are all TRUE. |
| Approximate lookups on a list sorted up (MATCH 1, VLOOKUP and HLOOKUP TRUE) are a binary search, which on a list that isn't sorted gives its own answer. From the middle: a blank or a value of another type is passed over to the right; an exact match is taken at once, then the last of the equal values after it. On a list sorted down (MATCH -1) it is a scan from the start, blanks and other types passed over, taking the first exact match and stopping at the first value below. | 400 unsorted lists, some with blanks and text, 1,200 answers (`test/lookup-excel.json`, from `tools/excel-compare/excel-lookups.ps1`): all as Excel's. `MATCH(8,{3,6,9,6,6,1,8,7},1)` is 7, not 2 or 8; `MATCH(6,{5,8,7,1},-1)` is #N/A. |

## Not yet

- **Spilling.** Dynamic arrays (FILTER, SORT, UNIQUE, SEQUENCE) aren't supported, and neither is the `@` operator.
  A formula's array result shows its first value.
- **Intersection and 3D references.** The space operator (intersection) and 3D references (`Sheet1:Sheet3!A1`)
  aren't read, nor are structured references to tables. Formulas using them keep Excel's values.
- **Iterative calculation.** Circular references are left as they were and announced.
- **Dates and times typed as entries** are read as dates and times (docs/ENTRY.md). Inside a formula, text such as
  "2024-01-02" or "1/2/2024" reads as a date.
