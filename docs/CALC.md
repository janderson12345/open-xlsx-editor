# Calculation

How formulas are worked out (`src/calc/`), and what Excel was seen to do. As with layout, Excel's own results are
the only measure. `test/fixtures/formulas.xlsx` holds 1029 formula cells with no saved values, and a table (Fruit) for table references. Excel opened it,
worked them out (`tools/excel-compare/excel-values.ps1`, COM, read-only, `CalculateFull`) and its results are in
`test/formulas-excel.json`. `npm test` checks every one: **1029 of 1029 as Excel** (to 15 digits; RATE, IRR and YIELD, searches (Excel promises RATE and IRR to 0.0000001), to 1e-9; XIRR, which it promises to 0.000001%, to 1e-8;
the t, chi-squared, F, gamma, beta, Weibull and negative binomial distributions, their inverses and tests, and VDB, to
1e-13 of the value: series, continued fractions and searches whose last digits are each one's own, Excel's within
2e-14 of these).

## How it works

- **Reading formulas.**
  - `tokens.js` reads a formula into tokens (`prefixed.js`: the newer functions the file names with `_xlfn.`).
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
  - `functions-arrays.js`: Excel 365's TAKE, DROP, VSTACK, HSTACK, CHOOSECOLS, CHOOSEROWS, TOCOL, TOROW, WRAPROWS,
    WRAPCOLS, EXPAND, TEXTSPLIT, VALUETOTEXT and ARRAYTOTEXT, and the LAMBDA helpers MAP, REDUCE, SCAN, BYROW, BYCOL,
    MAKEARRAY and ISOMITTED (docs/DYNAMIC.md for spilling).
  - `functions-finance.js`: IPMT, PPMT, CUMIPMT, CUMPRINC, SLN, DDB, DB, EFFECT, NOMINAL, XNPV, XIRR; TIMEVALUE (read as an
    entry is read), WORKDAY.INTL and NETWORKDAYS.INTL (Excel's weekend numbers and seven-day strings). As with PV and FV,
    each computed the way that gives Excel's last digit: DDB's book value and EFFECT by repeated squaring, IPMT's
    balance through logarithms.
  - `functions-stats.js`: CORREL, PEARSON, SLOPE, INTERCEPT, RSQ, FORECAST and FORECAST.LINEAR, COVARIANCE.P and .S,
    GEOMEAN, HARMEAN, AVEDEV, DEVSQ, SKEW, KURT, STANDARDIZE, PERMUT, FISHER; NORM.DIST, NORM.S.DIST, NORM.INV, NORM.S.INV
    (W. J. Cody's erfc, the inverse polished by Halley's method: to 15 digits), EXPON.DIST, POISSON.DIST, BINOM.DIST;
    SINH, COSH, TANH, SEC, CSC, COT; ROMAN, ARABIC, BASE, DECIMAL, DEC2BIN, BIN2DEC, DEC2HEX, HEX2DEC, MULTINOMIAL;
    ISFORMULA, FORMULATEXT, SHEET and SHEETS.
  - `functions-dist.js`: the distributions the others lack, on Lanczos' log-gamma and the regularised incomplete gamma
    (series and continued fraction) and beta (continued fraction) functions, inverses by bisection then Newton: T.DIST,
    .2T, .RT, T.INV, .2T; CHISQ.DIST, .RT, CHISQ.INV, .RT; F.DIST, .RT, F.INV, .RT; GAMMA, GAMMALN(.PRECISE), GAMMA.DIST,
    GAMMA.INV; BETA.DIST, BETA.INV; LOGNORM.DIST, LOGNORM.INV; WEIBULL.DIST; NEGBINOM.DIST; HYPGEOM.DIST (the ways
    multiplied out, so exact); BINOM.INV; CONFIDENCE.NORM and .T; Z.TEST, T.TEST (paired, equal and unequal
    variances), CHISQ.TEST, F.TEST; PHI, GAUSS, FISHERINV; and the names of Excel 2007 and before for each (TDIST,
    CHIDIST, FDIST, NORMDIST, CRITBINOM...). Degrees of freedom are cut to whole numbers, T.DIST's too.
  - `functions-agg.js`: AGGREGATE (its options leave out nested subtotals, hidden rows and errors; functions 1 to 13
    take references only, an array #VALUE!), FREQUENCY, TRIMMEAN, PERCENTRANK(.INC, .EXC), MODE.MULT; VARA, VARPA,
    STDEVA, STDEVPA (text 0 and TRUE 1 in ranges, only the numbers of an array), STEYX, PROB; MMULT, MDETERM, MINVERSE,
    MUNIT; SUMX2MY2, SUMX2PY2, SUMXMY2; LINEST, LOGEST, TREND and GROWTH. LINEST takes the xs and ys from their means
    first, as Excel does, so an exact line has residuals of exactly 0 and its F is #NUM!.
  - `functions-eng.js`: BITAND, BITOR, BITXOR, BITLSHIFT, BITRSHIFT; OCT2DEC, DEC2OCT, BIN2HEX, HEX2BIN, BIN2OCT,
    OCT2BIN, HEX2OCT, OCT2HEX; DELTA, GESTEP, ERF, ERFC (.PRECISE); ACOSH, ASINH, ATANH, ACOT, ACOTH, COTH, CSCH, SECH;
    COMBINA, PERMUTATIONA, FACTDOUBLE, SERIESSUM, SQRTPI; DAYS360 (Excel's US method leaves an end on the last of
    February as it is); HYPERLINK, AREAS; FVSCHEDULE, PDURATION, RRI, ISPMT, MIRR (from the two NPVs, as Excel defines
    it), SYD, VDB, DOLLARDE, DOLLARFR, TBILLPRICE, TBILLYIELD, TBILLEQ.
  - `functions-db.js`: the database functions (DSUM, DAVERAGE, DCOUNT, DCOUNTA, DGET, DMAX, DMIN, DPRODUCT, DSTDEV,
    DSTDEVP, DVAR, DVARP), their criteria read as Advanced Filter reads them (`criteria.js`, shared with
    `ooxml/advfilter.js`); CEILING.PRECISE, FLOOR.PRECISE, ISO.CEILING, SKEW.P, BINOM.DIST.RANGE, COVAR.
  - `functions-bonds.js`: securities. COUPPCD, COUPNCD, COUPNUM, COUPDAYBS, COUPDAYS, COUPDAYSNC; PRICE, YIELD,
    DURATION, MDURATION; ACCRINT, ACCRINTM; DISC, INTRATE, RECEIVED, PRICEDISC, YIELDDISC, PRICEMAT, YIELDMAT. Coupon
    dates step back from maturity, kept at the month's end when maturity is. What Excel was seen to do (each measured
    with Evaluate in a hidden Excel, then checked by the probe on cases it wasn't worked out from):
    - COUPDAYS with basis 1 runs from the coupon date before settlement to the same day 12/frequency months on
      (February 29 to August 29: 182 days, not the 184 to August 31). PRICE, YIELD and DURATION take the actual period
      (184) instead. COUPDAYSNC counts 30/360 days with basis 4 and takes COUPDAYS - COUPDAYBS with basis 0.
    - PRICE, YIELD and DURATION take the days to the next coupon as COUPDAYS - COUPDAYBS, but for basis 1.
    - ACCRINT: the coupon periods from issue to settlement, stepped from first_interest. A period that ends before
      settlement counts 1. A part counts its days (by the basis) over the period's length: 360 or 365 over the
      frequency, or with basis 1 the days of the period that ends at first_interest. A period ending on settlement is a
      part (August 31 to February 28: 178 of 180). With calc_method FALSE, the days from issue to settlement over that
      length, whatever Microsoft's page says about first_interest.
    - DISC and its kin with basis 1 take YEARFRAC's year (366 over a February 29); PRICEMAT and YIELDMAT the year up
      to settlement's. DISC is (1 - pr/redemption) x (year / days), to Excel's last digit.
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

## The Function Library

- **The list** (`calc/catalog.js`): every function the engine works out, in Excel's categories as Microsoft's "Excel
  functions (by category)" has them, and its syntax with Excel's argument names, the optional ones in brackets
  (`SUM(number1, [number2], ...)`). Where that list names a function under two categories, it is in one: CONCATENATE
  in Text, FORECAST in Statistical, FLOOR in Math & Trig, LET in Logical. DATEDIF is left out, as Excel leaves it out of
  its lists. `test/catalog.test.js` checks that the list and the engine agree.
- **Insert Function** (Shift+F3, or fx on the Formulas tab; `editor/functionbox.js`): Excel's box. Search for a
  function and Go lists the functions whose names hold the words, under Recommended. A category shows its functions:
  Most Recently Used (Excel's list to start with: SUM, AVERAGE, IF, HYPERLINK, COUNT, MAX, SIN, SUMIF, PMT, STDEV,
  then the ones picked, newest first), All, or one of Excel's. The syntax of the one picked is under the list. OK opens
  the active cell with `=NAME(`. In a cell being edited, Shift+F3 keeps it open and types `NAME(` at the caret.
- **The Formulas tab's lists**: Recently Used, Financial, Logical, Text, Date & Time, Lookup & Reference, Math & Trig,
  and More Functions (Statistical, Engineering, Information, Compatibility). Database is in Insert Function only, as
  in Excel. A function picked is typed as Insert Function types it. Each one's syntax is its tooltip.
- **The screen tip** (`editor/fntip.js`): inside a function's parentheses in a cell being edited, its syntax under the
  cell, the argument the caret is in in bold. Repeated arguments go on with the last one named (SUMIFS's pairs two by
  two). Text in quotes, sheet names and array constants are passed over.
- **Not yet**: Excel's description of each function, and its Function Arguments box (Excel opens it after Insert
  Function; here the arguments are typed).

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

- **ACCRINT with basis 1 over many coupon periods.** Excel's answer differs from the rule above in the fourth digit
  (ACCRINT(DATE(2005,7,10),DATE(2006,1,31),DATE(2009,6,20),0.06,100,2,1) is 23.6307 in Excel, 23.625 here); its own
  rule for that case isn't found yet. Shorter spans, and every other basis, are Excel's.
- **Not written yet**: the odd-period bonds (ODDFPRICE, ODDFYIELD, ODDLPRICE, ODDLYIELD), AMORDEGRC and AMORLINC,
  CONVERT, the complex-number functions (IMSUM and the rest), the BESSEL functions and the REGEX functions.
- **Spilling** is in docs/DYNAMIC.md (dynamic arrays, LET, LAMBDA, `D1#`, `@`); its own gaps are listed there.
- **Intersection and 3D references.** The space operator (intersection) and 3D references (`Sheet1:Sheet3!A1`)
  aren't read, nor are structured references to tables. Formulas using them keep Excel's values.
- **Iterative calculation.** Circular references are left as they were and announced.
- **Dates and times typed as entries** are read as dates and times (docs/ENTRY.md). Inside a formula, text such as
  "2024-01-02" or "1/2/2024" reads as a date.
