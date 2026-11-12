# Calculation

How formulas are worked out (`src/calc/`), and what Excel was seen to do. As with layout, Excel's own results are
the only measure. `test/fixtures/formulas.xlsx` holds 1601 formula cells with no saved values, and a table (Fruit) for table references. Excel opened it,
worked them out (`tools/excel-compare/excel-values.ps1`, COM, read-only, `CalculateFull`) and its results are in
`test/formulas-excel.json`. `npm test` checks every one: **1601 of 1601 as Excel** (to 15 digits; RATE, IRR, YIELD and ODDFYIELD, searches (Excel promises RATE and IRR to 0.0000001), to 1e-9; XIRR, which it promises to 0.000001%, to 1e-8;
the t, chi-squared, F, gamma, beta, Weibull and negative binomial distributions, their inverses and tests, and VDB, to
1e-13 of the value: series, continued fractions and searches whose last digits are each one's own, Excel's within
2e-14 of these).

## How it works

- **Reading formulas.**
  - `tokens.js` reads a formula into tokens; `fileform.js` turns it between the typed form and the file's (`prefixed.js`:
    the newer functions the file names with `_xlfn.`).
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
  - `functions-complex.js`: the complex numbers, text such as "3+4i" or "3+4j". COMPLEX, IMREAL, IMAGINARY, IMABS,
    IMARGUMENT, IMCONJUGATE, IMSUM, IMSUB, IMPRODUCT, IMDIV, IMPOWER, IMSQRT, IMEXP, IMLN, IMLOG10, IMLOG2, IMSIN,
    IMCOS, IMTAN, IMCOT, IMSEC, IMCSC, IMSINH, IMCOSH, IMSECH, IMCSCH. What Excel was seen to do (about 350 cases with
    Evaluate, then the probe's 105):
    - Reading: the real part, then the imaginary part from the first + or - after the first character that doesn't
      follow an E. Each part is read as a number typed into a cell is (" 4", "$2", "(2)" is -2, "2%", "1,000", "1e3").
      A space may come before that sign but not stand alone after it ("3 +i" is fine, "3+ i" is #NUM!). "3+-4i" is
      3-4i, "++i" and "3+4I" are #NUM!. A number is its real part, an empty cell 0, TRUE #VALUE!.
    - Writing: each part to 15 significant digits, written out in full up to 21 characters (one more than a number
      turned into text gets): "0.0000000000000000001+i", but "1.5E-19+i". A coefficient of 1 is left out ("i", "-i"),
      and a zero part too ("3", "0"). "i" and "j" mixed: IMSUM and IMPRODUCT #VALUE!, IMSUB and IMDIV #NUM!.
    - Past the largest number: IMPRODUCT #VALUE!, the others #NUM!. IMABS is #NUM! once |re| + |im| is
      ("1E308+1E308i"). cosh and sinh are #NUM! once e^x is (IMCOSH("710")), though their value would fit.
    - Powers and roots by the polar form, so IMPOWER("i",2) is -1+1.22514845490862E-16i. The tangent and the
      reciprocals by the double angle (tan(x+iy) = (sin 2x + i sinh 2y) / (cos 2x + cosh 2y)), the real functions'
      own for a real number; ±i and 0 where the double angle is past the largest number (IMTAN("1+400i") is i).
  - `functions-convert.js`: CONVERT, with Excel's 152 unit names in 13 kinds. Each factor is Excel's own, read to
    the last bit from CONVERT(1,unit,base) in a hidden Excel. A conversion is number × factor ÷ factor, which matches
    Excel to the last bit in most cases and to 2 bits in the rest (724 cases). The units that take a prefix were
    measured too: 40 of them, among them mph ("kmph"), ar ("kar" is 1000 ar) and the litre ("cl").
    - A prefix is squared or cubed for a length squared or cubed ("km2" is 1E+06 m2), and only one is allowed ("kmm"
      is #N/A). "da" and "e" are both deka. The binary prefixes (ki, Mi, ... Yi) are for bit and byte only.
    - Names are exact: "M", "FT" and " m" are #N/A, as are units of different kinds. TRUE is #VALUE!, and "2" is 2.
    - Temperatures go through kelvin, with "K" and "kel" taking a prefix ("mK"), and C, F, Rank and Reau not.
  - `functions-oddbonds.js`: ODDFPRICE, ODDFYIELD, ODDLPRICE and ODDLYIELD, by Microsoft's formulas with Excel's own
    counting, measured on 360 cases with cell values (prices to 1e-15, ODDFYIELD's search to 1e-9), then the
    probe's 54:
    - The quasi-coupon dates of an odd period step from the first coupon back, or from the last interest date on, each
      from the one before on the same day, not kept at the month's end: June 30, December 30, June 30; December 31,
      September 30, June 30, March 30. The days to the next coupon (bases 2 and 3) and from the last (the others)
      take COUPNCD's and COUPPCD's dates, which are.
    - A short first period's N counts the first coupon. A long one's whole quasi-periods from settlement to the first
      coupon are counted on from settlement, moved to its month's end when the first coupon is at one (which counts
      one): settlement on February 10, 2020 and a first coupon on June 30 make 1 half-year, not 0.
    - An odd last period's quasi-periods have their actual days for length with bases 1, 2 and 3, 360 over the
      frequency with 0 and 4.
    - Issue on settlement is #NUM!, as are dates out of order, a negative rate, a frequency other than 1, 2, 4.
  - `functions-bessel.js`: BESSELJ, BESSELY, BESSELI and BESSELK. Excel's are not the exact functions but Numerical
    Recipes' approximations, good to about 1e-8: =BESSELJ(0,0) is 1.00000000283141. The rational and polynomial
    forms for orders 0 and 1, then recurrences: up for Y, K, and J past its order; down (Miller's, 40 for the start)
    for I and J. The coefficients past 8 are two printings: Y0 shows 0.934945152e-7 where J0 has 0.934935152e-7,
    and J0's p -0.1098628267e-2 where Y0's is -0.1098628627e-2; J1's p has -0.3516396496e-4 twice, a term more
    than Y1's (so J1(8) is 0.2346283865070744, 8e-6 from the function). Found from 221 points from 8 to 30, to the
    last bit. Measured on about 970 cases with cell values, then the probe's 60; all as Excel's to the last bit or
    within 1e-15. n is made whole, below 0 #NUM!; Y and K need x above 0; J and Y are #NUM! from 2^27 on, as SIN is.
  - `iterate.js`: circular references (a formula that reads itself, or a circle of them), measured with a hidden
    Excel on 30 cases (each with iteration off, on with 100 and on with 10 iterations; entered, F9, Ctrl+Alt+F9) and
    an 11-step sequence in one workbook, all the same here, to the last bit:
    - Iteration off: a formula in a circle isn't worked out; it keeps its value, and the one that closes the circle
      shows 0 (=A1+1 in A1 is 0). The editor says which formulas are circular.
    - Iteration on (calcPr iterate="1"; iterateCount and iterateDelta written only when not 100 and 0.001): every
      recalculation goes round, whatever changed (=A1+1 is 100, then 200 after any edit). Each pass works out the
      formulas in circles and those reading them, row by row of each sheet in order (B1 before A2), reading the last
      pass's values; one with no value yet reads as 0. The passes stop when no number changed by Maximum Change or
      more, the whole workbook's (one circle still changing keeps every circle going), or after Maximum Iterations:
      =1+A1/2 stops at 1.9990234375 after 11 passes.
  - `threed.js`: 3D references (=SUM(Jan:Apr!A1), 'Feb:My Mar'!A1:B2) and the intersection operator, a space
    between references (=A1:C3 B2:D4 is B2:C3, spilled; =SUM(B:B 2:2) is B2). Measured on 86 formulas over four sheets
    in a hidden Excel, all the same here, then the probe's 17:
    - A 3D reference is the area on each sheet from the first to the last in the workbook's order. SUM, COUNT, COUNTA,
      AVERAGE(A), MAX(A), MIN(A), PRODUCT, STDEV, VAR and their kinds, and in Excel 365 also MEDIAN, LARGE, SMALL,
      PERCENTILE, QUARTILE, SUMSQ, AVEDEV, GEOMEAN, DEVSQ, SKEW, KURT, CONCAT, TEXTJOIN, AND, OR, VSTACK and HSTACK
      (sheet after sheet) and TOCOL take one; IF, CHOOSE and IFERROR pass it on; ISREF says FALSE.
    - Functions that want a reference of one sheet (COUNTIF(S), SUMIF(S), ROWS, INDEX, N, AGGREGATE, SUBTOTAL,
      COUNTBLANK) give #VALUE!; any other, and the reference as a value (=Jan:Apr!A1, +1, SUMPRODUCT, MODE), #REF!.
    - The first name reading as a cell makes it a range, as Excel reads it: S1:S4!A1 is S1 to S4!A1.
    - Intersection binds tighter than negation; no overlap is #NULL!; references on two sheets, or a 3D one, #VALUE!.
    - The file has them as typed: `SUM(Jan:Apr!A1)`, `'Feb:My Mar'!A1:B2`, `A1:C3 B2:D4`.
    - Typed, a 3D reference is put in the workbook's order (`Apr:Jan!A1` is `Jan:Apr!A1`) and one of one sheet made
      plain (`Jan:Jan!A1` is `Jan!A1`). As sheets go and move (16 cases measured, all as Excel's): an end sheet that
      leaves the span's other sheets, deleted or moved out, gives way to the next of them inward (Jan:Apr with Jan
      deleted, or moved after Apr, is Feb:Apr; Feb:Mar with Feb gone is Mar!); one moved but still among them, and any
      other sheet, leaves it as it was; a sheet moved or inserted between the ends is read. Renaming an end renames it.
    - The range operator between references (=SUM(A1:INDEX(A:A,3)), =OFFSET(A1,1,0):A4, =INDEX(..):INDEX(..)):
      the range holding both, measured on 25 formulas, all as Excel's, then the probe's 9. Tightest of all
      (=A1:B2 B2:C3:D4 is B2; =-A1:INDEX(A:A,2) spills -A1 and -A2); references on two sheets, or a 3D one, #VALUE!;
      Jan:'My Mar'!A1 is Jan, a name, to 'My Mar'!A1, so #NAME?; a number or text on one side isn't a formula.
      A formula whose range is worked out (INDEX's) is worked out at every change, as the cells between can be any.
    - Negation over a range of several cells spills in a dynamic formula, =-A1:A2 as -A1 and -A2.
    - Rows or columns inserted or deleted, or cells cut and pasted, on one of its sheets leave a 3D reference as it
      was: it reads what is there now (8 cases measured, all as Excel's).
  - `functions-amort.js`: AMORLINC and AMORDEGRC, the French depreciation. Microsoft's page gives other rules than
    Excel's; these are Excel's, from about 600 cases with cell values in a hidden Excel, then the probe's 48:
    - AMORDEGRC's coefficient by the life (1 / rate): up to 2 years #NUM!, to 4 years 1.5, to 6 years 2, past that 2.5.
    - The life has ceil(1 / rate) periods, the first (prorated) one with them. The last two take half of what is left
      each, and the ones after it nothing. A period starting below salvage takes nothing.
    - Each period's depreciation is rounded as returned, but the value it leaves isn't, except the first period's
      (5000 bought January 1 at 0.3: 1119, then 1746, not 1747).
    - The first period's part of a year: bases 0 and 4 as YEARFRAC; bases 1 and 3 actual days, February 29 counted as
      the 28th, over 365, or for basis 1 over the purchase year's days. None (bought on the first period's last day):
      a whole year, and for AMORDEGRC one period fewer.
    - basis 2 is #NUM!, as are a purchase after the first period, salvage above cost, a negative period, a rate of 0.
  - `functions-regex.js`, `regex.js`: REGEXTEST, REGEXEXTRACT and REGEXREPLACE. Excel's patterns are PCRE2's with
    Unicode properties; `regex.js` turns them into JavaScript's (the v flag). Measured on about 190 cases with a hidden
    Excel, then the probe's 125:
    - \w, \d, \s and \b are Unicode's (é is a letter, a no-break space a space, ٣ a digit). $ matches before a final
      newline too, . anything but a newline. (?i), (?s), (?m) and (?x) at the start apply to the whole pattern.
    - Read too: [[:alpha:]] and the other POSIX classes, \Q...\E, (?P<name>...) and (?P=name), \g1 and \g{-1},
      \x{..}, \o{..}, octal, \cJ, \h, \R, \N, a{,2}, comments (?#...). Atomic groups and possessive quantifiers are
      made from a lookahead and a reference to it (^a++a matches nothing in "aaa", as in Excel).
    - An unknown escape (\i), an unclosed group or a bad repeat is #VALUE!. case_sensitivity is 0 or 1, else #VALUE!.
    - REGEXEXTRACT: mode 0 the first match (#N/A if none); 1 every match, in a row; 2 the first match's groups, in a
      row, a group that took no part #N/A, and #VALUE! where there is no match or no group. An array given it is cut
      to its first value. REGEXTEST and REGEXREPLACE work out each value of an array.
    - REGEXREPLACE: occurrence 0 every match, n the nth, -n the nth from the end. The replacement is PCRE2's extended
      one: $1, ${1}, ${name}, $name, $0 and $&, $$; \1; \n, \t, \x41; \u, \l, \U...\E, \L...\E; ${2:-default}
      (where group 2 took no part) and ${2:+set:unset}. A group the pattern doesn't have, "$" or "\" alone, and an
      unknown escape (\q) are #VALUE!.
  - `trig.js`: sine and cosine as Excel works them out. The angle is brought near 0 with the x87's 64-bit π, not the
    true one, so =SIN(PI()) is 1.22514845490862E-16 and =COS(PI()/2) 6.1257422745431E-17 (the true values are
    1.22464679914735E-16 and 6.12323399573677E-17). From 2^27 on, #NUM! (=SIN(134217728)). TAN is SIN over COS
    (=TAN(PI()/2) is 1.63245522776191E+16), COT COS over SIN. SIN, COS, TAN, SEC, CSC, COT and the complex functions
    use it.
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

## Formula Auditing

Show Formulas (Formulas > Formula Auditing, Ctrl+`), as Excel does it (measured with a hidden Excel's
Window.DisplayFormulas, its Range.Text and the file it saved):

- Saved as `showFormulas="1"` on the sheet's view; a sheet opened shows formulas when its file says so.
- Each column is drawn twice as wide (Excel's 54.75 points to 109.5), the width saved unchanged.
- Each cell shows what it holds, unformatted, as the formula bar has it: =A1*2, 1234.5 for 1,234.50, 45292 for a
  date, 0.25 for 25%, 007 for '007, 100000000000000000000 for 1E+20; an array formula without its braces. Formulas sit
  left, as text does, unless the cell is aligned otherwise.

Trace Precedents and Trace Dependents (`ooxml/trace.js`, `editor/tracing.js`, `editor/arrows.js`), as Excel's arrows
go (measured with a hidden Excel's ShowPrecedents, ShowDependents and NavigateArrow on
=A1+A2+SUM(B1:B3)+Sheet2!A1+Rate+A1 and its dependents; `test/trace.test.js`):

- One arrow for each different cell or range read: A1 read twice is one arrow. A defined name is traced to what it names
  (Rate to D1). Everything on other sheets is one arrow of its own, from a worksheet icon (to it, for dependents).
- Each click goes a level further, from the formulas the last level reached, until there is nothing more. A trace from
  another cell starts afresh, the arrows drawn so far staying. Remove Arrows takes them all away; Remove Precedent
  Arrows and Remove Dependent Arrows those of one kind.
- Drawn as Excel draws them: a dot where the arrow starts and a head where it ends, a range boxed with the arrow from
  its top-left cell, blue, red from a cell holding an error, dashed black to or from another sheet. Each sheet keeps
  its own arrows; none is saved (Excel saves none).
- **Not yet**: Excel takes arrows away when a formula they point to changes, rows or columns are inserted or cells
  moved; here they stay until removed. The arrows' exact look (Excel's line widths, its icon) isn't measured. Double-
  clicking an arrow to go to its other end. How Excel aligns numbers in Show Formulas isn't measured (Range.Text can't
  say); here General's right.

Error indicators (`ooxml/errorcheck.js`): Excel's green triangle in the top-left corner of a cell that breaks one of
the rules Excel checks by default, each cell's message heard as its description (measured with a hidden Excel's
Range.Errors on about 60 cells; `test/errorcheck.test.js` has them):

- **Evaluates to an error**: a formula whose result is an error, but #N/A from a formula calling NA() (=NA(), =NA()+1;
  =1/0+NA() is flagged, and so is =A1 reading an =NA()). An error typed as a value isn't.
- **Number stored as text**: text that reads as a number ('123, '12.5, '1,234, '1e3, ' 7, '50%, '(5), '-3, '0012), not
  a date or time ('1/2/2024, '12:30), 'TRUE, or the text a formula gives.
- **Two-digit year**: text that reads as a date with a two-digit year ('1/2/24, 'Jan 2, 24, '2-Jan-24).
- **Inconsistent formula**: the cells either side, left and right or above and below, hold one formula between them
  (moved as a copy moves it) and this cell another (=E3*F3 between =E2+F2 and =E4+F4).
- **Formula omits adjacent cells**: a number (a date too; not text, TRUE or a formula's result) just before or after a
  range of one row or one column the formula reads, on whatever sheet that is (=SUM(N2:N4) with a number in N1). A
  range of two rows and columns isn't checked, nor a number past a gap.
- **Unlocked formula**: a formula in a cell whose format isn't locked.
- A sheet's ignoredErrors (Excel's Ignore Error, saved in the file) turn a rule off for their cells.
- **The error button** beside the active cell when it has an indicator (`editor/errormenu.js`, `ooxml/errorfix.js`;
  Alt+Shift+F10 too): Excel's menu, headed by the error (Number Stored as Text, Divide by Zero Error...), with its fixes
  (Convert to Number; Copy Formula from Above or Left for an inconsistent formula; Lock Cell), Ignore Error and Edit in
  Formula Bar, for the selected cells with that error, each one step of Undo. Ignore Error is written as Excel writes
  it (measured): one ignoredError a rule, its cells joined into ranges, `<ignoredError sqref="A1:A2 A4"
  numberStoredAsText="1"/>`, after pageMargins.
- **Error Checking** (Formulas > Formula Auditing; `editor/errorbox.js`): Excel's box going through the sheet's cells
  with an indicator in reading order (rows, then columns), from the active cell, each selected in turn: "Error in cell
  G2", what it holds, the error's name and what it means, its fix, Ignore Error, Edit in Formula Bar, Previous, Next;
  a fix or Ignore Error goes on to the next. Past the last it starts again at the top, and back where it started says
  "The error check is complete for the entire sheet." The box is a dialog COM can't drive: its order and words are
  Excel's as its documentation and the box itself show them, not measured; Excel's box lets the sheet be clicked
  while it is open, this one doesn't.
- **Not yet**: the two-digit year's Convert XX to 19XX and 20XX, Update Formula to Include Cells, Show Calculation Steps,
  Help and Error Checking Options; whether Convert to Number changes a Text number format (not measured: it is kept).
  The rules Excel has off by default (formulas referring to empty cells). One case Excel didn't flag here does: =SUM(M1:M3)
  in M5 with a number in M4, when M7 held =M1+M2 (alone, Excel flags it). The triangle's green isn't measured.

Evaluate Formula (`calc/stepper.js`, `editor/evalbox.js`; `test/stepper.test.js`): the active cell's formula worked out
a step at a time, the next part underlined and the value just worked out in italics; Evaluate, Restart at the end.
The steps go as Microsoft's own example shows them (=IF(AVERAGE(D2:D5)>50,SUM(E2:E5),0) to =IF(40>50,...), to
=IF(FALSE,...), to 0): a reference on its own or as an operand is a step of its own (its value; a range's as an array);
one handed to a function goes with it; a function or operator once its arguments are values; IF works out its
condition, then only the argument it returns, IFERROR its value_if_error only after an error; brackets go with what
is in them. Excel's box is a dialog COM can't drive, so the steps aren't measured beyond that example.

- **Not yet**: Step In and Step Out (there, not working); a defined name's own steps.

The Watch Window (`ooxml/watches.js`, `editor/watchwindow.js`; `test/watches.test.js`): the watched cells with their Book,
Sheet, Name (a defined name for just that cell), Cell, Value (as the cell shows it) and Formula, kept up to date as
the workbook changes; Add Watch watches the selected cells, each cell of a range on its own as Excel's does; Delete
Watch; a double click or Enter goes to the cell. Saved as Excel saves them (measured: Application.Watches.Add, the
file saved and opened again): a sheet's `<cellWatches><cellWatch r="B2"/>...`, in the order added, after pageMargins.

- **Not yet**: Excel's Add Watch box (here the selection is watched at once); the window floating (it is docked under
  the ribbon here); watches of other open workbooks. Adding or deleting a watch isn't a step of Undo, as in Excel.

## Calculation options

Formulas > Calculation (`src/ooxml/calcmode.js`): Calculation Options (Automatic, Automatic Except for Data Tables,
Manual), Calculate Now (F9), Calculate Sheet (Shift+F9); Ctrl+Alt+F9 works out every formula. Measured in a hidden
Excel (and checked the same way by `test/calcmode.test.js`):

- **Saved** in workbook.xml's calcPr: Manual as `calcMode="manual"`, Automatic Except for Data Tables as
  `calcMode="autoNoTable"`, Automatic as no calcMode; Recalculate workbook before saving off as `calcOnSave="0"`. A
  workbook opened is in the mode its file says.
- **Manual**: a change works out nothing that reads it (=A1*2 stays 10 when A1 goes from 5 to 7, and so does a formula
  on another sheet), and the status bar says Calculate. A formula entered is worked out at once, from what the cells
  hold then (=A2+1 is 11 while A2 still shows 10). Calculate Sheet works out the shown sheet's (A2 14), the other
  sheets' wait; Calculate Now the rest. Back to Automatic, everything waiting is worked out.
- **Saving** a manual workbook works it out first (`save()` in the editor and the agent), unless calcOnSave is off.
- There are no data tables here (What-If Analysis), so Automatic Except for Data Tables works as Automatic.
- **Not yet**: inserting or deleting rows and columns, and moving or copying sheets, still work out their formulas in
  Manual (Excel leaves them waiting); the Options box's iteration and "Recalculate before saving" (calcOnSave can be
  set through `setCalcOnSave`).

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
| YEARFRAC's US 30/360 (basis 0) makes a start on the last day of February the 30th, but then leaves an end on the 31st as the 31st; only a start on the 30th or 31st itself makes it the 30th. | `=YEARFRAC(DATE(2020,2,29),DATE(2020,12,31),0)` is 0.836111 (301 days), `=YEARFRAC(DATE(2021,2,28),DATE(2021,3,31),0)` 0.086111 (31). |
| MOD works on the binary values, without the 15-digit rounding. | `=MOD(0.3,0.1)` is 0.09999999999999998. |
| TIME is the fraction of a day with whole days taken off. | `=TIME(25,0,0)` is 0.04166666666666674. |
| ROUND, ROUNDUP and ROUNDDOWN work on the value's 15 significant decimal digits. | `=ROUND(2.675,2)` is 2.68. |
| Text compares without regard to case, in Excel's order. Numbers come before text, and text before FALSE and TRUE. A blank counts as 0, "" or FALSE, whichever the other value is. | `="a"="A"`, `=1<"a"`, `="a"<TRUE`, `=Data!F2=""` (blank) are all TRUE. |
| Approximate lookups on a list sorted up (MATCH 1, VLOOKUP and HLOOKUP TRUE) are a binary search, which on a list that isn't sorted gives its own answer. From the middle: a blank or a value of another type is passed over to the right; an exact match is taken at once, then the last of the equal values after it. On a list sorted down (MATCH -1) it is a scan from the start, blanks and other types passed over, taking the first exact match and stopping at the first value below. | 400 unsorted lists, some with blanks and text, 1,200 answers (`test/lookup-excel.json`, from `tools/excel-compare/excel-lookups.ps1`): all as Excel's. `MATCH(8,{3,6,9,6,6,1,8,7},1)` is 7, not 2 or 8; `MATCH(6,{5,8,7,1},-1)` is #N/A. |

## Not yet

- **ACCRINT with basis 1 over many coupon periods.** Excel's answer differs from the rule above in the fourth digit
  (ACCRINT(DATE(2005,7,10),DATE(2006,1,31),DATE(2009,6,20),0.06,100,2,1) is 23.6307 in Excel, 23.625 here); its own
  rule for that case isn't found yet. Shorter spans, and every other basis, are Excel's.
- **ODDLPRICE and ODDLYIELD from a last interest date at a month's end into a February 29 settlement** (August 31,
  2019 and February 29, 2020 measured): 1e-4 from Excel's for some frequencies and bases. Kept out of the probe.
- **AMORDEGRC bought on the first period's last day with a life of 3 years or less**: the rule here (declining, no
  halves, up to the life's periods) fits the two such assets measured, not more.
- **Regular expressions JavaScript can't say**: \K, recursion ((?1), (?R)), conditionals, \G, \X, and options partway
  through a pattern ("a(?i)b", "(?i:b)"). A formula using them keeps the value Excel saved. And in a case-insensitive
  pattern, \p{Lu} and [[:upper:]] match lower-case letters here too; in Excel they don't.
- **IMCSC("1+i")**: Excel's last digits there (0.621518017170428-0.303931001628427i) come from neither form tried;
  here 0.621518017170429-0.303931001628426i. Other arguments agree.
- **Spilling** is in docs/DYNAMIC.md (dynamic arrays, LET, LAMBDA, `D1#`, `@`); its own gaps are listed there.
- **A function Excel doesn't know** (=foo(1), #NAME?) is written in capitals here, FOO(1); Excel keeps it as typed.
  (Telling it from a function Excel has but this editor hasn't needs Excel's whole list.)
- **The range operator and intersection in Evaluate Formula** are one step each, as a reference is (=SUM(A1:INDEX(A:A,3))
  goes straight to its sum); Excel's steps there aren't measured.
- **Iterative calculation: what isn't yet as Excel.** No on-screen switch for it (Excel's File > Options > Formulas):
  it is read from the file and kept, and `setIteration` (calcmode.js) turns it on. A circle whose cells change only
  in text or errors isn't measured: here only numbers count toward Maximum Change. Array and dynamic formulas in a
  circle aren't measured either.
- **Dates and times typed as entries** are read as dates and times (docs/ENTRY.md). Inside a formula, text such as
  "2024-01-02" or "1/2/2024" reads as a date.
