// What each Math & Trig and Lookup & Reference function does, one sentence each, for the Function Arguments box
// (describe.js). Written for this project in its own words. One line per function: "NAME: sentence".

export const MATH = `
ABS: Gives a number without its sign.
ACOS: Gives the arccosine of a number, in radians from 0 to pi.
ACOSH: Gives the inverse hyperbolic cosine of a number.
ACOT: Gives the arccotangent of a number, in radians from 0 to pi.
ACOTH: Gives the inverse hyperbolic cotangent of a number.
AGGREGATE: Works out a total such as a sum, average or maximum, with the choice to leave out hidden rows and errors.
ARABIC: Turns a Roman numeral into a number.
ASIN: Gives the arcsine of a number, in radians from -pi/2 to pi/2.
ASINH: Gives the inverse hyperbolic sine of a number.
ATAN: Gives the arctangent of a number, in radians from -pi/2 to pi/2.
ATAN2: Gives the angle of the point (x, y) from the x axis, in radians.
ATANH: Gives the inverse hyperbolic tangent of a number.
BASE: Writes a number in another base (radix) as text.
CEILING: Rounds a number up to the nearest multiple of a significance.
CEILING.MATH: Rounds a number up to the nearest whole number or multiple, with a choice for negative numbers.
CEILING.PRECISE: Rounds a number up to the nearest whole number or multiple, whatever its sign.
COMBIN: Counts the ways to choose a number of items from a set, order not counting.
COMBINA: Counts the ways to choose a number of items from a set when items may repeat.
COS: Gives the cosine of an angle in radians.
COSH: Gives the hyperbolic cosine of a number.
COT: Gives the cotangent of an angle in radians.
COTH: Gives the hyperbolic cotangent of a number.
CSC: Gives the cosecant of an angle in radians.
CSCH: Gives the hyperbolic cosecant of a number.
DECIMAL: Turns text that writes a number in another base into a decimal number.
DEGREES: Turns radians into degrees.
EVEN: Rounds a number away from zero to the nearest even whole number.
EXP: Gives e raised to the power of a number.
FACT: Gives the factorial of a number.
FACTDOUBLE: Gives the double factorial of a number.
FLOOR: Rounds a number down to the nearest multiple of a significance.
FLOOR.MATH: Rounds a number down to the nearest whole number or multiple, with a choice for negative numbers.
FLOOR.PRECISE: Rounds a number down to the nearest whole number or multiple, whatever its sign.
GCD: Gives the greatest common divisor of whole numbers.
INT: Rounds a number down to the whole number below it.
ISO.CEILING: Rounds a number up to the nearest whole number or multiple, whatever its sign.
LCM: Gives the least common multiple of whole numbers.
LN: Gives the natural logarithm of a number.
LOG: Gives the logarithm of a number in a base you choose.
LOG10: Gives the base-10 logarithm of a number.
MDETERM: Gives the determinant of a square matrix.
MINVERSE: Gives the inverse of a square matrix.
MMULT: Gives the matrix product of two arrays.
MOD: Gives the remainder after a number is divided by a divisor.
MROUND: Rounds a number to the nearest multiple you choose.
MULTINOMIAL: Gives the multinomial coefficient of a set of numbers.
MUNIT: Gives the identity matrix of a size you choose.
ODD: Rounds a number away from zero to the nearest odd whole number.
PERCENTOF: Gives the share that a subset of values is of all of them.
PI: Gives the value of pi, 3.14159265358979.
POWER: Raises a number to a power.
PRODUCT: Multiplies its arguments together.
QUOTIENT: Gives the whole-number part of a division.
RADIANS: Turns degrees into radians.
RAND: Gives a random number from 0 up to 1, new at each calculation.
RANDARRAY: Gives an array of random numbers.
RANDBETWEEN: Gives a random whole number between two numbers you choose.
ROMAN: Writes a number as a Roman numeral, as text.
ROUND: Rounds a number to a number of digits.
ROUNDDOWN: Rounds a number toward zero.
ROUNDUP: Rounds a number away from zero.
SEC: Gives the secant of an angle in radians.
SECH: Gives the hyperbolic secant of a number.
SEQUENCE: Gives an array of numbers in sequence.
SERIESSUM: Adds up a power series.
SIGN: Gives 1 for a positive number, 0 for zero and -1 for a negative number.
SIN: Gives the sine of an angle in radians.
SINH: Gives the hyperbolic sine of a number.
SQRT: Gives the square root of a number.
SQRTPI: Gives the square root of a number times pi.
SUBTOTAL: Works out a subtotal of a list, leaving out other subtotals in it.
SUM: Adds up all the numbers in its arguments.
SUMIF: Adds up the cells that meet a condition.
SUMIFS: Adds up the cells that meet every one of a set of conditions.
SUMPRODUCT: Multiplies matching items of arrays and adds up the products.
SUMSQ: Adds up the squares of its arguments.
SUMX2MY2: Adds up the differences of the squares of matching values in two arrays.
SUMX2PY2: Adds up the sums of the squares of matching values in two arrays.
SUMXMY2: Adds up the squares of the differences of matching values in two arrays.
TAN: Gives the tangent of an angle in radians.
TANH: Gives the hyperbolic tangent of a number.
TRUNC: Cuts a number to a whole number, or to a number of digits, without rounding.
`;

export const LOOKUP = `
ADDRESS: Writes a cell's reference as text from a row and a column number.
AREAS: Counts the areas in a reference.
CHOOSE: Picks one value from a list by its position.
CHOOSECOLS: Gives the columns of an array you name by number.
CHOOSEROWS: Gives the rows of an array you name by number.
COLUMN: Gives the column number of a reference.
COLUMNS: Counts the columns in an array or reference.
DROP: Gives an array without a number of rows or columns at its start or end.
EXPAND: Grows an array to a number of rows and columns, filling the new places.
FILTER: Gives the rows or columns of an array that meet a condition.
FORMULATEXT: Gives the formula of a cell as text.
GETPIVOTDATA: Gives a value from a PivotTable.
GROUPBY: Groups the rows of a table by fields and works out totals for each group.
HLOOKUP: Looks for a value in the top row of a table and gives the value in a row you name below it.
HSTACK: Puts arrays side by side in one array.
HYPERLINK: Makes a link that opens a file, a web page or a place in the workbook.
INDEX: Gives the value or reference at a row and column of an array or range.
INDIRECT: Gives the reference a text names.
LOOKUP: Looks for a value in a row or column and gives the value in the same place of another.
MATCH: Gives the position of a value in a row or column.
OFFSET: Gives a reference a number of rows and columns away from another, of the size you choose.
PIVOTBY: Groups the rows of a table by row and column fields and works out totals for each group.
ROW: Gives the row number of a reference.
ROWS: Counts the rows in an array or reference.
SORT: Sorts an array or range.
SORTBY: Sorts an array or range by the values of other arrays.
TAKE: Gives a number of rows or columns from the start or end of an array.
TOCOL: Puts an array into a single column.
TOROW: Puts an array into a single row.
TRANSPOSE: Turns the rows of an array into columns and the columns into rows.
TRIMRANGE: Gives a range without the empty rows and columns at its edges.
UNIQUE: Gives the different values of an array, each once.
VLOOKUP: Looks for a value in the first column of a table and gives the value in a column you name to its right.
VSTACK: Puts arrays one under another in one array.
WRAPCOLS: Wraps a row or column of values into columns of a length you choose.
WRAPROWS: Wraps a row or column of values into rows of a length you choose.
XLOOKUP: Looks for a value in an array and gives the matching item of another, with choices for what isn't found.
XMATCH: Gives the position of a value in an array, with choices for how to match and search.
`;
