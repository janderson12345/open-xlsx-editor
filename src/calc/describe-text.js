// What each Text, Logical, Information, Engineering and Web function does, one sentence each, for the Function
// Arguments box (describe.js). Written for this project in its own words. One line per function: "NAME: sentence".

export const TEXT = `
ARRAYTOTEXT: Writes the values of an array as one text.
ASC: Turns full-width (double-byte) characters into half-width (single-byte) ones.
BAHTTEXT: Writes a number in Thai words, with the baht currency.
CHAR: Gives the character a code number stands for.
CLEAN: Takes the characters that can't be printed out of a text.
CODE: Gives the code number of the first character of a text.
CONCAT: Joins texts, and the values of ranges, into one text.
CONCATENATE: Joins texts into one text.
DBCS: Turns half-width (single-byte) characters into full-width (double-byte) ones.
DOLLAR: Writes a number as text in a currency format.
EXACT: Says whether two texts are exactly the same, letter case counting.
FIND: Gives where one text starts within another, letter case counting.
FINDB: Gives where one text starts within another, counted in bytes.
FIXED: Writes a number as text rounded to a number of decimals.
LEFT: Gives the first characters of a text.
LEFTB: Gives the first characters of a text, counted in bytes.
LEN: Counts the characters in a text.
LENB: Counts the bytes of a text.
LOWER: Turns a text into lower case.
MID: Gives the characters from the middle of a text, from a position on.
MIDB: Gives the characters from the middle of a text, counted in bytes.
NUMBERVALUE: Turns text into a number, with the decimal and group separators you choose.
PHONETIC: Gives the phonetic (furigana) characters of a text.
PROPER: Gives each word of a text a capital first letter.
REGEXEXTRACT: Gives the parts of a text that match a regular expression.
REGEXREPLACE: Replaces the parts of a text that match a regular expression.
REGEXTEST: Says whether any part of a text matches a regular expression.
REPLACE: Replaces part of a text, from a position, with another text.
REPLACEB: Replaces part of a text, counted in bytes, with another text.
REPT: Repeats a text a number of times.
RIGHT: Gives the last characters of a text.
RIGHTB: Gives the last characters of a text, counted in bytes.
SEARCH: Gives where one text starts within another, letter case not counting.
SEARCHB: Gives where one text starts within another, counted in bytes.
SUBSTITUTE: Puts a new text in place of an old one within a text.
T: Gives the text of a value, or empty text for anything else.
TEXT: Writes a value as text in a number format you give.
TEXTAFTER: Gives the part of a text after a delimiter.
TEXTBEFORE: Gives the part of a text before a delimiter.
TEXTJOIN: Joins texts with a delimiter between them.
TEXTSPLIT: Splits a text into rows and columns at delimiters.
TRIM: Takes the extra spaces out of a text.
UNICHAR: Gives the Unicode character a number stands for.
UNICODE: Gives the Unicode number of the first character of a text.
UPPER: Turns a text into upper case.
VALUE: Turns text that shows a number into the number.
VALUETOTEXT: Writes any value as text.
`;

export const LOGICAL = `
AND: Gives TRUE if all its arguments are TRUE.
BYCOL: Applies a LAMBDA to each column of an array.
BYROW: Applies a LAMBDA to each row of an array.
FALSE: Gives the logical value FALSE.
IF: Checks a condition and gives one value if it is TRUE and another if it is FALSE.
IFERROR: Gives a value of your own if an expression is an error, else the expression's value.
IFNA: Gives a value of your own if an expression is #N/A, else the expression's value.
IFS: Checks conditions in turn and gives the value of the first one that is TRUE.
LAMBDA: Makes a function of your own from parameters and a calculation.
LET: Gives names to values in a formula, then works out a calculation with them.
MAKEARRAY: Makes an array of a size you choose, each value worked out by a LAMBDA.
MAP: Works out a new array by applying a LAMBDA to each value of arrays.
NOT: Turns TRUE into FALSE and FALSE into TRUE.
OR: Gives TRUE if any of its arguments is TRUE.
REDUCE: Brings an array down to one value by applying a LAMBDA to each value in turn.
SCAN: Applies a LAMBDA to each value of an array in turn and gives each step's result.
SWITCH: Compares an expression with a list of values and gives the result of the first that matches.
TRUE: Gives the logical value TRUE.
XOR: Gives TRUE if an odd number of its arguments are TRUE.
`;

export const INFORMATION = `
CELL: Gives information about a cell's formatting, place or contents.
ERROR.TYPE: Gives the number of an error value.
INFO: Gives information about the working environment.
ISBLANK: Says whether a value is an empty cell.
ISERR: Says whether a value is an error other than #N/A.
ISERROR: Says whether a value is any error.
ISEVEN: Says whether a number is even.
ISFORMULA: Says whether a cell holds a formula.
ISLOGICAL: Says whether a value is TRUE or FALSE.
ISNA: Says whether a value is the #N/A error.
ISNONTEXT: Says whether a value is anything but text.
ISNUMBER: Says whether a value is a number.
ISODD: Says whether a number is odd.
ISOMITTED: Says whether a LAMBDA's parameter was left out.
ISREF: Says whether a value is a reference.
ISTEXT: Says whether a value is text.
N: Gives a value as a number.
NA: Gives the #N/A error, "no value available".
SHEET: Gives the number of a sheet.
SHEETS: Counts the sheets in a reference, or in the workbook.
TYPE: Gives a number for the kind of a value.
`;

export const ENGINEERING = `
BESSELI: Gives the modified Bessel function In(x).
BESSELJ: Gives the Bessel function Jn(x).
BESSELK: Gives the modified Bessel function Kn(x).
BESSELY: Gives the Bessel function Yn(x).
BIN2DEC: Turns a binary number into a decimal one.
BIN2HEX: Turns a binary number into a hexadecimal one.
BIN2OCT: Turns a binary number into an octal one.
BITAND: Gives the bitwise AND of two numbers.
BITLSHIFT: Shifts a number's bits to the left.
BITOR: Gives the bitwise OR of two numbers.
BITRSHIFT: Shifts a number's bits to the right.
BITXOR: Gives the bitwise exclusive OR of two numbers.
COMPLEX: Makes a complex number from a real and an imaginary part.
CONVERT: Converts a number from one unit of measure to another.
DEC2BIN: Turns a decimal number into a binary one.
DEC2HEX: Turns a decimal number into a hexadecimal one.
DEC2OCT: Turns a decimal number into an octal one.
DELTA: Says (1 or 0) whether two numbers are equal.
ERF: Gives the error function between two limits.
ERF.PRECISE: Gives the error function from 0 to a limit.
ERFC: Gives the complementary error function.
ERFC.PRECISE: Gives the complementary error function.
GESTEP: Says (1 or 0) whether a number is at least a step value.
HEX2BIN: Turns a hexadecimal number into a binary one.
HEX2DEC: Turns a hexadecimal number into a decimal one.
HEX2OCT: Turns a hexadecimal number into an octal one.
IMABS: Gives the absolute value (modulus) of a complex number.
IMAGINARY: Gives the imaginary part of a complex number.
IMARGUMENT: Gives the argument of a complex number, an angle in radians.
IMCONJUGATE: Gives the complex conjugate of a complex number.
IMCOS: Gives the cosine of a complex number.
IMCOSH: Gives the hyperbolic cosine of a complex number.
IMCOT: Gives the cotangent of a complex number.
IMCSC: Gives the cosecant of a complex number.
IMCSCH: Gives the hyperbolic cosecant of a complex number.
IMDIV: Divides one complex number by another.
IMEXP: Gives e raised to a complex number.
IMLN: Gives the natural logarithm of a complex number.
IMLOG10: Gives the base-10 logarithm of a complex number.
IMLOG2: Gives the base-2 logarithm of a complex number.
IMPOWER: Raises a complex number to a power.
IMPRODUCT: Multiplies complex numbers.
IMREAL: Gives the real part of a complex number.
IMSEC: Gives the secant of a complex number.
IMSECH: Gives the hyperbolic secant of a complex number.
IMSIN: Gives the sine of a complex number.
IMSINH: Gives the hyperbolic sine of a complex number.
IMSQRT: Gives the square root of a complex number.
IMSUB: Subtracts one complex number from another.
IMSUM: Adds complex numbers.
IMTAN: Gives the tangent of a complex number.
OCT2BIN: Turns an octal number into a binary one.
OCT2DEC: Turns an octal number into a decimal one.
OCT2HEX: Turns an octal number into a hexadecimal one.
`;

export const WEB = `
ENCODEURL: Writes a text so that it can be used in a web address.
FILTERXML: Gives the data an XPath selects from XML text.
`;
