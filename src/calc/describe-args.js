// What each argument means, one sentence each, for the Function Arguments box (describe.js): by the argument's name
// (Excel's, from calc/catalog.js, without brackets or the number of a repeated one: "number1" is "number"), and for
// a function where the name means something of its own, by "FUNCTION.name". Written for this project in its own words.
// One line per argument: "name: sentence".

export const BY_NAME = `
A: The lower end of the interval of x, 0 if left out.
B: The upper end of the interval of x, 1 if left out.
a: The leading coefficient of the series.
abs_num: The kind of reference: 1 absolute, 2 absolute row, 3 absolute column, 4 relative.
actual_range: The observed values.
alpha: A parameter of the distribution.
angle: An angle.
argument: The LAMBDA parameter to check.
array: An array or a range of cells.
array_x: The first array or range of values.
array_y: The second array or range of values.
average_range: The cells to average; the range itself if left out.
base: The base of the logarithm.
basis: The day count basis: 0 or left out US 30/360, 1 actual/actual, 2 actual/360, 3 actual/365, 4 European 30/360.
beta: A parameter of the distribution.
bins_array: The upper limits of the bins to count values into.
bottom: The smallest whole number it may give.
by_array: The array or range to sort by.
by_col: TRUE to work by columns; FALSE or left out to work by rows.
calc_method: TRUE or left out to count interest from issue; FALSE to count it from the first interest date.
calculation: The formula to work out, using the names or parameters.
calculation_or_name: The formula to work out, or the next name to define.
case_sensitivity: 0 or left out for letter case to count; 1 for it not to.
coefficients: The coefficients of the successive powers of x.
col_delimiter: The text that marks where to split into columns.
col_fields: The values to group by in columns.
col_index_num: The column of the table to take the value from, 1 for its first column.
col_num: The number of a column to take.
col_sort_order: How to sort the column groups: a column number, negative for descending.
col_total_depth: Whether and how column totals are shown.
cols: The number of columns.
column_num: The column number.
columns: The number of columns.
const: TRUE or left out to work out the constant b as usual; FALSE to set it to 0 or 1.
cost: The asset's starting cost.
coupon: The security's yearly coupon rate.
criteria: The condition a cell must meet: a number, text, an expression such as ">5", or a reference.
criteria_range: The range to check against the condition after it.
cumulative: TRUE for the cumulative distribution; FALSE for the probability density or mass.
data_all: All the values.
data_array: The values whose frequencies are counted.
data_field: The name of the data field to take, in quotes.
data_subset: The values that are part of all of them.
database: The list or table of records, its headings in the first row.
date: A date.
date_purchased: The date the asset was bought.
date_text: Text that shows a date.
dates: The date of each cash flow, the first being the start.
day: The day of the month.
days: The number of working days before or after the start date.
decimal_dollar: A decimal number.
decimal_separator: The character that separates whole numbers from decimals.
decimals: The number of digits after the decimal point.
default_or_value: A result if nothing matches, or the next value to compare.
deg_freedom: The degrees of freedom.
delimiter: The text that marks the place.
denominator: The number to divide by.
dimension: The number of rows and columns of the identity matrix.
discount: The discount rate.
divisor: The number to divide by.
effect_rate: The effective interest rate.
end_date: The last date.
end_period: The last period.
error_val: The error value.
exactly_once: TRUE for only the values that occur once; FALSE or left out for every different value.
expected_range: The expected values.
expression: The value to compare.
factor: The rate the balance declines at, 2 (double declining) if left out.
field: The column used: its heading in quotes, or its position.
field_headers: Whether the fields have headings and whether to show them.
field_relationship: How the fields given relate to each other.
filter_array: Which rows to take in.
finance_rate: The interest rate paid on the money borrowed.
find_text: The text to find.
first_coupon: The first coupon date.
first_interest: The first date interest is paid.
first_period: The date the first period ends.
form: The style of Roman numeral, 0 or left out for classic.
format: 0 or left out for a short form that is easy to read; 1 for a form that can be typed back in.
format_text: A number format as text, such as "0.00".
fraction: The whole number to use as the denominator of the fraction.
fractional_dollar: A number written as a whole number and a fraction.
frequency: The number of coupon payments a year: 1 yearly, 2 half-yearly, 4 quarterly.
friendly_name: The text to show in the cell; the link itself if left out.
from_unit: The unit the number is in.
function: The function to work out each total with, such as SUM.
function_num: A number that picks the function to use.
fv: The value wanted after the last payment, 0 if left out.
group_separator: The character that groups thousands.
guess: A first guess at the answer.
height: The number of rows of the reference given.
holidays: Dates to leave out as holidays.
hour: The hour, from 0 to 23.
i_num: The imaginary part.
if_empty: The value to give when nothing is taken in.
if_not_found: The value to give when nothing is found.
ignore: Which values to leave out: 0 none, 1 blanks, 2 errors, 3 both.
ignore_empty: TRUE to leave empty texts out.
include: TRUE or FALSE for each row or column, the ones to keep.
index_num: Which value to pick, from 1.
info_type: The kind of information wanted, as text.
initial_value: The value to start from.
instance_num: Which occurrence to take, from 1.
inumber: A complex number, as text such as "3+4i".
investment: The amount invested.
issue: The security's issue date.
item: An item of the field before it.
k: The position, or the percentile, wanted.
known_x's: The known x values.
known_y's: The known y values.
lambda: A LAMBDA to apply.
lambda_or_array: A LAMBDA to apply, or another array.
last_interest: The last coupon date before maturity.
life: The number of periods of the asset's useful life.
link_location: The address of the file, web page or place to open, as text.
logical: A condition that is TRUE or FALSE.
logical_test: A value or expression that is TRUE or FALSE.
lookup_array: The range or array to look in.
lookup_value: The value to look for.
lookup_vector: The row or column to look in.
lower_limit: The lower limit.
m: The step by which the power increases.
match_end: 1 to treat the end of the text as a delimiter.
match_mode: How to match: 0 or left out exact, -1 or 1 the next smaller or larger, 2 wildcards.
match_type: 1 or left out the largest value not above it, 0 an exact match, -1 the smallest value not below it.
max: The largest number it may give.
max_range: The cells to take the largest from.
maturity: The security's maturity date, when it expires.
mean: The mean of the distribution.
method: FALSE or left out for the US method; TRUE for the European method.
min: The smallest number it may give.
min_length: The fewest characters to give, zeros put in front.
min_range: The cells to take the smallest from.
minute: The minute, from 0 to 59.
mode: How to round negative numbers, toward or away from zero.
month: The month, from 1 to 12.
months: The number of months before (negative) or after the start date.
multiple: The multiple to round to.
n: The order of the function, or the power to start at.
name: A name to give to a value.
name_value: The value given to the name before it.
new_text: The text to put in.
new_x's: The x values to work out y values for.
no_commas: TRUE for no thousands separators.
no_switch: TRUE not to switch to straight-line depreciation.
nominal_rate: The nominal interest rate.
nper: The total number of payments.
npery: The number of compounding periods a year.
num_bytes: The number of bytes.
num_chars: The number of characters.
num_digits: The number of digits to round to.
number: A number, or a reference or range of numbers.
number_chosen: The number of items in each choice.
number_f: The number of failures.
number_pop: The size of the population.
number_s: The number of successes.
number_sample: The size of the sample.
number_times: How many times to repeat it.
numerator: The number to divide.
occurrence: Which match to replace; 0 or left out for every one.
old_text: The text to change.
options: Which values to leave out: hidden rows, errors, nested subtotals.
order: 0 or left out to rank from the largest; any other number from the smallest.
pad_with: The value to fill the new places with.
par: The security's face value.
parameter: A name for a value handed to the function.
pattern: The regular expression.
per: The period, from 1 to nper.
percent: The share of values to leave out, from 0 to 1.
period: The period to work out.
pivot_table: A reference to a cell of the PivotTable.
places: The number of characters to give.
pmt: The payment each period.
population_s: The number of successes in the population.
power: The power to raise it to.
pr: The security's price per 100 of face value.
principal: The present value.
prob_range: The probability of each value.
probability: A probability, from 0 to 1.
probability_s: The probability of success in each trial.
pv: The present value, what a series of future payments is worth now.
quart: Which quartile: 0 to 4.
radix: The base, from 2 to 36.
range: The cells to check.
range_lookup: TRUE or left out for the closest match in a sorted column; FALSE for an exact match.
rate: The interest rate per period.
real_num: The real part.
redemption: The security's redemption value per 100 of face value.
ref: A reference.
ref_text: Text that names a reference.
reference: A cell or range of cells.
reinvest_rate: The interest rate earned on the money reinvested.
relative_to: What a share is worked out relative to.
replacement: The text to put in place of each match.
result: The result for the value before it.
result_vector: The row or column to take the result from.
return_array: The array or range to take the result from.
return_mode: 0 for the first match, 1 for every match, 2 for the parts of the first match.
return_type: A number for the numbering to use.
row_delimiter: The text that marks where to split into rows.
row_fields: The values to group by in rows.
row_index_num: The row of the table to take the value from, 1 for its first row.
row_num: The row number.
row_sort_order: How to sort the row groups: a column number, negative for descending.
row_total_depth: Whether and how row totals are shown.
rows: The number of rows.
salvage: The asset's value at the end of its life.
sample_s: The number of successes in the sample.
scan_by_column: FALSE or left out to read by rows; TRUE to read by columns.
schedule: The interest rates to apply.
search_mode: 1 or left out to search from the first; -1 from the last; 2 or -2 a binary search of sorted values.
second: The second, from 0 to 59.
serial_number: A date or time.
settlement: The security's settlement date, when it is bought.
sheet_text: The name of the sheet, as text.
shift_amount: The number of bits to shift by.
sigma: The population's standard deviation; the sample's if left out.
significance: The multiple to round to.
size: The size of the sample.
sort_index: The row or column to sort by.
sort_order: 1 or left out ascending; -1 descending.
standard_dev: The standard deviation of the distribution.
start: The first number.
start_date: The first date.
start_num: The character to start at, from 1.
start_period: The first period.
stats: TRUE to give more statistics; FALSE or left out for the coefficients only.
step: The amount to step by.
suffix: "i" or "j", the letter for the imaginary part.
sum_range: The cells to add up; the range itself if left out.
table_array: The table to look in.
tails: 1 for a one-tailed distribution, 2 for two-tailed.
text: A text, or a reference to a cell holding one.
time_text: Text that shows a time.
to_unit: The unit to convert to.
top: The largest whole number it may give.
total_depth: Whether and how totals are shown.
trials: The number of independent trials.
trim_cols: Which empty columns to trim.
trim_rows: Which empty rows to trim.
type: 0 or left out for payments at the end of each period; 1 for the start.
type_text: The kind of information wanted, as text.
upper_limit: The upper limit.
value: A value, a reference or an expression.
value_if_error: The value to give if it is an error.
value_if_false: The value to give if the test is FALSE.
value_if_na: The value to give if it is #N/A.
value_if_true: The value to give if the test is TRUE.
values: The cash flows: payments negative, income positive.
vector: A row or column of values.
weekend: The weekend days, as a number or a seven-character text of 1s and 0s.
whole_number: TRUE for whole numbers; FALSE or left out for decimals.
width: The number of columns of the reference given.
within_text: The text to look in.
wrap_count: The number of values in each row or column.
x: The value to work it out at.
x_num: The x coordinate of the point.
x_range: The values.
xml: The XML, as text.
xpath: The XPath that selects the data, as text.
y: The value to work it out at.
y_num: The y coordinate of the point.
year: The year.
yld: The security's yearly yield.
z: The value to work it out at.
`;

/** Arguments that mean something of their own in one function. */
export const BY_FUNCTION = `
AGGREGATE.ref: The values to work out the total of.
CELL.reference: The cell to tell about; the last one changed if left out.
CHOOSE.value: A value to pick from.
COLUMN.reference: The cell or range; the cell of the formula if left out.
COUNT.value: Values or ranges whose numbers are counted.
COUNTA.value: Values or ranges whose values are counted.
COUNTIF.range: The cells to count.
DCOUNT.field: The column used: its heading in quotes, or its position; all records if left out.
DCOUNTA.field: The column used: its heading in quotes, or its position; all records if left out.
DELTA.number: A number to compare; the second is 0 if left out.
ERF.lower_limit: The lower limit; from 0 to it if no upper limit is given.
FREQUENCY.data_array: The values whose frequencies are counted.
GESTEP.step: The step value, 0 if left out.
IFERROR.value: The value or expression to check for an error.
IFNA.value: The value or expression to check for #N/A.
INDEX.array: A range of cells or an array.
INDIRECT.a1: TRUE or left out for an A1 reference; FALSE for an R1C1 one.
ADDRESS.a1: TRUE or left out for an A1 reference; FALSE for an R1C1 one.
ISOMITTED.argument: The LAMBDA parameter to check.
LOG.base: The base of the logarithm, 10 if left out.
MATCH.lookup_array: The row or column to look in.
MOD.number: The number to divide.
NPV.value: A cash flow, at the end of a period.
OFFSET.reference: The cell or range to count from.
OFFSET.rows: The number of rows down (up, negative) from it.
OFFSET.cols: The number of columns right (left, negative) from it.
POWER.number: The number to raise.
RANK.ref: The list of numbers to rank it in.
RANK.AVG.ref: The list of numbers to rank it in.
RANK.EQ.ref: The list of numbers to rank it in.
RANK.number: The number to rank.
RANK.AVG.number: The number to rank.
RANK.EQ.number: The number to rank.
ROUND.number: The number to round.
ROUNDDOWN.number: The number to round.
ROUNDUP.number: The number to round.
ROW.reference: The cell or range; the cell of the formula if left out.
SHEET.value: A sheet's name or a reference; the formula's sheet if left out.
SHEETS.reference: A reference; the whole workbook if left out.
SUBTOTAL.ref: The ranges to work out the subtotal of.
SUM.number: Numbers, or ranges of them, to add up.
TRUNC.num_digits: The number of digits to keep, 0 if left out.
T.TEST.type: The kind of test: 1 paired, 2 equal variances, 3 unequal variances.
TTEST.type: The kind of test: 1 paired, 2 equal variances, 3 unequal variances.
TDIST.tails: 1 for one tail, 2 for two.
TYPE.value: Any value.
WEEKDAY.return_type: 1 or left out Sunday is 1; 2 Monday is 1; 3 Monday is 0.
WEEKNUM.return_type: 1 or left out weeks start on Sunday; 2 on Monday.
`;
