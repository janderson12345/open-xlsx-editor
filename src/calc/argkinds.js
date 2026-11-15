// The kind of value each argument takes, for the Function Arguments box: the hint shown beside an empty box (Excel's
// box shows "= number", "= text", "= logical", "= reference", "= array" or "= any" there) and whether text typed in it
// is put in quotes (text and any). Excel keeps its own list of each argument's kind in no file or COM property, so
// these are given by what each argument means (calc/describe-args.js): by the argument's name (without the number of a
// repeated one), and by "FUNCTION.name" where a function's argument takes another kind. Any other is a number.

const KINDS = {
  text: `text find_text within_text old_text new_text format_text date_text time_text delimiter col_delimiter
    row_delimiter info_type type_text ref_text sheet_text link_location friendly_name from_unit to_unit decimal_separator
    group_separator pattern replacement suffix inumber xml xpath data_field`,
  any: `value lookup_value criteria value_if_true value_if_false value_if_error value_if_na if_not_found if_empty pad_with
    expression result default_or_value initial_value name_value calculation calculation_or_name field item error_val
    lambda lambda_or_array argument parameter name`,
  logical: `logical logical_test cumulative range_lookup ignore_empty by_col exactly_once no_commas no_switch stats const
    whole_number calc_method method scan_by_column a1`,
  reference: `reference ref range criteria_range sum_range average_range max_range min_range pivot_table database`,
  array: `array table_array lookup_array return_array lookup_vector result_vector known_x's known_y's new_x's values
    data_array bins_array actual_range expected_range x_range prob_range array_x array_y by_array vector include
    holidays schedule data_subset data_all row_fields col_fields filter_array coefficients dates`,
};
/** Arguments that take another kind in one function. */
const OWN = {
  'DAVERAGE.criteria': 'reference', 'DCOUNT.criteria': 'reference', 'DCOUNTA.criteria': 'reference', 'DGET.criteria': 'reference',
  'DMAX.criteria': 'reference', 'DMIN.criteria': 'reference', 'DPRODUCT.criteria': 'reference', 'DSTDEV.criteria': 'reference',
  'DSTDEVP.criteria': 'reference', 'DSUM.criteria': 'reference', 'DVAR.criteria': 'reference', 'DVARP.criteria': 'reference',
  'COUNT.value': 'any', 'COUNTA.value': 'any', 'TYPE.value': 'any', 'N.value': 'any', 'T.value': 'any',
  'TEXT.value': 'any', 'ISFORMULA.reference': 'reference', 'BIN2DEC.number': 'text', 'HEX2DEC.number': 'text',
  'OCT2DEC.number': 'text', 'BIN2HEX.number': 'text', 'BIN2OCT.number': 'text', 'HEX2BIN.number': 'text',
  'HEX2OCT.number': 'text', 'OCT2BIN.number': 'text', 'OCT2HEX.number': 'text', 'ARABIC.text': 'text', 'DECIMAL.text': 'text',
};
const BY_NAME = new Map(Object.entries(KINDS).flatMap(([kind, names]) => names.trim().split(/\s+/).map((n) => [n, kind])));

/** The kind of argument `arg` (as the box names it: "text2", "value_if_true") of function `name`. */
export function argKind(name, arg) {
  const base = String(arg).replace(/\d+$/, '');
  const fn = String(name).toUpperCase();
  return OWN[`${fn}.${base}`] || OWN[`${fn}.${arg}`] || BY_NAME.get(arg) || BY_NAME.get(base) || 'number';
}
