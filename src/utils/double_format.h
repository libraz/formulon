//
// Locale-independent shortest-form formatter for `double` matching Mac
// Excel 365's General-format rendering used by VALUETOTEXT / coerce-to-text.
//
// `format_double` exists so multiple subsystems (the AST S-expression dumper,
// the tree-walk evaluator's `coerce_to_text`, the CLI JSON renderer, IM*
// complex-number formatting) emit
// numeric strings identically. The contract:
//
//   * NaN  -> "nan"
//   * +inf -> "inf",  -inf -> "-inf"
//   * Negative zero collapses to "0".
//   * Excel General format: 15 sig digits of `excel_decimal_view`;
//     decimal form when the unsigned-magnitude width is <= 20 chars,
//     otherwise scientific in Excel style (`E+NN` / `E-NN`, exponent
//     zero-padded to two digits, three+ digits when |exp| >= 100). A leading
//     `-` sign is NOT counted against the 20-char ceiling.
//
// Per-case examples (from `tests/oracle/cases/valuetotext_general_threshold_*`):
//
//        7.123456E-9 ->  "0.000000007123456"   (17 chars,  decimal)
//        1E-18       ->  "0.000000000000000001"(20 chars,  decimal)
//        1E-19       ->  "1E-19"               (sci, decimal would be 21 chars)
//        1E+19       ->  "10000000000000000000"(20 chars,  decimal)
//        1E+20       ->  "1E+20"               (sci, decimal would be 21 chars)
//
// Implementation is dependency-free (`<cmath>`, `<string>`, `number_text.h`).

#ifndef FORMULON_UTILS_DOUBLE_FORMAT_H_
#define FORMULON_UTILS_DOUBLE_FORMAT_H_

#include <string>

namespace formulon {

/// Snaps `v` to the decimal Excel shows with `significant_digits` (15 or 14)
/// digits: 15 digits round to nearest with an exact tie toward zero (Mac Excel
/// shows 100000000000001.5 as 100000000000001), and 14 digits round that
/// 15-digit decimal half up.
double excel_decimal_view(double v, int significant_digits) noexcept;

/// Appends a locale-independent textual form of `v` to `out`.
///
/// `out` is left untouched on entry; the formatter only appends. See the
/// header comment for the per-case formatting rules. `significant_digits`
/// caps the mantissa precision (Excel's General text uses 15).
void format_double(std::string& out, double v, int significant_digits = 15);

}  // namespace formulon

#endif  // FORMULON_UTILS_DOUBLE_FORMAT_H_
