// Copyright 2026 libraz. Licensed under the MIT License.
//
// Locale-independent shortest-form formatter for `double` matching Mac
// Excel 365's General-format rendering used by VALUETOTEXT / coerce-to-text.
//
// `format_double` exists so multiple subsystems (the AST S-expression dumper,
// the tree-walk evaluator's `coerce_to_text`, the CLI JSON renderer, IM*
// complex-number formatting, future bytecode-level text coercion) emit
// numeric strings identically. The contract:
//
//   * NaN  -> "nan"
//   * +inf -> "inf",  -inf -> "-inf"
//   * Negative zero collapses to "0".
//   * Excel General format: 15 sig digits via `%.15g` (round-to-nearest-even);
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
// Implementation is dependency-free (`<cstdio>`, `<cmath>`, `<string>`).

#ifndef FORMULON_UTILS_DOUBLE_FORMAT_H_
#define FORMULON_UTILS_DOUBLE_FORMAT_H_

#include <string>

namespace formulon {

/// Appends a locale-independent textual form of `v` to `out`.
///
/// `out` is left untouched on entry; the formatter only appends. See the
/// header comment for the per-case formatting rules.
void format_double(std::string& out, double v);

}  // namespace formulon

#endif  // FORMULON_UTILS_DOUBLE_FORMAT_H_
