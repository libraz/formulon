// Copyright 2026 libraz. Licensed under the MIT License.
//
// Locale-independent shortest-form formatter for `double`.
//
// `format_double` exists so multiple subsystems (the AST S-expression dumper,
// the tree-walk evaluator's `coerce_to_text`, future bytecode-level text
// coercion) emit numeric strings identically. The contract:
//
//   * NaN  -> "nan"
//   * +inf -> "inf",  -inf -> "-inf"
//   * Negative zero collapses to "0".
//   * Any value whose magnitude is below 1e16 and whose fractional part is
//     zero is printed in plain integer form (no decimal point).
//   * Otherwise we use `std::to_string` and trim trailing zeros plus any
//     stranded trailing dot.
//
// The implementation is intentionally dependency-free; it uses only `<cmath>`
// and `<string>`. When double-conversion is wired in we will replace the
// fallback path with Grisu3 here, and every caller will pick up the change
// without modification.

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
