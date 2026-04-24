// Copyright 2026 libraz. Licensed under the MIT License.
//
// Scalar coercion helpers shared by the tree-walk evaluator and the built-in
// function implementations. Each helper returns an `Expected<T, ErrorCode>`
// carrying the Excel-visible error code on failure (`#VALUE!`, `#NUM!`, ...).
//
// The semantics match Excel 365's implicit conversion rules for scalar
// arithmetic and string contexts; see `backup/plans/02-calc-engine.md`
// §2.1.1 for the authoritative table.

#ifndef FORMULON_EVAL_COERCE_H_
#define FORMULON_EVAL_COERCE_H_

#include <string>

#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {

/// Coerces `v` to a finite numeric value following Excel's implicit-conversion
/// rules.
///
/// * `Number` is returned as-is, or `#NUM!` if non-finite.
/// * `Bool` becomes 1.0 / 0.0.
/// * `Blank` becomes 0.0.
/// * `Text` is parsed via `std::strtod` after trimming; on strtod failure a
///   date / datetime fallback (`date_parse::parse_date_time_text`) accepts
///   ISO 8601 (`"2024-01-10"`), slash (`"2024/01/10"`), kanji
///   (`"2024年1月10日"`), and date+time (`"2024-01-10 12:00"`) shapes,
///   matching Mac Excel 365 ja-JP coercion. Empty / whitespace-only text
///   yields `#VALUE!`; an otherwise unparseable string yields `#VALUE!`;
///   a parse that produces a non-finite double yields `#NUM!`.
/// * `Error` propagates its code unchanged.
/// * `Array`, `Ref`, and `Lambda` are unsupported in scalar contexts and
///   yield `#VALUE!` defensively.
Expected<double, ErrorCode> coerce_to_number(const Value& v);

/// Coerces `v` to its Excel-visible string representation.
///
/// * `Number` is rendered via `format_double` (Grisu3 shortest round-trip).
/// * `Bool` becomes the literal `"TRUE"` / `"FALSE"`.
/// * `Blank` becomes the empty string.
/// * `Text` is returned verbatim.
/// * `Error` propagates its code unchanged.
/// * `Array`, `Ref`, and `Lambda` yield `#VALUE!`.
Expected<std::string, ErrorCode> coerce_to_text(const Value& v);

/// Coerces `v` to a boolean following Excel's truthiness rules.
///
/// * `Bool` is returned as-is.
/// * `Number` is `false` iff exactly zero (NaN / Inf yield `#NUM!`).
/// * `Blank` is `false`.
/// * `Text` is coerced to a number first; on success the numeric rule
///   applies (`"0"` -> false, `"1"` -> true). Empty / whitespace-only text
///   parses to 0 -> false. Non-numeric text (including the literal strings
///   `"TRUE"` / `"FALSE"`) yields `#VALUE!`, matching Excel's actual
///   behaviour for `AND` / `OR` / `NOT` / `IF` argument coercion.
/// * `Error` propagates its code unchanged.
/// * `Array`, `Ref`, and `Lambda` yield `#VALUE!`.
Expected<bool, ErrorCode> coerce_to_bool(const Value& v);

/// Computes `base ^ exp` with Excel's edge-case handling: a NaN/Inf result
/// is reported as `#NUM!`. This is shared between the `^` operator and the
/// `POWER()` builtin so the two paths cannot diverge.
///
/// * `POWER(0, 0)` yields `1` (matches IEEE-754 `std::pow`).
/// * Negative base with a non-integer exponent yields `#NUM!` (NaN from pow).
/// * Overflow / underflow to Inf yields `#NUM!`.
Expected<double, ErrorCode> apply_pow(double base, double exp);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_COERCE_H_
