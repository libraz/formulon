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
/// * `Text` is parsed via `std::strtod` after trimming; empty / whitespace
///   text is 0.0; an unparseable suffix yields `#VALUE!`; a parse that
///   produces a non-finite double yields `#NUM!`.
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
/// * `Text` accepts `"TRUE"` / `"FALSE"` case-insensitively (ASCII); any
///   other text yields `#VALUE!`.
/// * `Error` propagates its code unchanged.
/// * `Array`, `Ref`, and `Lambda` yield `#VALUE!`.
Expected<bool, ErrorCode> coerce_to_bool(const Value& v);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_COERCE_H_
