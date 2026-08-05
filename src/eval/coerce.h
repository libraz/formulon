//
// Scalar coercion helpers shared by the tree-walk evaluator and the built-in
// function implementations. Each helper returns an `Expected<T, ErrorCode>`
// carrying the Excel-visible error code on failure (`#VALUE!`, `#NUM!`, ...).
//
// The semantics match Excel 365's implicit conversion rules for scalar
// arithmetic and string contexts.

#ifndef FORMULON_EVAL_COERCE_H_
#define FORMULON_EVAL_COERCE_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

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

/// Allocation-free overload that runs the same `Text`-branch coercion logic
/// as `coerce_to_number(const Value&)` directly against a `string_view`.
/// Used by hot paths (criterion parsing in COUNTIFS / SUMIFS) that would
/// otherwise wrap a `string_view` in a temporary `Value::text(...)` solely
/// to satisfy the Value-shaped overload, paying for a heap allocation per
/// criterion × cell. The behaviour matches the `ValueKind::Text` arm of
/// the Value-shaped overload byte-for-byte: empty / whitespace-only input
/// yields `#VALUE!`, plain numerics / percent / currency / date strings
/// pass through their respective fallbacks.
Expected<double, ErrorCode> coerce_text_to_number(std::string_view text);

/// Coerces `v` to its Excel-visible string representation.
///
/// * `Number` is rendered via `format_double` (Grisu3 shortest round-trip).
/// * `Bool` becomes the literal `"TRUE"` / `"FALSE"`.
/// * `Blank` becomes the empty string.
/// * `Text` is returned verbatim.
/// * `Error` propagates its code unchanged.
/// * `Array`, `Ref`, and `Lambda` yield `#VALUE!`.
Expected<std::string, ErrorCode> coerce_to_text(const Value& v);

/// Coerces `v` to a boolean following Mac Excel 365 (ja-JP) truthiness rules.
///
/// * `Bool` is returned as-is.
/// * `Number` is `false` iff exactly zero (NaN / Inf yield `#NUM!`).
/// * `Blank` is `false`.
/// * `Text` accepts ONLY the literal strings `"TRUE"` / `"FALSE"` (ASCII
///   case-insensitive, no whitespace tolerance). Every other text —
///   including numeric strings (`"0"`, `"1"`, `"0.5"`), whitespace-padded
///   forms (`"  TRUE  "`), and localised truth-words — yields `#VALUE!`.
///   This matches the `text_to_bool_probes` oracle suite and keeps
///   `coerce_to_bool` in sync with the stricter `logical_coerce` helper
///   (`eval/logical_coerce.h`) used by AND / OR / XOR / IFS.
/// * `Error` propagates its code unchanged.
/// * `Array`, `Ref`, and `Lambda` yield `#VALUE!`.
Expected<bool, ErrorCode> coerce_to_bool(const Value& v);

/// Parses `s` as a full double via `std::strtod`. Returns `true` iff the
/// entire input was consumed (no trailing bytes). The numeric value is
/// written to `*out` even if it is `NaN` / `Inf`; callers that require a
/// finite result must run their own `std::isnan` / `std::isinf` guard.
///
/// This is the same scanner that `coerce_text_to_number` uses internally
/// for the numeric-fast-path; it is exposed so the IM* / COMPLEX
/// parsing helpers can avoid duplicating the stack-buffer + heap-fallback
/// dance.
bool strtod_full(std::string_view s, double* out) noexcept;

/// Computes `base ^ exp` with Excel's edge-case handling: a NaN/Inf result
/// is reported as `#NUM!`. This is shared between the `^` operator and the
/// `POWER()` builtin so the two paths cannot diverge.
///
/// * `POWER(0, 0)` yields `#NUM!` — Excel treats it as indeterminate,
///   diverging from IEEE-754 `std::pow` (which returns 1).
/// * Negative base with a non-integer exponent yields `#NUM!` (NaN from pow).
/// * Overflow / underflow to Inf yields `#NUM!`.
Expected<double, ErrorCode> apply_pow(double base, double exp);

// ---------------------------------------------------------------------------
// Matrix-strict numeric coercion
// ---------------------------------------------------------------------------
//
// Functions that consume a "matrix" of numbers — `LINEST`, `LOGEST`,
// `TREND`, `GROWTH`, `FORECAST.ETS`, and friends — refuse non-numeric
// cells outright rather than silently dropping them the way `AVERAGE` /
// `SUM` do. Numbers pass through; Booleans coerce to 1/0; everything
// else (Blank, Text, Array, Ref, Lambda) surfaces `#VALUE!`. Errors
// propagate verbatim.

/// Matrix-strict numeric coercion. Used by `LINEST` / `LOGEST` /
/// `FORECAST.ETS` etc. for every cell of a "numbers only" matrix
/// argument.
///
/// * `Number` is returned as-is (including NaN / Inf; matrix-strict
///   callers expect raw passthrough rather than the `coerce_to_number`
///   finiteness guard, mirroring `forecast_ets_lazy.cpp::coerce_strict_numeric`).
/// * `Bool` becomes 1.0 / 0.0.
/// * `Error` propagates its code unchanged.
/// * `Blank`, `Text`, `Array`, `Ref`, `Lambda` all yield `#VALUE!`. In
///   particular Text — even numeric-looking text like `"3.14"` — is
///   rejected, which is the whole point of the strict variant.
Expected<double, ErrorCode> matrix_strict_number(const Value& v);

/// Same as `matrix_strict_number` but with a `(row, col)` context for
/// future structured-log enrichment. The semantics are identical; the
/// row / col arguments are accepted today so call sites in `regression`
/// / `forecast` families can be migrated without churn when richer
/// diagnostics land.
Expected<double, ErrorCode> matrix_strict_number_cell(const Value& v, std::uint32_t row, std::uint32_t col);

// ---------------------------------------------------------------------------
// Configurable numeric collection
// ---------------------------------------------------------------------------

/// Policy flags controlling how `collect_numerics` treats each
/// non-`Number` kind it encounters while flattening a `Value`.
///
/// The defaults model the production `stats_detail::collect_numerics`
/// behaviour (the AVERAGE / SUM / VAR family): Numbers are kept,
/// every other kind is silently dropped. Flag combinations recover
/// the "A"-family (`include_bool` + `include_text_numeric_literal` +
/// `error_on_text`) and the `SMALL` / `LARGE` direct-scalar variant.
struct NumericCollectPolicy {
  /// If `true`, `Bool` cells contribute 1.0 / 0.0 instead of being
  /// dropped. Matches the "A"-family and the `SMALL` / `LARGE` direct-
  /// scalar rule.
  bool include_bool = false;

  /// If `true`, `Text` cells whose contents parse via `coerce_to_number`
  /// (numeric literals like `"3.14"`, percent / currency / date strings)
  /// contribute the parsed value. If `false`, `Text` cells are skipped
  /// (the default `AVERAGE`-family rule).
  bool include_text_numeric_literal = false;

  /// Only meaningful when `include_text_numeric_literal` is `true`:
  /// if a `Text` cell fails `coerce_to_number` (e.g. `"hello"`), the
  /// collection aborts with `#VALUE!` instead of silently skipping.
  /// Matches the "A"-family contract and the `SMALL` / `LARGE`
  /// direct-scalar variant.
  bool error_on_text = false;

  /// If `true` (default), encountering an `Error` cell aborts the
  /// collection by propagating that error code. Set to `false` only
  /// when the caller has its own error-propagation pass (e.g.
  /// regression families that walk both arrays first to surface the
  /// leftmost error).
  bool error_on_error_cell = true;

  /// If `true`, `Blank` cells contribute `0.0` instead of being dropped.
  /// Matches the "A"-family contract for direct `Blank` arguments
  /// (`AVERAGEA(A1, 1)` where A1 is blank counts the blank slot as 0).
  /// Default `false` keeps the AVERAGE / SUM rule (Blank dropped).
  bool blank_as_zero = false;
};

/// Flattens `v` into a vector of `double` according to `policy`.
///
/// Iteration: `Array` cells are visited in row-major order; any other
/// kind is treated as a single-cell input. `Blank` cells are always
/// dropped (counting blanks as zero is `collect_a`'s job, not this
/// helper's). `Ref` and `Lambda` are always dropped — callers that
/// care about them must resolve refs before calling in.
///
/// Returns the propagated `ErrorCode` if `policy.error_on_error_cell`
/// is set and an `Error` cell is encountered, or if
/// `policy.error_on_text` is set and a `Text` cell fails `coerce_to_number`.
Expected<std::vector<double>, ErrorCode> collect_numerics(const Value& v, NumericCollectPolicy policy);

/// Multi-`Value` overload: flattens each of `args[0..count-1]` into the
/// same output vector, applying `policy` per cell. Used by builtin
/// dispatchers that receive a flat `Value*` slice (the dispatcher has
/// already expanded any range arguments into scalar cells), so iterating
/// the slice directly is equivalent to walking each cell as a scalar
/// `Value` and accumulating. Errors short-circuit as soon as they are
/// encountered (no partial result).
Expected<std::vector<double>, ErrorCode> collect_numerics(const Value* args, std::uint32_t count,
                                                          NumericCollectPolicy policy);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_COERCE_H_
