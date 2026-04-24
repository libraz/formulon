// Copyright 2026 libraz. Licensed under the MIT License.
//
// Criterion parsing and matching for the conditional aggregators `COUNTIF`,
// `SUMIF`, and `AVERAGEIF`. The helpers here operate purely on scalar
// `Value`s so they are reusable across the eager dispatcher and the lazy
// range-aware impls in `tree_walker.cpp`.
//
// A criterion is a scalar "predicate template" that Excel compiles out of
// the second argument to `COUNTIF`/`SUMIF`/`AVERAGEIF`. Supported shapes:
//
//   * Bare numeric or boolean: `5`, `TRUE` -> equality comparison against
//     that number (booleans fold to 1.0 / 0.0).
//   * Text with no leading comparator: treated as `Op::Eq` against the
//     text, with DOS-style wildcards (`*`, `?`, `~` escape).
//   * Text with a leading comparator (`<`, `<=`, `>`, `>=`, `<>`, `=`):
//     the RHS is probed as a number (via the `Text` branch of
//     `coerce_to_number`); on success the criterion is numeric, otherwise
//     the RHS is compared as text (case-insensitive ASCII).
//   * Blank Value: treated as `Op::Eq` against text "".
//
// Wildcard matching uses byte-level semantics: `?` matches one byte, `*`
// matches any run of bytes, and `~` escapes the next metacharacter. For
// ASCII-only text this is Excel-exact; multibyte UTF-8 is an accepted
// divergence (see comment in `matches_criterion` below). This matches the
// range-vs-direct divergence already documented in `eval_context.cpp`.

#ifndef FORMULON_EVAL_CRITERIA_H_
#define FORMULON_EVAL_CRITERIA_H_

#include <cstdint>
#include <string>
#include <string_view>

#include "value.h"

namespace formulon {
namespace eval {

/// Comparator encoded by a `COUNTIF` / `SUMIF` / `AVERAGEIF` criterion.
///
/// `Eq` is the default when the criterion text carries no leading
/// comparator, and also the representation used for bare numeric / boolean
/// criterion values. The remaining operators correspond to the
/// Excel-visible prefixes `<>`, `<`, `<=`, `>`, `>=`.
enum class CriteriaOp : std::uint8_t {
  Eq,     ///< "=5", "foo", or bare scalar number/bool (no prefix).
  NotEq,  ///< "<>5"
  Lt,     ///< "<5"
  LtEq,   ///< "<=5"
  Gt,     ///< ">5"
  GtEq,   ///< ">=5"
};

/// Parsed form of a `COUNTIF` / `SUMIF` / `AVERAGEIF` criterion.
///
/// Exactly one shape is populated:
///   * `rhs_is_number == true` — numeric comparison. `rhs_number` carries
///     the RHS, `rhs_text` is ignored.
///   * `rhs_is_number == false` — text comparison. `rhs_text` carries the
///     RHS (possibly empty), and `has_wildcard` is true iff the RHS
///     contains an unescaped `*` or `?`.
///
/// The `rhs_text` view may alias the input `Value::text` storage when no
/// rewriting was required, or it may point into the owned `rhs_storage`
/// string when the input needed a comparator prefix stripped. Either way,
/// the view stays valid for the lifetime of the `ParsedCriterion`.
struct ParsedCriterion {
  CriteriaOp op = CriteriaOp::Eq;
  bool rhs_is_number = false;
  double rhs_number = 0.0;
  std::string_view rhs_text;  ///< Valid when `!rhs_is_number`; may be empty.
  bool has_wildcard = false;  ///< Unescaped `*` or `?` present in `rhs_text`.
  /// True when the numeric RHS originated from a Bool criterion (the caller
  /// passed `TRUE` or `FALSE`, not a number or `">1"`). Excel's criterion
  /// matching is type-strict: a Bool criterion matches only Bool cells, and
  /// a Number criterion matches only Number cells. The matcher branches on
  /// this flag in the numeric path.
  bool rhs_from_bool = false;
  /// True when the criterion itself is an error Value (e.g. a cell holding
  /// `#N/A` passed as `COUNTIF(range, #N/A)`). The matcher counts cells whose
  /// error code equals `rhs_error_code` and every other kind is a miss.
  /// Observed in Excel 365: `COUNTIF` does not propagate an error criterion
  /// but uses it as a filter over error cells in the range.
  bool rhs_is_error = false;
  ErrorCode rhs_error_code = ErrorCode::NA;

 private:
  /// Owns the re-buffered RHS when the parser needed to strip a comparator
  /// prefix from an interned `Value::text` view. External callers should
  /// not read this directly; use `rhs_text` instead.
  std::string rhs_storage;

  friend ParsedCriterion parse_criterion(const Value&);
};

/// Parses a scalar `Value` into a `ParsedCriterion`.
///
/// Semantics:
///   * `Number` -> `Op::Eq`, numeric RHS.
///   * `Bool`   -> `Op::Eq`, numeric RHS (TRUE = 1.0, FALSE = 0.0).
///   * `Blank`  -> `Op::Eq` with text RHS "" (matches blank cells and
///     formula-result empty strings; see `matches_criterion`).
///   * `Text`   -> scanned for a leading comparator (`<>`, `<=`, `>=`,
///     `<`, `>`, `=` — longer forms checked first). The remaining RHS is
///     probed as a number by constructing a temporary `Value::text` and
///     calling `coerce_to_number`; on success the criterion becomes
///     numeric, else text. Wildcards in text RHS set `has_wildcard`.
///
/// Error `Value`s must be rejected by the caller before invoking this
/// helper; this function assumes the input is not an error.
ParsedCriterion parse_criterion(const Value& criterion);

/// Tests whether a cell `Value` matches a parsed criterion.
///
/// Matching rules (summary):
///   * Blank cell:
///       - `Op::Eq`  with text RHS ""  -> true (matches blanks AND "")
///       - `Op::NotEq` with text RHS "" -> false
///       - `Op::NotEq` with any non-empty RHS -> true (blank != value)
///       - All other ops against blank -> false
///   * Error cell: always false. Excel silently skips errors in a
///     criteria range; `COUNTIF`/`SUMIF`/`AVERAGEIF` do not propagate them.
///   * Numeric criterion (`rhs_is_number == true`): the cell is coerced
///     to a number via `coerce_to_number`; on failure the cell does not
///     match (numeric-text that fails to parse is simply rejected, not an
///     error).
///   * Text criterion without wildcards: case-insensitive ASCII equality
///     against the cell's `coerce_to_text` rendering for `Op::Eq`/`NotEq`;
///     ordering ops use `case_insensitive_compare`.
///   * Text criterion with wildcards (`Op::Eq`/`NotEq` only): byte-level
///     two-pointer match where `*` matches any byte run, `?` matches
///     exactly one byte, and `~` escapes the next metacharacter. Ordering
///     ops with wildcard RHS fall back to literal-text compare — Excel
///     treats `*`/`?` as literals for non-equality comparators.
///
/// Accepted divergence (documented here and in the range-expansion code):
/// wildcard matching operates on bytes, not Unicode code points, so `?`
/// only matches one byte of a multibyte UTF-8 character. ASCII-only text
/// behaves exactly like Excel.
bool matches_criterion(const Value& cell, const ParsedCriterion& criterion);

/// Byte-level wildcard matcher. `pattern` may contain unescaped `*` / `?`
/// and `~`-escaped literals; `text` is matched verbatim. Shared between
/// `matches_criterion` (for `COUNTIF`/`SUMIF`/`AVERAGEIF` equality paths)
/// and scalar lookups such as `MATCH(..., 0)` that honour the same
/// DOS-style wildcard dialect.
///
/// Iterative backtrack (no recursion) keeps stack depth constant. See the
/// header-level divergence note: `?` matches one byte, not one code point.
bool wildcard_match(std::string_view pattern, std::string_view text);

/// Scans `rhs` for an unescaped `*` or `?`. A leading `~` escapes the next
/// byte, so `"~*"` contains no wildcard but `"*~*"` does. Returns `true` on
/// the first unescaped metacharacter encountered.
bool scan_has_wildcard(std::string_view rhs);

/// Returns the byte offset where `pattern` first matches a prefix of
/// `text` starting at that offset (wildcards: `*`, `?`, `~` for escape).
/// Returns `std::string_view::npos` when no position in `text` begins a
/// match. The match is anchored at its start but not at its end — it
/// succeeds as soon as `pattern` is consumed, regardless of any unmatched
/// suffix in `text`. This mirrors the SEARCH contract where the pattern
/// need only match somewhere inside the haystack.
///
/// Callers that need case-insensitive matching must lower-case both inputs
/// first (the matcher is byte-exact). UTF-16 offset mapping is likewise the
/// caller's responsibility; this helper only reports a byte offset into
/// `text`.
std::size_t wildcard_find(std::string_view pattern, std::string_view text);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_CRITERIA_H_
