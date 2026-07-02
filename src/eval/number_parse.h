// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Locale-aware numeric string parsing shared by the VALUE / NUMBERVALUE
// builtins and by the implicit text->number coercion used by arithmetic
// operators and the COUNTIF / SUMIF / AVERAGEIF criteria engine. Excel
// applies the same numeric normalisation in all these contexts (full-width
// digits, accounting parentheses, thousands grouping, currency symbols, a
// trailing percent), so the logic lives here in one place.

#ifndef FORMULON_EVAL_NUMBER_PARSE_H_
#define FORMULON_EVAL_NUMBER_PARSE_H_

#include <string>
#include <string_view>

namespace formulon {
namespace eval {

/// Locale-independent `std::strtod`. Excel's numeric grammar always uses `.`
/// as the decimal separator, but `std::strtod` honours the host process's
/// `LC_NUMERIC` category — so a native embedder that set (e.g.)
/// `LC_NUMERIC=de_DE` would misparse `"1.5"`. This wrapper evaluates the
/// conversion under a cached C locale on the calling thread (via
/// `uselocale`, which is thread-local and therefore safe under the
/// scheduler's worker threads), then restores the previous locale. The
/// contract matches `std::strtod`: `*endptr` points past the consumed
/// prefix. Every numeric parse in the evaluator routes through here.
double parse_double_c_locale(const char* str, char** endptr) noexcept;

/// Parses a numeric string using `decimal_sep` and `group_sep`. `group_sep`
/// may appear only in the integer part and must form valid 3-digit groups
/// (the first group being 1-3 digits); pass `'\0'` to disable grouping.
/// A leading sign, an optional currency symbol (`$` / `¥` / `￥` / `€`),
/// a trailing Euro suffix, and any number of trailing `%` signs (each
/// dividing the result by 100) are accepted. Surrounding ASCII whitespace
/// is trimmed. Returns true and writes the value into `*out` on success.
bool parse_numeric(std::string_view s, char decimal_sep, char group_sep, double* out) noexcept;

/// Locale-input pre-pass shared by VALUE / NUMBERVALUE. Folds full-width
/// ASCII forms (U+FF01..U+FF5E) and the ideographic space (U+3000) to their
/// ASCII equivalents, and strips accounting-style outer parentheses
/// (`"(1234)"` -> `"1234"` with `*paren_negated` set). Currency-symbol
/// stripping is left to `parse_numeric`.
std::string normalize_locale_numeric(std::string_view raw, bool* paren_negated);

/// Convenience wrapper reproducing the VALUE() function's numeric phase:
/// `normalize_locale_numeric` followed by `parse_numeric` with the en-US
/// separators (`.` decimal, `,` grouping) and accounting-paren negation.
/// Does NOT handle date/time strings — callers that need the DATEVALUE-style
/// fallback layer it separately. Returns true and writes `*out` on success.
bool parse_excel_number(std::string_view text, double* out);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_NUMBER_PARSE_H_
