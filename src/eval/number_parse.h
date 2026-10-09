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

/// Parses a numeric string using `decimal_sep` and `group_sep`. `group_sep`
/// may appear only in the integer part and must form valid 3-digit groups
/// (the first group being 1-3 digits); pass `'\0'` to disable grouping.
/// With `strict_groups` false (NUMBERVALUE) group separators are dropped
/// wherever they appear in the integer part.
/// A leading sign, the profile's accepted currency symbols, a trailing Euro
/// suffix, and any number of trailing `%` signs (each dividing the result by
/// 100) are accepted. Surrounding ASCII whitespace is trimmed. Returns true
/// and writes the value into `*out` on success.
bool parse_numeric(std::string_view s, char decimal_sep, char group_sep, double* out,
                   bool strict_groups = true) noexcept;

/// Locale-input pre-pass shared by VALUE / NUMBERVALUE. Japanese profiles
/// fold full-width ASCII forms (U+FF01..U+FF5E) and the ideographic space
/// (U+3000) to their ASCII equivalents; all profiles strip ASCII
/// accounting-style outer parentheses (`"(1234)"` -> `"1234"` with
/// `*paren_negated` set). Currency-symbol stripping is left to
/// `parse_numeric`.
std::string normalize_locale_numeric(std::string_view raw, bool* paren_negated);

/// True when decimal numeric text (optional sign, digits with an optional
/// point, optional exponent) denotes a magnitude above 9.99999999999999E+307
/// once the digits past the 15th are dropped. Excel rejects such text as a
/// number. Text of any other shape yields false.
bool numeric_text_above_excel_max(std::string_view text);

/// Convenience wrapper reproducing the VALUE() function's numeric phase:
/// `normalize_locale_numeric` followed by `parse_numeric` with the active
/// profile's separators and accounting-paren negation.
/// Does NOT handle date/time strings — callers that need the DATEVALUE-style
/// fallback layer it separately. Returns true and writes `*out` on success.
bool parse_excel_number(std::string_view text, double* out);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_NUMBER_PARSE_H_
