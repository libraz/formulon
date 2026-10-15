//
// Locale-aware numeric string parsing shared by the VALUE / NUMBERVALUE
// builtins and by the implicit text->number coercion used by arithmetic
// operators and the COUNTIF / SUMIF / AVERAGEIF criteria engine. Excel
// applies the same numeric normalisation in all these contexts (full-width
// digits, accounting parentheses, thousands grouping, currency symbols, a
// trailing percent), so the logic lives here in one place.

#ifndef FORMULON_EVAL_NUMBER_PARSE_H_
#define FORMULON_EVAL_NUMBER_PARSE_H_

#include <cstddef>
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
/// `text_percent` >= 0 selects the text rule (VALUE and implicit coercion;
/// NUMBERVALUE passes -1): the `%` count was already taken out by
/// `normalize_locale_numeric`, no `%` may remain, and the value follows
/// `excel_text_number`.
bool parse_numeric(std::string_view s, char decimal_sep, char group_sep, double* out, bool strict_groups = true,
                   int text_percent = -1) noexcept;

/// Longest text, in UTF-16 units and spaces included, that Excel converts to
/// a number.
inline constexpr std::size_t kMaxNumericTextLength = 255;

/// True when the raw text exceeds `kMaxNumericTextLength`; checked on the
/// input before any normalisation.
bool numeric_text_too_long(std::string_view raw) noexcept;

/// Converts the canonical decimal `[digits][.digits][e[+-]digits]` (no sign)
/// the way Excel converts text: digits past the 15th are dropped, not rounded
/// (an integer digit still scales), `percent_count` folds into the exponent
/// before the range check, a leading-digit exponent outside [-308, 308] fails,
/// and a result below the smallest normal double becomes 0. The caller has
/// already validated the grammar.
bool excel_text_number(std::string_view magnitude, bool negative, int percent_count, double* out);

/// Locale-input pre-pass shared by VALUE / NUMBERVALUE. Japanese profiles
/// fold full-width ASCII forms (U+FF01..U+FF5E) and the ideographic space
/// (U+3000) to their ASCII equivalents; all profiles strip ASCII
/// accounting-style outer parentheses (`"(1234)"` -> `"1234"` with
/// `*paren_negated` set). Currency-symbol stripping is left to
/// `parse_numeric`. A non-null `text_percent` additionally takes the single
/// allowed `%` out of the text and reports its count (0 or 1), or -1 when the
/// placement is invalid.
std::string normalize_locale_numeric(std::string_view raw, bool* paren_negated, int* text_percent = nullptr);

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
