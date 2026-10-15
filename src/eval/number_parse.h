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
#include <cstdint>
#include <string>
#include <string_view>

namespace formulon {
namespace eval {

/// Parses numeric text, the grammar shared by VALUE, NUMBERVALUE and implicit
/// coercion, using `decimal_sep` and `group_sep`. `group_sep` may appear only
/// in the integer part: the first group is 1+ digits not starting with `0`,
/// every later group 3+ digits; pass `'\0'` to disable grouping.
/// Accepted around the digits: a sign, the profile's currency symbols on one
/// side only (a trailing Euro included), and accounting parentheses
/// (negative), in the orders Excel takes them. A mixed fraction `N a/b` (an
/// integer, one blank, then digits `/` digits, each at most 32767, `b > 0`) is
/// `N + a/b`. `percent_count` is the number
/// of `%` signs already taken out by `normalize_locale_numeric`; currency
/// combined with a percent is not a number. The value follows
/// `excel_text_number`. Returns true and writes the value into `*out` on
/// success.
bool parse_numeric(std::string_view s, char decimal_sep, char group_sep, int percent_count, double* out) noexcept;

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

/// Locale-input pre-pass shared by VALUE / NUMBERVALUE. Folds the decimal
/// digit scripts Excel reads (Arabic-Indic, Devanagari, Thai, ...) to ASCII;
/// Japanese profiles also fold full-width ASCII forms (U+FF01..U+FF5E, the
/// full-width dollar sign included) and the ideographic space (U+3000).
/// Currency-symbol and parenthesis handling is left to `parse_numeric`. The
/// single allowed `%` is taken out of the text and counted in `*percent_count`
/// (0 or 1), or -1 when the placement is invalid.
std::string normalize_locale_numeric(std::string_view raw, int* percent_count);

/// `text` with the decimal digit scripts (Unicode Nd; not the Ethiopic digits
/// numbers read) folded to ASCII and nothing else changed; what the date /
/// time recogniser sees.
std::string fold_script_digits(std::string_view text);

/// The first UTF-16 unit of `text` (a lead surrogate for a supplementary-plane
/// character), or 0 for empty text.
std::uint32_t first_utf16_unit(std::string_view text) noexcept;

/// NUMBERVALUE's text pre-pass. Fails when the raw text exceeds
/// `kMaxNumericTextLength` or the two separators are the same unit. Otherwise
/// deletes U+0020 / TAB / LF / CR, replaces each code point equal to
/// `decimal_unit` with the profile's decimal separator, drops each one equal
/// to `group_unit` (a failure once a decimal separator was seen) and takes the
/// trailing ASCII `%` signs into `*percent_count`. The caller hands `*out` to
/// the VALUE recogniser and divides the result by 100 `*percent_count` times.
bool numbervalue_prepass(std::string_view text, std::uint32_t decimal_unit, std::uint32_t group_unit, std::string* out,
                         int* percent_count);

/// Convenience wrapper reproducing the VALUE() function's numeric phase:
/// `normalize_locale_numeric` followed by `parse_numeric` with the active
/// profile's separators.
/// Does NOT handle date/time strings — callers that need the DATEVALUE-style
/// fallback layer it separately. Returns true and writes `*out` on success.
bool parse_excel_number(std::string_view text, double* out);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_NUMBER_PARSE_H_
