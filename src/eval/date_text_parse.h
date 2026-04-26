// Copyright 2026 libraz. Licensed under the MIT License.
//
// Shared date / time text-parsing helpers used by DATEVALUE, TIMEVALUE, and
// VALUE. The primitives here recognise the shapes that Mac Excel 365 accepts
// in ja-JP locale without touching the clock: ISO 8601 dashed dates,
// slash-separated dates, the kanji (年/月/日) form, the wareki era forms
// (令和/平成/昭和/大正/明治 with strict 年/月/日 separators, plus the
// single-letter abbreviations R/H/S/T/M with dot separators), and
// time-of-day tokens with optional fractional seconds, AM/PM markers, and
// the kanji time form `H時M分[S秒]`. Full-width digits (U+FF10..U+FF19)
// are folded to ASCII before tokenisation.

#ifndef FORMULON_EVAL_DATE_TEXT_PARSE_H_
#define FORMULON_EVAL_DATE_TEXT_PARSE_H_

#include <string_view>

namespace formulon {
namespace eval {
namespace date_parse {

/// Trims leading and trailing ASCII whitespace from `s`. The Japanese
/// fullwidth space (U+3000) is NOT trimmed here; Mac Excel 365 rejects
/// `DATEVALUE("　 2024-03-15 　")` as #VALUE! when the U+3000
/// appears at the head, so the leading side must stay untouched. The
/// tail cleanup after a successful parse is performed by
/// `parse_date_time_text` itself.
std::string_view trim_date_text(std::string_view s) noexcept;

/// Tries to parse a leading date + time stream out of `s`, consuming the
/// entire input (after the tail-cleanup rules applied internally). On
/// success returns true and writes:
///
///   * `*out_date_serial` - Excel serial for the date component (0 if none)
///   * `*out_time_frac`   - fractional day for the time component (0 if none)
///   * `*out_has_date`    - true iff a date token was parsed
///   * `*out_has_time`    - true iff a time token was parsed
///
/// At least one of `*out_has_date` / `*out_has_time` is true on success.
/// On any parse failure returns false and leaves the out parameters
/// untouched. The recognised shapes are documented in the implementation.
bool parse_date_time_text(std::string_view s, double* out_date_serial, double* out_time_frac, bool* out_has_date,
                          bool* out_has_time) noexcept;

}  // namespace date_parse
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DATE_TEXT_PARSE_H_
