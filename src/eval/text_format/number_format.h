// Copyright 2026 libraz. Licensed under the MIT License.
//
// Excel TEXT() format-string engine.
//
// `apply_format(value, format, out)` renders `value` through `format`
// following the Excel 365 surface documented in
// https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-.
// The implemented subset is:
//
//   Number tokens:  `0`, `#`, `?`, `.`, `,`, `%`, `E+/E-/e+/e-`, sign.
//   Date tokens:    `y`/`yy`/`yyyy`, `m`/`mm`/`mmm`/`mmmm`, `d`/`dd`/`ddd`/
//                   `dddd`, `h`/`hh`, `m`/`mm` (minute), `s`/`ss`, `.0*`
//                   fractional seconds, `AM/PM` / `A/P`, `[h]`/`[m]`/`[s]`
//                   elapsed brackets.
//   Literal:        `"..."`, `\x`, `!x`, plus any character that does not
//                   match a token (so e.g. `円`, `-`, ` ` pass through).
//   Section split:  `;` (up to 4 sections: positive; negative; zero; text).
//   Discarded:      `[Red]` / `[Blue]` / ... colour specifiers, currency
//                   locale prefixes like `[$-409]` (treated as inert).
//   Text-section:   `@` substitutes the original text input in the text
//                   section of the format.
//
// Deferred (documented as divergences via `tests/divergence.yaml` if they
// trip the oracle): wareki eras, fractional formats `# ?/?`, conditional
// sections `[>100]...`, and DBNum digit styles.
//
// The engine is stateless: a call with identical (value, format) returns
// the same string on every thread.

#ifndef FORMULON_EVAL_TEXT_FORMAT_NUMBER_FORMAT_H_
#define FORMULON_EVAL_TEXT_FORMAT_NUMBER_FORMAT_H_

#include <string>
#include <string_view>

namespace formulon {
namespace eval {
namespace text_format {

/// Return codes for `apply_format`. The engine surfaces `Value` / `Num`
/// errors to the caller rather than throwing; parsing failures and
/// unrepresentable renderings both map to `kValueError` so the caller can
/// forward them as Excel's `#VALUE!`.
enum class FormatStatus : int {
  kOk = 0,
  kValueError = 1,
};

/// Renders the numeric scalar `value` through `format`, appending the result
/// to `out`. Returns `kOk` on success, `kValueError` on any format-parse or
/// rendering failure.
///
/// An empty `format` yields an empty result. `original_text` is the raw text
/// form of the caller's value; if the format contains an `@` token in the
/// text section, `original_text` is substituted there. When the caller
/// passed a non-text value, `original_text` should be empty.
///
/// The engine does not perform any argument coercion itself: the caller is
/// responsible for supplying a finite `value` (non-NaN, non-Inf) and a
/// non-owning `std::string_view` for `original_text`. The format string is
/// interpreted as UTF-8 bytes; quoted literal text and non-token bytes are
/// copied verbatim.
FormatStatus apply_format(double value, std::string_view format, std::string_view original_text, std::string& out);

/// Convenience overload for the pure-numeric path. Equivalent to
/// `apply_format(value, format, {}, out)`.
inline FormatStatus apply_format(double value, std::string_view format, std::string& out) {
  return apply_format(value, format, std::string_view{}, out);
}

}  // namespace text_format
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_TEXT_FORMAT_NUMBER_FORMAT_H_
