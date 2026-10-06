//
// Excel TEXT() format-string engine.
//
// `apply_format(value, format, out)` renders `value` through `format`
// following the Excel 365 surface documented in
// https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-.
// The implemented subset is:
//
//   Number tokens:  `0`, `#`, `?`, `.`, `,`, `%`, `E+/E-/e+/e-`, sign,
//                   and proper/improper fractions such as `# ?/?`.
//   Date tokens:    `y`/`yy`/`yyyy`, `m`/`mm`/`mmm`/`mmmm`, `d`/`dd`/`ddd`/
//                   `dddd`, `h`/`hh`, `m`/`mm` (minute), `s`/`ss`, `.0*`
//                   fractional seconds, `AM/PM` / `A/P`, `[h]`/`[m]`/`[s]`
//                   elapsed brackets.
//   Literal:        `"..."`, `\x`, `!x`, plus any character that does not
//                   match a token (so e.g. `円`, `-`, ` ` pass through).
//   Section split:  `;` (up to 4 sections: positive; negative; zero; text),
//                   and conditional selectors such as `[>100]`.
//   Digit styles:    `[DBNum1]`, `[DBNum2]`, `[DBNum3]`.
//   Discarded:      `[赤]` / `[Red]` / ... colour specifiers, currency
//                   locale prefixes like `[$-409]` (treated as inert).
//                   Colour names follow the code's `FormatDialect`: the
//                   ja-JP spellings for a TEXT() argument, the English
//                   ones for a stored cell format.
//   Text-section:   `@` substitutes the original text input in the text
//                   section of the format.
//
// Deferred (documented as divergences via `tests/divergence.yaml` if they
// trip the oracle): wareki eras and locale/currency directives whose locale
// semantics require an Excel-compatible locale database.
//
// The engine is stateless: a call with identical (value, format) returns
// the same string on every thread.

#ifndef FORMULON_EVAL_TEXT_FORMAT_NUMBER_FORMAT_H_
#define FORMULON_EVAL_TEXT_FORMAT_NUMBER_FORMAT_H_

#include <string>
#include <string_view>

namespace formulon {
namespace text_format {

/// Spelling of a format code. A TEXT() format argument is written in the UI
/// locale (`[赤]`, `G/標準`); a cell's stored code (styles.xml / XLSB) is the
/// locale-independent canonical form (`[Red]`, `[Color10]`, `General`).
/// Colour names are accepted only in the dialect's own spelling, exactly as
/// Excel rejects `=TEXT(5,"[Red]0")` on a ja-JP host.
enum class FormatDialect : int {
  kLocalized = 0,
  kStored = 1,
};

/// Return codes for `apply_format`. The engine surfaces failures to the
/// caller rather than throwing; callers that only need Excel's TEXT()
/// semantics map every non-`kOk` value to `#VALUE!`.
enum class FormatStatus : int {
  kOk = 0,
  /// The format code is malformed (e.g. an unrecognised bracket).
  kValueError = 1,
  /// The code is valid but the value has no rendering in it: a date or time
  /// section given a serial outside the workbook's calendar range, a value no
  /// conditional arm accepts, or numeric scaling/normalization overflowed to
  /// a non-finite value.
  kOverflow = 2,
};

/// Renders the numeric scalar `value` through `format`, appending the result
/// to `out`. An empty `format` yields an empty result.
///
/// The caller supplies a finite `value` (non-NaN, non-Inf); no coercion is
/// performed. The format string is interpreted as UTF-8 bytes; quoted literal
/// text and non-token bytes are copied verbatim.
/// `date1904` selects the workbook date epoch for date/time tokens (the
/// serial->calendar conversion shifts by 1462 days under the 1904 system).
/// Callers that render a value from a date1904 workbook (the TEXT builtin)
/// must pass the workbook flag; the pure-numeric callers leave it false.
FormatStatus apply_format(double value, std::string_view format, std::string& out, bool date1904 = false,
                          FormatDialect dialect = FormatDialect::kLocalized);

/// Renders the text value `text` through the text section of `format`: the
/// fourth section, or a single section containing `@`, with `@` replaced by
/// `text`. A format without a text section shows `text` unchanged. An empty
/// `text` is still a text value and takes the text section.
FormatStatus apply_text_format(std::string_view text, std::string_view format, std::string& out,
                               FormatDialect dialect = FormatDialect::kLocalized);

/// True when the section of `format` that renders negative numbers (the
/// second section, or the only one) carries a colour qualifier spelled in
/// `dialect`.
bool negative_section_has_color(std::string_view format, FormatDialect dialect);

}  // namespace text_format
}  // namespace formulon

#endif  // FORMULON_EVAL_TEXT_FORMAT_NUMBER_FORMAT_H_
