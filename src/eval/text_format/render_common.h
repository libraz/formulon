//
// Internal header -- do not include outside `src/eval/text_format/`.
//
// Shared helpers used by the renderer translation units
// (`number_format_render.cpp`, `render_numeric.cpp`, `render_date.cpp`,
// `render_fraction.cpp`). The bodies live in `number_format_render.cpp`
// so the symbols have a single owner.

#ifndef FORMULON_EVAL_TEXT_FORMAT_RENDER_COMMON_H_
#define FORMULON_EVAL_TEXT_FORMAT_RENDER_COMMON_H_

#include <cstddef>
#include <string>
#include <string_view>

#include "eval/text_format/number_format_types.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {

// --- Digit-style substitution ----------------------------------------
//
// A section's digit style (`Section::digit_style`, null for ASCII) writes
// digit placeholders and full years digit by digit (`=TEXT(1234,"[DBNum1]0")`
// -> 一二三四). A style with place units writes General, elapsed time and the
// other date fields with them (千二百三十四); a digit set substitutes every digit.

// Returns the substitution for `c` under `style`, or an empty string if none
// applies (caller falls back to `c` verbatim).
std::string_view dbnum_digit_subst(const DbnumStyle* style, char c) noexcept;

// Append a single ASCII digit `c`, substituting it via `style` if applicable.
// Non-digit characters fall through verbatim.
void append_digit_dbnum(std::string& out, const DbnumStyle* style, char c);

// Append every character of `chars`, substituting digits via `style`.
void append_chars_dbnum(std::string& out, const DbnumStyle* style, std::string_view chars);

// Appends `value` to `out` with each digit substituted per `style`.
void append_int_dbnum(std::string& out, long long value, const DbnumStyle* style);

// Appends the decimal integer `digits` with the place units of `style`;
// digit by digit when it has none.
void append_dbnum_positional(std::string& out, const DbnumStyle* style, std::string_view digits);

// True when `style` writes numerals with place units of its own, rather than
// leaving ASCII digits or substituting digit by digit.
bool dbnum_writes_numerals(const DbnumStyle* style) noexcept;

// Appends a date or time field zero-padded to `width` digits, with place
// units where `style` writes date fields that way.
void append_dbnum_date_field(std::string& out, const DbnumStyle* style, unsigned value, std::size_t width);

// Append `n` zero-padded to two characters without any DBNum substitution
// (ASCII output).
void append_pad2(std::string& out, unsigned n);

// Appends `value` zero-padded to 2 digits, with DBNum substitution applied.
void append_pad2_dbnum(std::string& out, unsigned value, const DbnumStyle* style);

// Writes a percent token: its kept full-width glyph, else `%`.
void append_percent(std::string& out, std::string_view fmt, const Token& token);

// True when every byte in `digits` is the ASCII '0' character. Used by
// the numeric walker to suppress a stray minus sign when a tiny negative
// magnitude rounds down to a representation of zero.
bool decimal_digits_all_zero(std::string_view digits) noexcept;

// Zeros every digit of `digits` past the 15th significant one, rounding the
// 15th half-away-from-zero against the 16th. `digits` holds decimal digits
// only (no sign); its length is preserved, so only trailing digits collapse
// to zero. Excel shows no more than 15 significant digits: `2^60` displays
// as `1152921504606850000`, not the exact 19-digit integer `%.0f` prints.
void cap_integer_significant_digits(std::string* digits);

}  // namespace number_format_detail
}  // namespace text_format
}  // namespace formulon

#endif  // FORMULON_EVAL_TEXT_FORMAT_RENDER_COMMON_H_
