//
// Internal header -- do not include outside `src/eval/text_format/`.
//
// Shared helpers used by the renderer translation units
// (`number_format_render.cpp`, `render_numeric.cpp`, `render_date.cpp`,
// `render_fraction.cpp`). The bodies live in `number_format_render.cpp`
// so the symbols have a single owner.

#ifndef FORMULON_EVAL_TEXT_FORMAT_RENDER_COMMON_H_
#define FORMULON_EVAL_TEXT_FORMAT_RENDER_COMMON_H_

#include <string>
#include <string_view>

#include "eval/text_format/number_format_types.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {

// --- DBNum digit substitution -----------------------------------------
//
// `[DBNum1]` / `[DBNum2]` / `[DBNum3]` are *per-digit* substitutions in
// Mac Excel 365 ja-JP. Despite their popular description as "positional
// kanji" formats, the oracle corpus shows that Excel does NOT decompose
// integers into 千 / 百 / 十 groups -- it simply rewrites each ASCII
// digit through a fixed table. Concretely:
//
//   * `=TEXT(1234, "[DBNum1]0")` -> `一二三四` (NOT `一千二百三十四`).
//   * `=TEXT(1234, "[DBNum2]0")` -> `壱弐参四` (NOT `壱阡弐百参拾四`).
//   * `=TEXT(1234, "[DBNum3]0")` -> `１２３４` (full-width Arabic).

// Returns the per-digit substitution for `c` under `mode`, or an empty
// string if no substitution applies (caller falls back to `c` verbatim).
std::string_view dbnum_digit_subst(DbNumMode mode, char c) noexcept;

// Append a single ASCII digit `c`, substituting it via `mode` if applicable.
// Non-digit characters fall through verbatim.
void append_digit_dbnum(std::string& out, DbNumMode mode, char c);

// Append every character of `chars`, substituting digits via `mode`.
void append_chars_dbnum(std::string& out, DbNumMode mode, std::string_view chars);

// Appends `value` to `out` with each digit substituted per `mode`. Used
// for date components (era year, m, d, h, min, s) where positional kanji
// do NOT apply -- only per-digit substitution.
void append_int_dbnum(std::string& out, long long value, DbNumMode mode);

// Append `n` zero-padded to two characters without any DBNum substitution
// (ASCII output).
void append_pad2(std::string& out, unsigned n);

// Appends `value` zero-padded to 2 digits, with DBNum substitution applied.
void append_pad2_dbnum(std::string& out, unsigned value, DbNumMode mode);

// True when every byte in `digits` is the ASCII '0' character. Used by
// the numeric walker to suppress a stray minus sign when a tiny negative
// magnitude rounds down to a representation of zero.
bool decimal_digits_all_zero(std::string_view digits) noexcept;

}  // namespace number_format_detail
}  // namespace text_format
}  // namespace formulon

#endif  // FORMULON_EVAL_TEXT_FORMAT_RENDER_COMMON_H_
