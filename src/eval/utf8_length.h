// Copyright 2026 libraz. Licensed under the MIT License.
//
// Tiny standalone helper that walks a UTF-8 byte sequence and returns the
// equivalent UTF-16 code-unit count. Used by `LEN` and other text functions
// that report length in Excel's UTF-16-unit semantics.

#ifndef FORMULON_EVAL_UTF8_LENGTH_H_
#define FORMULON_EVAL_UTF8_LENGTH_H_

#include <cstdint>
#include <string_view>

namespace formulon {
namespace eval {

/// Returns the number of UTF-16 code units required to represent `s` after
/// decoding it as UTF-8.
///
/// * BMP codepoints (U+0000..U+FFFF) contribute 1 unit each.
/// * Supplementary-plane codepoints (U+10000..U+10FFFF) contribute 2 units
///   each (surrogate pair).
/// * Malformed leading bytes contribute 1 unit each, matching how the
///   tokenizer's `peek_codepoint` treats them.
std::uint32_t utf16_units_in(std::string_view s) noexcept;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_UTF8_LENGTH_H_
