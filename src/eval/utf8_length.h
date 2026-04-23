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

/// Returns the Excel DBCS byte count that Mac Excel 365 (ja-JP) assigns to a
/// single Unicode codepoint for the LENB / LEFTB / RIGHTB / MIDB family.
///
/// * ASCII (U+0000..U+007F) -> 1 byte.
/// * Half-width katakana (U+FF61..U+FF9F) -> 1 byte.
/// * All other codepoints (BMP and supplementary plane) -> 2 bytes.
///
/// This classification matches Mac Excel ja-JP's observed behaviour:
/// supplementary-plane emoji such as "😀" count as 2 bytes, not 4
/// (oracle-verified 2026-04-23). Any future deviation must be captured in
/// `tests/divergence.yaml`.
int byte_count_jajp(std::uint32_t codepoint) noexcept;

/// Returns the Excel DBCS byte count of `s` interpreted as UTF-8 under the
/// ja-JP rule above. Sum of `byte_count_jajp(cp)` for every decoded codepoint.
/// Malformed leading bytes contribute 1 byte each (mirroring
/// `utf16_units_in`).
std::uint64_t bytes_in_jajp(std::string_view s) noexcept;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_UTF8_LENGTH_H_
