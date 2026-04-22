// Copyright 2026 libraz. Licensed under the MIT License.
//
// Pure UTF-8 / UTF-16 helpers used by Formulon's text built-ins. Excel
// measures every text position and length in UTF-16 code units, while we
// store strings as UTF-8 byte sequences. The functions in this header are
// the bridge: they convert between UTF-16 unit offsets and UTF-8 byte
// offsets, slice substrings on UTF-16 boundaries, and apply ASCII case
// folding without pulling in a locale layer.
//
// This translation unit deliberately depends only on the standard library:
// no `Value`, no `Arena`, no parser headers. That keeps the helpers reusable
// from both the evaluator and any future text-rendering code, and keeps the
// dependency graph tight for the WASM size budget.

#ifndef FORMULON_EVAL_TEXT_OPS_H_
#define FORMULON_EVAL_TEXT_OPS_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

namespace formulon {
namespace eval {

/// Returns the byte offset in `text` that begins at UTF-16 unit `units_offset`.
/// `units_offset` is clamped to `utf16_units_in(text)`. The returned byte offset
/// always aligns to a UTF-8 codepoint boundary (never lands inside a multi-byte
/// sequence).
///
/// If `units_offset` falls between the two halves of a surrogate pair (i.e. the
/// caller asked for an offset that splits a supplementary-plane codepoint), the
/// returned byte offset is positioned AFTER the full codepoint. This matches
/// Excel's behavior — Excel rounds up to the next codepoint boundary in the
/// rare case a slice would split a surrogate pair.
std::size_t utf16_to_byte_offset(std::string_view text, std::uint32_t units_offset) noexcept;

/// Returns a substring of `text` covering UTF-16 units [start_units, start_units + length_units).
/// Both bounds are clamped to the text's UTF-16 length. Returns an owned `std::string`
/// because the result may differ in byte length from any view into `text`.
std::string utf16_substring(std::string_view text, std::uint32_t start_units, std::uint32_t length_units);

/// In-place ASCII case folding. Bytes 'A'..'Z' map to 'a'..'z' (and vice versa
/// for `to_upper_ascii`). All other bytes — including UTF-8 continuation bytes
/// and non-ASCII leading bytes — are unchanged. This is the MVP behavior; Excel
/// uses Unicode case folding which we will implement when the locale layer lands.
std::string to_lower_ascii(std::string_view text);
std::string to_upper_ascii(std::string_view text);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_TEXT_OPS_H_
