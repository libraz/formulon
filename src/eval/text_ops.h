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

/// Encodes a single Unicode codepoint as a UTF-8 byte sequence (1-4 bytes).
///
/// The caller is responsible for validating `codepoint` (rejecting values
/// outside [0, 0x10FFFF] and the surrogate range [0xD800, 0xDFFF]) before
/// invoking this helper. When passed an invalid codepoint, the function
/// returns an empty string. This signature mirrors the helpers above:
/// pure UTF-8 in, owned `std::string` out, no `Value`/`Arena` dependencies.
std::string encode_utf8_codepoint(std::uint32_t codepoint);

/// Result of decoding the first UTF-8 codepoint in a byte sequence.
/// `valid` is false when `text` is empty or the leading bytes form a
/// malformed / truncated sequence.
struct Utf8DecodeResult {
  bool valid;
  std::uint32_t codepoint;
  std::size_t byte_len;
};

/// Decodes the first UTF-8 codepoint in `text`. On a malformed leading byte,
/// truncated continuation, or empty input, returns `{false, 0, 0}`. Used by
/// UNICODE() to read the leading codepoint of a string.
Utf8DecodeResult decode_first_utf8_codepoint(std::string_view text) noexcept;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_TEXT_OPS_H_
