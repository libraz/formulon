// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the UTF-8 / UTF-16 helpers declared in `text_ops.h`.
// The codepoint walker mirrors `utf8_length.cpp`: we read the leading byte,
// derive the byte length (1/2/3/4) and the UTF-16 unit count (1 for byte_len
// <= 3, 2 for byte_len == 4). Malformed leading bytes and truncated
// sequences each consume one byte and one UTF-16 unit, matching the
// behaviour of `utf16_units_in`.

#include "eval/text_ops.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

namespace formulon {
namespace eval {
namespace {

// Decodes the leading byte at `text[i]` into a byte length and UTF-16 unit
// length. On a malformed or truncated sequence, sets both lengths to 1 so
// the caller advances by a single byte / unit (matching `utf16_units_in`).
struct Step {
  std::size_t byte_len;
  std::uint32_t unit_len;
};

Step next_step(std::string_view text, std::size_t i) noexcept {
  const auto c0 = static_cast<unsigned char>(text[i]);
  if (c0 < 0x80) {
    return {1, 1};
  }
  std::size_t need = 0;
  if ((c0 & 0xE0) == 0xC0) {
    need = 1;
  } else if ((c0 & 0xF0) == 0xE0) {
    need = 2;
  } else if ((c0 & 0xF8) == 0xF0) {
    need = 3;
  } else {
    return {1, 1};
  }
  if (i + need >= text.size()) {
    return {1, 1};
  }
  for (std::size_t k = 0; k < need; ++k) {
    const auto ck = static_cast<unsigned char>(text[i + 1 + k]);
    if ((ck & 0xC0) != 0x80) {
      return {1, 1};
    }
  }
  // need == 3 means a 4-byte UTF-8 sequence -> supplementary plane -> 2 units.
  return {need + 1, need == 3 ? 2u : 1u};
}

}  // namespace

std::size_t utf16_to_byte_offset(std::string_view text, std::uint32_t units_offset) noexcept {
  if (units_offset == 0) {
    return 0;
  }
  std::uint32_t units = 0;
  std::size_t i = 0;
  while (i < text.size()) {
    const Step step = next_step(text, i);
    units += step.unit_len;
    i += step.byte_len;
    // Rounding-up rule: if we just crossed (or hit) the target, return the
    // byte position AFTER the codepoint. For a surrogate-pair midpoint this
    // yields the byte position past the entire supplementary codepoint.
    if (units >= units_offset) {
      return i;
    }
  }
  return text.size();
}

std::string utf16_substring(std::string_view text, std::uint32_t start_units, std::uint32_t length_units) {
  const std::size_t start_byte = utf16_to_byte_offset(text, start_units);
  // Saturating add on `start_units + length_units` to avoid overflow when the
  // caller passes very large bounds (e.g. a text-length-derived end).
  std::uint32_t end_units = start_units + length_units;
  if (end_units < start_units) {
    end_units = UINT32_MAX;
  }
  const std::size_t end_byte = utf16_to_byte_offset(text, end_units);
  if (end_byte <= start_byte) {
    return std::string();
  }
  return std::string(text.substr(start_byte, end_byte - start_byte));
}

std::string to_lower_ascii(std::string_view text) {
  std::string out(text);
  for (char& c : out) {
    if (c >= 'A' && c <= 'Z') {
      c = static_cast<char>(c + 32);
    }
  }
  return out;
}

std::string to_upper_ascii(std::string_view text) {
  std::string out(text);
  for (char& c : out) {
    if (c >= 'a' && c <= 'z') {
      c = static_cast<char>(c - 32);
    }
  }
  return out;
}

std::string encode_utf8_codepoint(std::uint32_t codepoint) {
  // Invalid codepoints (beyond Unicode max or in the surrogate range) yield
  // an empty result. Caller is expected to validate before calling, but we
  // double-guard so this helper never emits malformed UTF-8.
  if (codepoint > 0x10FFFFu || (codepoint >= 0xD800u && codepoint <= 0xDFFFu)) {
    return std::string();
  }
  std::string out;
  if (codepoint < 0x80u) {
    out.push_back(static_cast<char>(codepoint));
  } else if (codepoint < 0x800u) {
    out.push_back(static_cast<char>(0xC0u | (codepoint >> 6)));
    out.push_back(static_cast<char>(0x80u | (codepoint & 0x3Fu)));
  } else if (codepoint < 0x10000u) {
    out.push_back(static_cast<char>(0xE0u | (codepoint >> 12)));
    out.push_back(static_cast<char>(0x80u | ((codepoint >> 6) & 0x3Fu)));
    out.push_back(static_cast<char>(0x80u | (codepoint & 0x3Fu)));
  } else {
    out.push_back(static_cast<char>(0xF0u | (codepoint >> 18)));
    out.push_back(static_cast<char>(0x80u | ((codepoint >> 12) & 0x3Fu)));
    out.push_back(static_cast<char>(0x80u | ((codepoint >> 6) & 0x3Fu)));
    out.push_back(static_cast<char>(0x80u | (codepoint & 0x3Fu)));
  }
  return out;
}

Utf8DecodeResult decode_first_utf8_codepoint(std::string_view text) noexcept {
  if (text.empty()) {
    return {false, 0u, 0u};
  }
  const auto c0 = static_cast<unsigned char>(text[0]);
  if (c0 < 0x80u) {
    return {true, static_cast<std::uint32_t>(c0), 1u};
  }
  std::size_t need = 0;
  std::uint32_t value = 0;
  if ((c0 & 0xE0u) == 0xC0u) {
    need = 1;
    value = c0 & 0x1Fu;
  } else if ((c0 & 0xF0u) == 0xE0u) {
    need = 2;
    value = c0 & 0x0Fu;
  } else if ((c0 & 0xF8u) == 0xF0u) {
    need = 3;
    value = c0 & 0x07u;
  } else {
    return {false, 0u, 0u};
  }
  if (text.size() < need + 1) {
    return {false, 0u, 0u};
  }
  for (std::size_t k = 0; k < need; ++k) {
    const auto ck = static_cast<unsigned char>(text[k + 1]);
    if ((ck & 0xC0u) != 0x80u) {
      return {false, 0u, 0u};
    }
    value = (value << 6) | (ck & 0x3Fu);
  }
  return {true, value, need + 1};
}

}  // namespace eval
}  // namespace formulon
