// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of `utf16_units_in`. See `utf8_length.h` for the contract.

#include "eval/utf8_length.h"

#include <cstdint>
#include <string_view>

namespace formulon {
namespace eval {

int byte_count_jajp(std::uint32_t codepoint) noexcept {
  if (codepoint <= 0x7Fu) {
    return 1;
  }
  if (codepoint >= 0xFF61u && codepoint <= 0xFF9Fu) {
    // Half-width katakana block (single-byte region in Shift-JIS / CP932).
    return 1;
  }
  // All other codepoints (BMP and supplementary plane) count as 2 bytes.
  // Oracle-verified on Mac Excel ja-JP: LENB("\U0001F600") == 2.
  return 2;
}

std::uint64_t bytes_in_jajp(std::string_view s) noexcept {
  std::uint64_t bytes = 0;
  std::size_t i = 0;
  while (i < s.size()) {
    const auto c0 = static_cast<unsigned char>(s[i]);
    if (c0 < 0x80) {
      bytes += 1;
      i += 1;
      continue;
    }
    std::uint32_t need = 0;
    std::uint32_t value = 0;
    if ((c0 & 0xE0) == 0xC0) {
      need = 1;
      value = c0 & 0x1F;
    } else if ((c0 & 0xF0) == 0xE0) {
      need = 2;
      value = c0 & 0x0F;
    } else if ((c0 & 0xF8) == 0xF0) {
      need = 3;
      value = c0 & 0x07;
    } else {
      // Malformed leading byte: count as a single DBCS byte and advance one.
      bytes += 1;
      i += 1;
      continue;
    }
    if (i + need >= s.size()) {
      bytes += 1;
      i += 1;
      continue;
    }
    bool ok = true;
    for (std::uint32_t k = 0; k < need; ++k) {
      const auto ck = static_cast<unsigned char>(s[i + 1 + k]);
      if ((ck & 0xC0) != 0x80) {
        ok = false;
        break;
      }
      value = (value << 6) | (ck & 0x3F);
    }
    if (!ok) {
      bytes += 1;
      i += 1;
      continue;
    }
    bytes += static_cast<std::uint64_t>(byte_count_jajp(value));
    i += need + 1;
  }
  return bytes;
}

std::uint32_t utf16_units_in(std::string_view s) noexcept {
  std::uint32_t units = 0;
  std::size_t i = 0;
  while (i < s.size()) {
    const auto c0 = static_cast<unsigned char>(s[i]);
    if (c0 < 0x80) {
      units += 1;
      i += 1;
      continue;
    }
    std::uint32_t need = 0;
    std::uint32_t value = 0;
    if ((c0 & 0xE0) == 0xC0) {
      need = 1;
      value = c0 & 0x1F;
    } else if ((c0 & 0xF0) == 0xE0) {
      need = 2;
      value = c0 & 0x0F;
    } else if ((c0 & 0xF8) == 0xF0) {
      need = 3;
      value = c0 & 0x07;
    } else {
      // Malformed leading byte: count as one unit and skip a single byte.
      units += 1;
      i += 1;
      continue;
    }
    if (i + need >= s.size()) {
      // Truncated multibyte sequence: count one unit and skip one byte.
      units += 1;
      i += 1;
      continue;
    }
    bool ok = true;
    for (std::uint32_t k = 0; k < need; ++k) {
      const auto ck = static_cast<unsigned char>(s[i + 1 + k]);
      if ((ck & 0xC0) != 0x80) {
        ok = false;
        break;
      }
      value = (value << 6) | (ck & 0x3F);
    }
    if (!ok) {
      units += 1;
      i += 1;
      continue;
    }
    units += value > 0xFFFF ? 2 : 1;
    i += need + 1;
  }
  return units;
}

}  // namespace eval
}  // namespace formulon
