// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of `utf16_units_in`. See `utf8_length.h` for the contract.

#include "eval/utf8_length.h"

#include <cstdint>
#include <string_view>

namespace formulon {
namespace eval {

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
