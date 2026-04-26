// Copyright 2026 libraz. Licensed under the MIT License.
//
// Hand-written lookup helpers backing the JIS X 0208 reverse table generated
// into `jis0208_table.cpp`. Kept in a separate translation unit so that the
// generator (`tools/jis0208/generate_table.py`) can clobber the data file
// without touching this code.

#include <cstddef>
#include <cstdint>

#include "eval/jis0208_table.h"

namespace formulon {
namespace eval {

std::uint16_t lookup_unicode_to_jis0208(std::uint32_t unicode_codepoint) {
  // U+0000 is the sentinel for "unassigned" and is never mapped, so reject it
  // up front. Codepoints outside the BMP cannot appear in JIS X 0208 either.
  if (unicode_codepoint == 0u || unicode_codepoint > 0xFFFFu) {
    return 0u;
  }
  const auto needle = static_cast<std::uint16_t>(unicode_codepoint);
  constexpr std::size_t kSize = kJis0208RowCount * kJis0208CellCount;
  for (std::size_t i = 0; i < kSize; ++i) {
    if (kJis0208Reverse[i] == needle) {
      const std::size_t row = i / kJis0208CellCount;   // 0-based
      const std::size_t cell = i % kJis0208CellCount;  // 0-based
      return static_cast<std::uint16_t>(((row + 0x21u) << 8) | (cell + 0x21u));
    }
  }
  return 0u;
}

std::uint16_t lookup_jis0208_to_unicode(std::uint8_t jis_hi_byte, std::uint8_t jis_lo_byte) {
  if (jis_hi_byte < 0x21u || jis_hi_byte > 0x7Eu || jis_lo_byte < 0x21u || jis_lo_byte > 0x7Eu) {
    return 0u;
  }
  const std::size_t row = static_cast<std::size_t>(jis_hi_byte - 0x21u);   // 0-based
  const std::size_t cell = static_cast<std::size_t>(jis_lo_byte - 0x21u);  // 0-based
  return kJis0208Reverse[row * kJis0208CellCount + cell];
}

}  // namespace eval
}  // namespace formulon
