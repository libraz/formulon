//
// Implementation of `utf16_units_in`. See `utf8_length.h` for the contract.

#include "eval/utf8_length.h"

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "eval/text_ops.h"

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
    std::size_t step = 0;
    const std::uint32_t cp = decode_utf8_step(s, i, &step);
    // Lenient decoder reports U+FFFD with a 1-byte advance on malformed
    // input. Both ASCII and replacement-codepoint paths cost 1 DBCS byte
    // (`byte_count_jajp(0xFFFD) == 2` would mis-count broken input as
    // 2 bytes; the prior implementation explicitly counted these as 1).
    if (step == 1 && cp == 0xFFFDu) {
      bytes += 1;
      i += 1;
      continue;
    }
    bytes += static_cast<std::uint64_t>(byte_count_jajp(cp));
    i += step;
  }
  return bytes;
}

std::uint32_t utf16_units_in(std::string_view s) noexcept {
  std::uint32_t units = 0;
  std::size_t i = 0;
  while (i < s.size()) {
    std::size_t step = 0;
    const std::uint32_t cp = decode_utf8_step(s, i, &step);
    // Surrogate-pair codepoints (>0xFFFF) take two UTF-16 units; everything
    // else (including the U+FFFD replacement emitted on malformed input
    // since it's <= 0xFFFF) takes a single unit. This matches the prior
    // line-by-line implementation byte-for-byte.
    units += cp > 0xFFFFu ? 2u : 1u;
    i += step;
  }
  return units;
}

}  // namespace eval
}  // namespace formulon
