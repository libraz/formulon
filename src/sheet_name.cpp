// Worksheet-name UTF-8 decoding and Unicode simple-fold identity.

#include "sheet_name.h"

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "sheet_name_casefold_data.h"

namespace formulon {
namespace sheet_names {
namespace {

constexpr bool is_continuation(std::uint8_t byte) noexcept {
  return (byte & 0xC0U) == 0x80U;
}

// Decodes one scalar at `*offset`, applying the strict UTF-8 constraints from
// RFC 3629. On failure, the offset is left at the malformed lead byte so the
// caller can terminate without reading past the input.
bool decode_one(std::string_view text, std::size_t* offset, std::uint32_t* scalar) noexcept {
  const std::size_t start = *offset;
  if (start >= text.size()) {
    return false;
  }
  const std::uint8_t first = static_cast<std::uint8_t>(text[start]);
  if (first < 0x80U) {
    *scalar = first;
    *offset = start + 1U;
    return true;
  }

  std::size_t length = 0;
  std::uint32_t value = 0;
  std::uint32_t minimum = 0;
  if (first >= 0xC2U && first <= 0xDFU) {
    length = 2U;
    value = first & 0x1FU;
    minimum = 0x80U;
  } else if (first >= 0xE0U && first <= 0xEFU) {
    length = 3U;
    value = first & 0x0FU;
    minimum = 0x800U;
  } else if (first >= 0xF0U && first <= 0xF4U) {
    length = 4U;
    value = first & 0x07U;
    minimum = 0x10000U;
  } else {
    return false;
  }

  if (length > text.size() - start) {
    return false;
  }
  for (std::size_t i = 1U; i < length; ++i) {
    const std::uint8_t byte = static_cast<std::uint8_t>(text[start + i]);
    if (!is_continuation(byte)) {
      return false;
    }
    value = (value << 6U) | static_cast<std::uint32_t>(byte & 0x3FU);
  }
  if (value < minimum || value > 0x10FFFFU || (value >= 0xD800U && value <= 0xDFFFU)) {
    return false;
  }

  // Tighten the lead/second-byte boundaries that are not captured by the
  // scalar-range check: E0 must not encode below U+0800, ED must not encode a
  // surrogate, F0 must not encode below U+10000, and F4 must stay <= U+10FFFF.
  const std::uint8_t second = static_cast<std::uint8_t>(text[start + 1U]);
  if ((first == 0xE0U && second < 0xA0U) || (first == 0xEDU && second > 0x9FU) || (first == 0xF0U && second < 0x90U) ||
      (first == 0xF4U && second > 0x8FU)) {
    return false;
  }

  *scalar = value;
  *offset = start + length;
  return true;
}

}  // namespace

bool valid_utf8(std::string_view name) noexcept {
  std::size_t offset = 0;
  std::uint32_t scalar = 0;
  while (offset < name.size()) {
    if (!decode_one(name, &offset, &scalar)) {
      return false;
    }
  }
  return true;
}

bool equal(std::string_view lhs, std::string_view rhs) noexcept {
  if (!valid_utf8(lhs) || !valid_utf8(rhs)) {
    return false;
  }
  if (lhs == rhs) {
    return true;
  }
  std::size_t lhs_offset = 0;
  std::size_t rhs_offset = 0;
  while (lhs_offset < lhs.size() && rhs_offset < rhs.size()) {
    std::uint32_t lhs_scalar = 0;
    std::uint32_t rhs_scalar = 0;
    if (!decode_one(lhs, &lhs_offset, &lhs_scalar) || !decode_one(rhs, &rhs_offset, &rhs_scalar)) {
      return false;
    }
    if (detail::fold_scalar(lhs_scalar) != detail::fold_scalar(rhs_scalar)) {
      return false;
    }
  }
  if (lhs_offset != lhs.size() || rhs_offset != rhs.size()) {
    return false;
  }
  return true;
}

}  // namespace sheet_names
}  // namespace formulon
