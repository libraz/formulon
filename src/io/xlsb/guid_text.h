//
// `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}` text of a [MS-DTYP] GUID and its
// 16-byte wire form (Data1..Data3 little-endian, Data4 a byte string).

#ifndef FORMULON_IO_XLSB_GUID_TEXT_H_
#define FORMULON_IO_XLSB_GUID_TEXT_H_

#include <array>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "utils/strings.h"

namespace formulon {
namespace io {
namespace xlsb {

inline constexpr std::size_t kGuidBytes = 16;

namespace guid_detail {
/// Wire byte index per text byte; -1 marks a '-' separator.
inline constexpr int kTextOrder[] = {3, 2, 1, 0, -1, 5, 4, -1, 7, 6, -1, 8, 9, -1, 10, 11, 12, 13, 14, 15};
}  // namespace guid_detail

/// Braced upper-case text of the GUID at `b` (16 bytes).
inline std::string FormatGuid(const std::uint8_t* b) {
  std::string out("{");
  for (const int i : guid_detail::kTextOrder) {
    if (i < 0) {
      out.push_back('-');
      continue;
    }
    out.push_back(strings::hex_digit_upper(b[i] >> 4U));
    out.push_back(strings::hex_digit_upper(b[i]));
  }
  out.push_back('}');
  return out;
}

/// Parses braced GUID text (either hex case) into wire bytes; false on any
/// other shape.
inline bool ParseGuid(std::string_view text, std::array<std::uint8_t, kGuidBytes>& out) {
  if (text.size() != 38U || text.front() != '{' || text.back() != '}') {
    return false;
  }
  std::array<std::uint8_t, kGuidBytes> bytes{};
  std::size_t pos = 1;
  for (const int i : guid_detail::kTextOrder) {
    if (i < 0) {
      if (text[pos++] != '-') {
        return false;
      }
      continue;
    }
    const int hi = strings::hex_digit_value(text[pos]);
    const int lo = strings::hex_digit_value(text[pos + 1U]);
    if (hi < 0 || lo < 0) {
      return false;
    }
    bytes[static_cast<std::size_t>(i)] = static_cast<std::uint8_t>((hi << 4) | lo);
    pos += 2U;
  }
  out = bytes;
  return true;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_GUID_TEXT_H_
