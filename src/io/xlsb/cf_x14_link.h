// Shared pieces of the BrtBeginCFRule x14 link: the FRT block framing, the
// BrtColor and GUID codecs used by both the legacy rule records and the x14
// data-bar records.

#ifndef FORMULON_IO_XLSB_CF_X14_LINK_H_
#define FORMULON_IO_XLSB_CF_X14_LINK_H_

#include <array>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>

#include "cf/cf_types.h"
#include "io/xlsb/brt_color.h"
#include "io/xlsb/record.h"
#include "utils/number_text.h"

namespace formulon {
namespace io {
namespace xlsb {

inline constexpr std::uint16_t kFrtBegin = 35;
inline constexpr std::uint16_t kFrtEnd = 36;

// BrtFRTBegin payload Excel writes around the x14 link, and the
// BrtCFRuleExt layout: a zero u32, then the GUID in Windows byte order.
inline constexpr std::array<std::uint8_t, 4> kFrtVersion = {0x02, 0x0E, 0x00, 0x00};
inline constexpr std::size_t kGuidBytes = 16;

// BrtColor: [flags (fValidRGB bit 0, XColorType bits 1..7), index, tint:i16,
// r, g, b, a]. XColorType 0 is automatic, 1 a palette index, 2 literal RGB,
// 3 a theme index with tint (tint / 32767). Theme, palette and automatic
// colours keep their selector in `Color::spec`; the channels stay the
// placeholder the OOXML reader leaves.

inline std::optional<cf::Color> DecodeColor(ByteSpan p) {
  if (p.size != 8U) {
    return std::nullopt;
  }
  const std::uint8_t type = static_cast<std::uint8_t>(p.data[0] >> 1U);
  const std::uint8_t index = p.data[1];
  cf::Color out{};
  switch (type) {
    case kBrtColorTypeRgb:
      if (p.data[0] != 0x05U || index != 0xFFU || p.data[2] != 0U || p.data[3] != 0U) {
        return std::nullopt;
      }
      return cf::Color{p.data[4], p.data[5], p.data[6], p.data[7]};
    case kBrtColorTypeTheme: {
      out.spec.kind = ColorSpec::Kind::kTheme;
      out.spec.theme = index;
      const auto raw = static_cast<std::int16_t>(static_cast<std::uint16_t>(p.data[2] | (p.data[3] << 8U)));
      out.spec.tint = static_cast<double>(raw) / kBrtColorTintScale;
      return out;
    }
    case kBrtColorTypeIndexed:
      out.spec.kind = ColorSpec::Kind::kIndexed;
      out.spec.indexed = index;
      return out;
    case kBrtColorTypeAuto:
      out.spec.kind = ColorSpec::Kind::kAuto;
      return out;
    default:
      return std::nullopt;
  }
}

inline std::array<std::uint8_t, 8> EncodeColor(const cf::Color& c) {
  const std::uint32_t argb = (static_cast<std::uint32_t>(c.a) << 24U) | (static_cast<std::uint32_t>(c.r) << 16U) |
                             (static_cast<std::uint32_t>(c.g) << 8U) | static_cast<std::uint32_t>(c.b);
  return encode_brt_color(argb, c.spec);
}

inline std::string FormatGuid(const std::uint8_t* g) {
  // Bytes in GUID text order: the first three groups are stored little-endian.
  static constexpr int kTextOrder[16] = {3, 2, 1, 0, 5, 4, 7, 6, 8, 9, 10, 11, 12, 13, 14, 15};
  std::string out = "{";
  for (int i = 0; i < 16; ++i) {
    if (i == 4 || i == 6 || i == 8 || i == 10) {
      out.push_back('-');
    }
    char hex[4];
    format_hex(hex, sizeof(hex), g[kTextOrder[i]], 2, true);
    out.append(hex);
  }
  out.push_back('}');
  return out;
}

inline bool ParseGuid(const std::string& id, std::array<std::uint8_t, kGuidBytes>& out) {
  // {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}
  if (id.size() != 38U || id.front() != '{' || id.back() != '}') {
    return false;
  }
  std::array<std::uint8_t, kGuidBytes> be{};
  std::size_t n = 0;
  for (std::size_t i = 1; i + 1U < id.size(); ++i) {
    const char c = id[i];
    if (c == '-') {
      if (i != 9U && i != 14U && i != 19U && i != 24U) {
        return false;
      }
      continue;
    }
    const int d = (c >= '0' && c <= '9')   ? c - '0'
                  : (c >= 'A' && c <= 'F') ? c - 'A' + 10
                  : (c >= 'a' && c <= 'f') ? c - 'a' + 10
                                           : -1;
    if (d < 0 || n >= 2U * kGuidBytes) {
      return false;
    }
    be[n / 2U] = static_cast<std::uint8_t>((be[n / 2U] << 4U) | static_cast<std::uint8_t>(d));
    ++n;
  }
  if (n != 2U * kGuidBytes) {
    return false;
  }
  // Data1..Data3 are little-endian on the wire; Data4 is a byte string.
  constexpr std::array<std::uint8_t, kGuidBytes> kOrder = {3, 2, 1, 0, 5, 4, 7, 6, 8, 9, 10, 11, 12, 13, 14, 15};
  for (std::size_t i = 0; i < kGuidBytes; ++i) {
    out[i] = be[kOrder[i]];
  }
  return true;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_CF_X14_LINK_H_
