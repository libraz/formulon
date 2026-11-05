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
#include <vector>

#include "cf/cf_types.h"
#include "io/xlsb/brt_color.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/guid_text.h"
#include "io/xlsb/record.h"
#include "utils/number_text.h"

namespace formulon {
namespace io {
namespace xlsb {

// BrtFRTBegin payload Excel writes around the x14 link, and the
// BrtCFRuleExt layout: a zero u32, then the GUID in Windows byte order.
inline constexpr std::array<std::uint8_t, 4> kFrtVersion = {0x02, 0x0E, 0x00, 0x00};

/// A future-record container `open` .. `close` around `body`, wrapped in
/// BrtFRTBegin / BrtFRTEnd. Excel refuses an x14 container that directly
/// follows the last legacy block, so a sheet whose `recs` carry no print
/// records gets Excel's default BrtPrintOptions ahead of it.
std::vector<std::uint8_t> frt_container_block(const std::vector<FramedRecord>& recs, std::uint16_t open,
                                              std::uint16_t close, const std::vector<std::uint8_t>& body);

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

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_CF_X14_LINK_H_
