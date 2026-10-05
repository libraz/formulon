//
// BrtColor encoding shared by the styles and conditional-formatting
// writers: [flags (fValidRGB bit 0, XColorType bits 1..7), index,
// tint:i16, r, g, b, a].

#ifndef FORMULON_IO_XLSB_BRT_COLOR_H_
#define FORMULON_IO_XLSB_BRT_COLOR_H_

#include <algorithm>
#include <array>
#include <cmath>
#include <cstdint>

#include "styles.h"

namespace formulon {
namespace io {
namespace xlsb {

/// `XColorType` values.
inline constexpr std::uint8_t kBrtColorTypeAuto = 0;
inline constexpr std::uint8_t kBrtColorTypeIndexed = 1;
inline constexpr std::uint8_t kBrtColorTypeRgb = 2;
inline constexpr std::uint8_t kBrtColorTypeTheme = 3;
/// The system palette slot Excel writes beside an automatic colour.
inline constexpr std::uint8_t kBrtColorAutoIndex = 0x40U;
/// A tint is stored as a signed 16-bit fraction of this scale.
inline constexpr double kBrtColorTintScale = 32767.0;

/// Encodes `spec`, keeping a theme / indexed / auto selector; the RGBA
/// bytes are `argb`, literal for an RGB selector and a fallback otherwise.
/// Bit 7 of the flags byte stays clear: the reserved type 0x41 it would
/// produce makes Excel reject the whole part.
inline std::array<std::uint8_t, 8> encode_brt_color(std::uint32_t argb, const ColorSpec& spec) {
  std::uint8_t type = kBrtColorTypeRgb;
  std::uint8_t index = 0xFFU;  // what Excel writes beside an RGB colour
  std::int16_t tint = 0;
  switch (spec.kind) {
    case ColorSpec::Kind::kTheme:
      type = kBrtColorTypeTheme;
      index = static_cast<std::uint8_t>(std::min<std::uint32_t>(spec.theme, 0xFFU));
      tint = static_cast<std::int16_t>(
          std::clamp(std::round(spec.tint * kBrtColorTintScale), -kBrtColorTintScale, kBrtColorTintScale));
      break;
    case ColorSpec::Kind::kIndexed:
      type = kBrtColorTypeIndexed;
      index = static_cast<std::uint8_t>(std::min<std::uint32_t>(spec.indexed, 0xFFU));
      break;
    case ColorSpec::Kind::kAuto:
      type = kBrtColorTypeAuto;
      index = kBrtColorAutoIndex;
      break;
    case ColorSpec::Kind::kNone:
    case ColorSpec::Kind::kRgb:
      break;
  }
  const auto raw = static_cast<std::uint16_t>(tint);
  return {static_cast<std::uint8_t>((static_cast<unsigned>(type) << 1U) | 0x01U),
          index,
          static_cast<std::uint8_t>(raw & 0xFFU),
          static_cast<std::uint8_t>(raw >> 8U),
          static_cast<std::uint8_t>((argb >> 16U) & 0xFFU),
          static_cast<std::uint8_t>((argb >> 8U) & 0xFFU),
          static_cast<std::uint8_t>(argb & 0xFFU),
          static_cast<std::uint8_t>((argb >> 24U) & 0xFFU)};
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_BRT_COLOR_H_
