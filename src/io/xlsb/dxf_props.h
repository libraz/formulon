//
// Record ids and property types of the XLSB differential-format block
// (`BrtBeginDxfs` / `BrtDXF` / `BrtEndDxfs`), shared by the styles reader and
// writer, plus the BrtFill gradient constants the dxf gradient mirrors.
// A `BrtDXF` payload is `u16 flags, u16 reserved, u16 count` followed
// by `count` properties, each `u16 type, u16 size (header included), data`.

#ifndef FORMULON_IO_XLSB_DXF_PROPS_H_
#define FORMULON_IO_XLSB_DXF_PROPS_H_

#include <cstddef>
#include <cstdint>

namespace formulon {
namespace io {
namespace xlsb {

inline constexpr std::uint16_t kBrtBeginDxfs = 505;
inline constexpr std::uint16_t kBrtEndDxfs = 506;
inline constexpr std::uint16_t kBrtDxf = 507;

/// BrtDXF property types as Excel saves them (each matches the
/// <dxf> Excel wrote for the same book).
inline constexpr std::uint16_t kDxfPropFillPattern = 0;
inline constexpr std::uint16_t kDxfPropFillFg = 1;
inline constexpr std::uint16_t kDxfPropFillBg = 2;
/// XFPropGradient: i32 type (0 linear, 1 rectangular), then degree, left,
/// right, top, bottom as Xnum doubles ([MS-XLSB] 2.5.162).
inline constexpr std::uint16_t kDxfPropFillGradient = 3;
/// XFPropGradientStop: u16 unused, Xnum position, XFPropColor (2.5.163).
inline constexpr std::uint16_t kDxfPropGradientStop = 4;
inline constexpr std::uint16_t kDxfPropFontColor = 5;
inline constexpr std::uint16_t kDxfPropBorderTop = 6;
inline constexpr std::uint16_t kDxfPropBorderBottom = 7;
inline constexpr std::uint16_t kDxfPropBorderLeft = 8;
inline constexpr std::uint16_t kDxfPropBorderRight = 9;
inline constexpr std::uint16_t kDxfPropBorderDiagonal = 10;
inline constexpr std::uint16_t kDxfPropDiagonalUp = 13;
inline constexpr std::uint16_t kDxfPropDiagonalDown = 14;
inline constexpr std::uint16_t kDxfPropFontName = 24;
inline constexpr std::uint16_t kDxfPropFontWeight = 25;
inline constexpr std::uint16_t kDxfPropUnderline = 26;
inline constexpr std::uint16_t kDxfPropVertAlign = 27;
inline constexpr std::uint16_t kDxfPropItalic = 28;
inline constexpr std::uint16_t kDxfPropStrike = 29;
inline constexpr std::uint16_t kDxfPropOutline = 30;
inline constexpr std::uint16_t kDxfPropShadow = 31;
inline constexpr std::uint16_t kDxfPropCondense = 32;
inline constexpr std::uint16_t kDxfPropExtend = 33;
inline constexpr std::uint16_t kDxfPropFontSize = 36;
inline constexpr std::uint16_t kDxfPropNumFmtCode = 38;
inline constexpr std::uint16_t kDxfPropNumFmtId = 41;

/// BrtFill (45) `fls` of a gradient fill ([MS-XLSB] 2.4.681).
inline constexpr std::uint32_t kBrtFillGradient = 0x28U;
/// `cNumStop` ceiling of BrtFill; XFProps uses the same bound.
inline constexpr std::size_t kMaxBrtGradientStops = 256U;

/// `FontRecord::underline` ordinal -> the wire underline value.
inline constexpr std::uint8_t kDxfUnderlineWire[] = {0U, 1U, 2U, 0x21U, 0x22U};

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_DXF_PROPS_H_
