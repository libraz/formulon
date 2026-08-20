//
// Bit layout of the two flag words carried by an MS-XLSB `BrtXF` record.
//
// A `BrtXF` payload is 8 x u16:
//
//   [ixfeParent, iFmt, iFont, iFill, ixBorder, trot|indent, flags,
//    xfGrbitAtr]
//
// The sixth word is two u8 fields (text rotation, then indentation); the
// seventh and eighth are the bit fields named below. Both the reader and
// the writer take their masks from here so the two directions cannot
// disagree about where a field lives.
//
// The layout was verified byte-for-byte against Excel-365-produced
// `xl/styles.bin` parts whose `.xlsx` twin spells the same xf table in
// `xl/styles.xml`.
//
// The 3-bit alignment enums use the same ordinals as the OOXML
// `ST_HorizontalAlignment` / `ST_VerticalAlignment` types, so they carry
// straight into `CellXf::horizontal_align` / `CellXf::vertical_align`
// without a mapping table.

#ifndef FORMULON_IO_XLSB_XF_FLAGS_H_
#define FORMULON_IO_XLSB_XF_FLAGS_H_

#include <cstdint>

namespace formulon {
namespace io {
namespace xlsb {

/// `BrtXF::flags` -- alignment, protection and the quote-prefix bit.
constexpr std::uint16_t kXfHorizontalAlignMask = 0x0007U;
constexpr std::uint16_t kXfVerticalAlignMask = 0x0038U;
constexpr unsigned kXfVerticalAlignShift = 3U;
constexpr std::uint16_t kXfWrapText = 0x0040U;
constexpr std::uint16_t kXfJustifyLastLine = 0x0080U;
constexpr std::uint16_t kXfShrinkToFit = 0x0100U;
/// `fMergeCell` and `fSxButton` describe a sheet-level condition (a merge
/// range, a PivotTable drop-down) rather than a cell format, and have no
/// `<xf>` attribute to carry them; they are neither read nor written.
constexpr std::uint16_t kXfReadingOrderMask = 0x0C00U;
constexpr unsigned kXfReadingOrderShift = 10U;
constexpr std::uint16_t kXfLocked = 0x1000U;
constexpr std::uint16_t kXfHidden = 0x2000U;
constexpr std::uint16_t kXfQuotePrefix = 0x8000U;

/// `BrtXF::xfGrbitAtr` -- the `apply*` attribute set. Note that the fill
/// bit sits after the border bit, the reverse of the `<xf>` attribute
/// order.
constexpr std::uint16_t kXfApplyNumberFormat = 0x0001U;
constexpr std::uint16_t kXfApplyFont = 0x0002U;
constexpr std::uint16_t kXfApplyAlignment = 0x0004U;
constexpr std::uint16_t kXfApplyBorder = 0x0008U;
constexpr std::uint16_t kXfApplyFill = 0x0010U;
constexpr std::uint16_t kXfApplyProtection = 0x0020U;

/// `ixfeParent` value a `<cellStyleXfs>` entry carries: such an entry has
/// no parent record to name, and OOXML spells that by omitting `xfId`.
constexpr std::uint16_t kXfNoParent = 0xFFFFU;

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_XF_FLAGS_H_
