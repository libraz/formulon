//
// Bit layout of the flag fields of the worksheet records BrtTable,
// BrtRowHdr and BrtColInfo. The reader and the writer take their masks from
// here so the two directions cannot disagree about where a field lives.

#ifndef FORMULON_IO_XLSB_SHEET_RECORD_FLAGS_H_
#define FORMULON_IO_XLSB_SHEET_RECORD_FLAGS_H_

#include <cstdint>

namespace formulon {
namespace io {
namespace xlsb {

/// `BrtTable` flag byte.
constexpr std::uint8_t kTableRowInput = 0x01U;
constexpr std::uint8_t kTableTwoDimensional = 0x02U;
constexpr std::uint8_t kTableDeletedInput1 = 0x04U;
constexpr std::uint8_t kTableDeletedInput2 = 0x08U;
constexpr std::uint8_t kTableAlwaysCalc = 0x10U;

/// `BrtRowHdr` second flag byte.
constexpr std::uint8_t kRowOutlineLevelMask = 0x07U;
constexpr std::uint8_t kRowCollapsed = 0x08U;
constexpr std::uint8_t kRowZeroHeight = 0x10U;  ///< fDyZero: the row is hidden.
constexpr std::uint8_t kRowUnsynced = 0x20U;    ///< Custom height.
constexpr std::uint8_t kRowGhostDirty = 0x40U;  ///< `iStyleRef` is an effective row style.

/// `BrtColInfo` flag word.
constexpr std::uint16_t kColHidden = 0x0001U;
constexpr std::uint16_t kColUserSet = 0x0002U;
constexpr unsigned kColOutlineShift = 8U;
constexpr std::uint16_t kColOutlineLevelMask = 0x07U;
constexpr std::uint16_t kColCollapsed = 0x0800U;

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_SHEET_RECORD_FLAGS_H_
