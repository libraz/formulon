// Record ids shared by the future-record (FRT / AC) wrappers, the x14
// conditional-formatting, data-validation, sparkline and print-option records
// around them, and the autofilter / table-style records several readers
// dispatch on.

#ifndef FORMULON_IO_XLSB_FRT_RECORDS_H_
#define FORMULON_IO_XLSB_FRT_RECORDS_H_

#include <cstdint>

namespace formulon {
namespace io {
namespace xlsb {

inline constexpr std::uint16_t kFrtBegin = 35;
inline constexpr std::uint16_t kFrtEnd = 36;
inline constexpr std::uint16_t kAcBegin = 37;
inline constexpr std::uint16_t kAcEnd = 38;

inline constexpr std::uint16_t kBeginAFilter = 161;
inline constexpr std::uint16_t kTableStyleClient = 513;

/// BrtTableStyleClient flag bits.
inline constexpr std::uint16_t kStyleClientFirstColumn = 0x01U;
inline constexpr std::uint16_t kStyleClientLastColumn = 0x02U;
inline constexpr std::uint16_t kStyleClientRowStripes = 0x04U;
inline constexpr std::uint16_t kStyleClientColumnStripes = 0x08U;
inline constexpr std::uint16_t kStyleClientRowHeaders = 0x10U;
inline constexpr std::uint16_t kStyleClientColumnHeaders = 0x20U;

inline constexpr std::uint16_t kMargins = 476;
inline constexpr std::uint16_t kPrintOptions = 477;
inline constexpr std::uint16_t kPageSetup = 478;

inline constexpr std::uint16_t kBeginSparklineGroup = 1041;
inline constexpr std::uint16_t kEndSparklineGroup = 1042;
inline constexpr std::uint16_t kSparkline = 1043;
inline constexpr std::uint16_t kBeginCondFmt14 = 1046;
inline constexpr std::uint16_t kEndCondFmt14 = 1047;
inline constexpr std::uint16_t kBeginCfRule14 = 1048;
inline constexpr std::uint16_t kCfvo14 = 1050;
inline constexpr std::uint16_t kDVal14 = 1053;
inline constexpr std::uint16_t kBeginDVals14 = 1054;
inline constexpr std::uint16_t kBeginSparklines = 1056;
inline constexpr std::uint16_t kEndSparklines = 1057;
inline constexpr std::uint16_t kBeginSparklineGroups = 1058;
inline constexpr std::uint16_t kEndSparklineGroups = 1059;
inline constexpr std::uint16_t kBeginCfs14 = 1135;
inline constexpr std::uint16_t kEndCfs14 = 1136;
inline constexpr std::uint16_t kEndDVals14 = 1154;
inline constexpr std::uint16_t kUid = 3072;

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_FRT_RECORDS_H_
