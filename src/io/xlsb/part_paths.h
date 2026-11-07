//
// Fixed package paths of the MS-XLSB workbook-level parts.

#ifndef FORMULON_IO_XLSB_PART_PATHS_H_
#define FORMULON_IO_XLSB_PART_PATHS_H_

#include <string_view>

namespace formulon {
namespace io {
namespace xlsb {

inline constexpr std::string_view kWorkbookBinPath = "xl/workbook.bin";
inline constexpr std::string_view kWorkbookBinRelsPath = "xl/_rels/workbook.bin.rels";
inline constexpr std::string_view kStylesBinPath = "xl/styles.bin";
inline constexpr std::string_view kSharedStringsBinPath = "xl/sharedStrings.bin";
inline constexpr std::string_view kMetadataBinPath = "xl/metadata.bin";

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PART_PATHS_H_
