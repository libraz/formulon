//
// Sparkline groups as worksheet future records ([MS-XLSB] 2.1.7.62
// SPARKLINEGROUPS): BrtFRTBegin, BrtBeginSparklineGroups (1058), per group
// an optional revision-uid block (BrtACBegin, BrtUid, BrtACEnd),
// BrtBeginSparklineGroup (1041), BrtBeginSparklines (1056), one BrtSparkline
// (1043) per sparkline, BrtEndSparklines (1057), BrtEndSparklineGroup
// (1042), then BrtEndSparklineGroups (1059) and BrtFRTEnd.
//
// The model is the `x14:sparklineGroups` XML of `Sheet::ext_lst_xml()` that
// `io/sparkline_groups.h` reads and edits: the reader moves each group it can
// decode out of the retained tail into that XML, and the writer encodes the
// XML's groups back. A group holding anything outside the supported shape
// stays in the retained tail, verbatim.

#ifndef FORMULON_IO_XLSB_SPARKLINE_RECORDS_H_
#define FORMULON_IO_XLSB_SPARKLINE_RECORDS_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "io/xlsb/feature_formula.h"
#include "sheet_passthrough.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Moves every decodable sparkline group of `tail` into `ext_lst_xml`,
/// dropping its records and the container once it is empty.
void take_sparkline_groups(XlsbSheetTail& tail, const FeatureFormulaReadContext& ctx, std::string& ext_lst_xml);

/// True when `tail` still retains sparkline records.
bool has_retained_sparkline_groups(const XlsbSheetTail& tail);

/// Writes the sparkline groups of `ext_lst_xml` into the tail slot `records`
/// (the slot after the hyperlinks): inside a retained container, or as a new
/// container after the x14 conditional formats and data validations. A
/// sparkline formula without a sheet is read on `sheet_name`. A group with a
/// formula that is not one reference or name, or a location that is not one
/// cell, is left out and added to `*skipped`.
Expected<void, Error> add_sparkline_groups(std::vector<std::uint8_t>& records, std::string_view ext_lst_xml,
                                           const FeatureFormulaWriteContext& ctx, std::string_view sheet_name,
                                           std::uint32_t* skipped = nullptr);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_SPARKLINE_RECORDS_H_
