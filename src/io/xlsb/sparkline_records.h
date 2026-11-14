//
// Sparkline groups as worksheet future records ([MS-XLSB] 2.1.7.62
// SPARKLINEGROUPS): BrtFRTBegin, BrtBeginSparklineGroups (1058), per group
// an optional revision-uid block (BrtACBegin, BrtUid, BrtACEnd),
// BrtBeginSparklineGroup (1041), BrtBeginSparklines (1056), one BrtSparkline
// (1043) per sparkline, BrtEndSparklines (1057), BrtEndSparklineGroup
// (1042), then BrtEndSparklineGroups (1059) and BrtFRTEnd.
//
// The model is the `x14:sparklineGroups` XML of `Sheet::ext_lst_xml()` that
// `io/sparkline_groups.h` reads and edits: the reader moves each group out of
// the tail into that XML, and the writer encodes the XML's groups back. A
// deleted source is a BrtSparkline without a formula, and an absent colour is
// an automatic BrtColor with index 0.

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

/// Moves every sparkline group of `tail` into `ext_lst_xml` and drops the
/// containers. A group that does not decode, or that `add_sparkline_group`
/// refuses, is not loaded: it is added to `*skipped` (when non-null) and
/// logged as `xlsb.reader.skipped`.
void take_sparkline_groups(XlsbSheetTail& tail, const FeatureFormulaReadContext& ctx, std::string& ext_lst_xml,
                           std::uint32_t* skipped = nullptr);

/// Writes the sparkline groups of `ext_lst_xml` into the tail slot `records`
/// (the slot after the hyperlinks): inside a retained container, or as a new
/// container after the x14 conditional formats and data validations. A
/// sparkline formula without a sheet is read on `sheet_name`, and a location
/// rectangle is written as its top-left cell. A formula or date-axis
/// reference that is not one reference or name, or a location that is not one
/// rectangle, fails with `kIoFeatureNotWritable` (context
/// `kind=sparklines sheet_index=<sheet_index>`).
Expected<void, Error> add_sparkline_groups(std::vector<std::uint8_t>& records, std::string_view ext_lst_xml,
                                           const FeatureFormulaWriteContext& ctx, std::string_view sheet_name,
                                           std::size_t sheet_index);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_SPARKLINE_RECORDS_H_
