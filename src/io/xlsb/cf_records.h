//
// MS-XLSB conditional-formatting block codec: one
// BrtBeginConditionalFormatting .. BrtEndConditionalFormatting run to and
// from a `cf::ConditionalFormat`. Layouts were measured from Excel 365
// (macOS) re-saves, cross-checked against the .xlsx Excel wrote for the
// same workbook; `cf_records.cpp` carries the field tables.
//
// Decoding is all-or-nothing per block: a block holding any value outside
// the measured set (a theme colour, an unobserved template, a pivot-scoped
// block, a formula the Ptg codec refuses) decodes to nothing, and the
// caller keeps its bytes verbatim instead.

#ifndef FORMULON_IO_XLSB_CF_RECORDS_H_
#define FORMULON_IO_XLSB_CF_RECORDS_H_

#include <cstdint>
#include <optional>
#include <vector>

#include "cf/cf_types.h"
#include "io/xlsb/feature_formula.h"
#include "io/zip_reader.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

inline constexpr std::uint16_t kBrtBeginConditionalFormatting = 461;
inline constexpr std::uint16_t kBrtEndConditionalFormatting = 462;

/// Decodes one framed block, starting at its BrtBeginConditionalFormatting
/// and ending with its BrtEndConditionalFormatting.
std::optional<cf::ConditionalFormat> decode_cf_block(ByteSpan block, const FeatureFormulaReadContext& ctx);

/// Emits `format` as one framed block. Fails on a rule the records cannot
/// carry (a formula the Ptg codec refuses, an x14 link id that is not a
/// GUID, a malformed number threshold).
Expected<void, Error> emit_cf_block(std::vector<std::uint8_t>& dst, const cf::ConditionalFormat& format,
                                    const FeatureFormulaWriteContext& ctx);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_CF_RECORDS_H_
