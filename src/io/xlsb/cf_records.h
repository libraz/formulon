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
#include <string>
#include <unordered_set>
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

/// Folds the x14 data-bar settings the legacy records cannot carry
/// (negative fill and border, axis position and colour, border, gradient,
/// bar lengths) onto the linked rules of `formats`, as the .xlsx reader
/// does with `<x14:dataBar>`. `records` is a run of framed tail records;
/// the x14 blocks stay retained verbatim. An x14 rule whose records hold
/// anything unmeasured leaves its legacy rule as decoded.
void apply_x14_data_bar_overlays(ByteSpan records, std::vector<cf::ConditionalFormat>& formats);

/// Brings the x14 data-bar rules in a sheet's retained tail `records` in
/// line with the model: a rule linked to a model data bar has its
/// model-owned settings rewritten (thresholds and direction are kept), a
/// data-bar rule whose model rule is gone is dropped, along with any block
/// or container that leaves empty. Every data-bar id still present is
/// added to `linked`. Records this does not recognise pass through.
void reconcile_x14_data_bars(std::vector<std::uint8_t>& records, const std::vector<cf::ConditionalFormat>& formats,
                             std::unordered_set<std::string>& linked);

/// Adds an x14 data-bar rule to `records` (the tail slot after the
/// hyperlinks) for every model data bar that needs one and is not in
/// `linked`, inside the retained x14 container when there is one, and adds
/// its id to `linked`. Fails for a formula threshold, whose x14 form is
/// unmeasured.
Expected<void, Error> add_x14_data_bars(std::vector<std::uint8_t>& records,
                                        const std::vector<cf::ConditionalFormat>& formats,
                                        std::unordered_set<std::string>& linked);

/// Emits `format` as one framed block. A data bar whose id is in `linked`
/// has an x14 counterpart carrying its lengths, and keeps the pre-2010
/// 10/90 in the legacy record as Excel does. Fails on a rule the records
/// cannot carry (a formula the Ptg codec refuses, an x14 link id that is
/// not a GUID, a malformed number threshold).
Expected<void, Error> emit_cf_block(std::vector<std::uint8_t>& dst, const cf::ConditionalFormat& format,
                                    const FeatureFormulaWriteContext& ctx,
                                    const std::unordered_set<std::string>& linked);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_CF_RECORDS_H_
