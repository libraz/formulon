//
// Scenarios as xlsb records: BrtBeginScenMan (u16 current, u16 show, both 0xFFFF when absent, then an Sqrfx whose
// count is 0xFFFFFFFF when the sqref is absent), per scenario BrtBeginSct (u16 input-cell count, locked at byte 2,
// hidden at byte 6, the other bytes of the 10-byte head zero, then name, nullable comment, user) with one BrtSlc per
// input cell (u32 row, u32 col, u32 deleted, u32 undone, u16 numFmtId, value text) and BrtEndSct, then
// BrtEndScenMan. The model keeps them as the retained `<scenarios>` worksheet child, so both formats read and write
// the one element.

#ifndef FORMULON_IO_XLSB_SCENARIO_RECORDS_H_
#define FORMULON_IO_XLSB_SCENARIO_RECORDS_H_

#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "io/xlsb/record.h"
#include "sheet_passthrough.h"

namespace formulon::io::xlsb {

/// The `<scenarios>` element a scenario-manager block describes, from its records from BrtBeginScenMan up to but not
/// including BrtEndScenMan; nullopt when any record falls outside the layout above.
std::optional<std::string> decode_scenarios(const std::vector<XlsbRecord>& records);

/// Replaces the retained `<scenarios>` child with `xml`.
void write_scenarios(WorksheetRawExtensions& children, std::string xml);

/// Inserts the scenario-manager block for the retained `<scenarios>` child at the head of the pre-merge tail, where
/// the grammar puts it after the protection records; nothing when the sheet has none, when the tail already holds a
/// block kept as read, or when any part of the element falls outside what the records carry.
void splice_scenarios(std::vector<std::uint8_t>& before_merges, const WorksheetRawExtensions& children);

}  // namespace formulon::io::xlsb

#endif  // FORMULON_IO_XLSB_SCENARIO_RECORDS_H_
