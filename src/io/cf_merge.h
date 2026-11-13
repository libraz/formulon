//
// Conditional-format rule merge on load. Excel coalesces the rules of a sheet
// when it opens a workbook; both loaders run this pass once every rule is in
// the model, so a loaded workbook holds the rules Excel holds after opening
// the same file. Rules added or edited through the API are never merged.

#ifndef FORMULON_IO_CF_MERGE_H_
#define FORMULON_IO_CF_MERGE_H_

#include "workbook.h"

namespace formulon::io {

/// Merges the conditional-format rules of every sheet of `wb` as Excel does at
/// load. Two rules merge when they have the same definition (type, operator,
/// stopIfTrue, type fields, formulas compared relative to each block's anchor,
/// dxf content), their cell sets union to a rectangle or are equal, and no
/// other rule with a priority between theirs intersects the range of the one
/// with the higher priority number. The survivor is the one with the lower
/// priority number (the earlier in model order on a tie), takes the bounding
/// box (or keeps its sqref when the cell sets are equal), and moves to its own
/// block when its block holds other rules. Repeats until nothing merges.
///
/// Only expression, cellIs, text, blank, error and timePeriod rules take part,
/// and only outside pivot-scoped blocks and raw extension content. A sheet is
/// left untouched when its retained extension content holds a rule the model
/// does not, since the intermediate-priority test cannot see that rule. Must
/// run after `ingest_feature_formulas`, when formulas are canonical and the
/// dxf table is complete.
void merge_cf_rules_on_load(Workbook& wb);

}  // namespace formulon::io

#endif  // FORMULON_IO_CF_MERGE_H_
