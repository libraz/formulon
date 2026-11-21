//
// Conditional-format rule merge on load. Excel coalesces the rules of a sheet
// when it opens a workbook; both loaders run this pass once every rule is in
// the model, so a loaded workbook holds the rules Excel holds after opening
// the same file. Rules added or edited through the API are never merged in the
// model; a save writes them as the merge leaves them (`saved_conditional_formats`).

#ifndef FORMULON_IO_CF_MERGE_H_
#define FORMULON_IO_CF_MERGE_H_

#include <vector>

#include "cf/cf_types.h"
#include "sheet.h"
#include "styles.h"
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
/// Only expression, cellIs, text, blank, error and timePeriod rules take part
/// in the merge, and only outside pivot-scoped blocks and raw extension
/// content. A sheet's rules are left unmerged when its retained extension
/// content holds a rule the model does not, since the intermediate-priority
/// test cannot see that rule. Must run after `ingest_feature_formulas`, when
/// formulas are canonical and the dxf table is complete.
///
/// Independently of the merge, the sqref pieces of every such block holding
/// only those kinds are joined where two tile one rectangle, then stably
/// ordered by top row. A block of one data bar or colour scale rule has a
/// single-cell first piece trade places with the second; other kinds, and a
/// block of several such rules, keep their pieces as stored.
void merge_cf_rules_on_load(Workbook& wb);

/// The conditional formats a save of `sheet` writes: a copy of the model's
/// blocks with the rule merge of `merge_cf_rules_on_load` applied (without the
/// sqref piece normalization, which is not idempotent for data bars and colour
/// scales). The model is untouched, so rules added through the API keep their
/// own blocks in memory. Rules the merge keeps out (extension content, x14
/// links, a sheet whose retained extension holds an unmodelled rule) are
/// written as held, and formats already merged come back unchanged. Every
/// part of a save that reads a sheet's rules (blocks, x14 payloads, dxf
/// references) must read this view. `dxfs` is the workbook's dxf table.
std::vector<cf::ConditionalFormat> saved_conditional_formats(const Sheet& sheet,
                                                             const std::vector<DifferentialFormat>& dxfs);

}  // namespace formulon::io

#endif  // FORMULON_IO_CF_MERGE_H_
