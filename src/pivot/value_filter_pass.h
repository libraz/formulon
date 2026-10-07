//
// Value-axis filter pass (Top-N, GreaterThan, Between) of the pivot
// evaluator, run after grand totals are emitted.
//

#ifndef FORMULON_PIVOT_VALUE_FILTER_PASS_H_
#define FORMULON_PIVOT_VALUE_FILTER_PASS_H_

#include <cstddef>
#include <vector>

#include "pivot/aggregator.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"

namespace formulon::pivot {

/// Applies every value-axis filter of `table` to `result`.
///
/// A filter on the field at axis depth d keeps or drops whole depth-d items,
/// scored by the item's own aggregate over the records still standing and
/// ranked within each depth-(d-1) parent independently, as Excel does. The
/// row/col tree is then collapsed by dropping unkept leaves and any interior
/// node whose subtree becomes empty; a subtotal whose group is filtered away
/// entirely is dropped along with the leaves it covered.
///
/// Every total -- grand totals, per-leaf totals, and surviving subtotal
/// values -- is then recomputed from `buckets` restricted to the surviving
/// leaves: Excel's Top-N grand total covers only the visible rows, not the
/// pre-filter set (verified against pivot_value_date_filters.xlsx /
/// pivot_recurring_period_filter.xlsx).
///
/// Both the active (slicer) filters and the authored `<filters>` entries feed
/// this pass; the latter are projected onto the same shape. The subtotal
/// leaf sets are recompacted in step with the pruned leaves.
void apply_value_filters(const PivotTable& table, const PivotCache& cache, const RecordBuckets& buckets,
                         PivotResult& result, std::vector<std::vector<std::size_t>>& row_subtotal_leaf_sets,
                         std::vector<std::vector<std::size_t>>& col_subtotal_leaf_sets);

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_VALUE_FILTER_PASS_H_
