//
// Value-axis filter pass (top / bottom N, percent and sum, value
// comparisons) of the pivot evaluator, run after grand totals are emitted.
//

#ifndef FORMULON_PIVOT_VALUE_FILTER_PASS_H_
#define FORMULON_PIVOT_VALUE_FILTER_PASS_H_

#include <cstddef>
#include <vector>

#include "pivot/aggregator.h"
#include "pivot/calculated.h"
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
/// The value types of `active_filters()` run in document order, each pruning
/// what the previous ones left. An item scores the value its row or column
/// shows: with calculated cells (`has_calculated_cells`) or on a calculated
/// field that is the fold of its cells, the field's formula applied last;
/// otherwise its aggregate over its records. An item with no records that
/// shows nothing scores 0. The subtotal leaf sets are recompacted in step
/// with the pruned leaves.
///
/// A filter that keeps only items without records leaves every total as it
/// was. Returns true when the totals were re-derived.
///
/// The leaf counts are passed explicitly because either may be 0 (an axis
/// field with no surviving records), leaving `buckets` empty.
bool apply_value_filters(const PivotTable& table, const PivotCache& cache, const RecordBuckets& buckets,
                         std::size_t row_leaf_count, std::size_t col_leaf_count, PivotResult& result,
                         std::vector<std::vector<std::size_t>>& row_subtotal_leaf_sets,
                         std::vector<std::vector<std::size_t>>& col_subtotal_leaf_sets, const CalculatedPlan& plan,
                         bool has_calculated_cells);

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_VALUE_FILTER_PASS_H_
