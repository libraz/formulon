//
// "Show values as" transforms applied after the raw aggregation has
// finished. Every cell of the report for a data field with
// `show_as != Normal` -- leaves, subtotals, margin totals and the grand
// total -- is recomputed from a snapshot of the untransformed grid, so a
// transformed cell never feeds another cell's reference.
//
// Cells are addressed by item path on each axis. Modes along a base field
// (difference, percent, running total, rank, percent of parent) act on the
// siblings of the cell at the base field's depth; the cells at or above
// that depth (the totals along the base field) render blank.
//
// After mutation, the legacy mirror `result.subtotals[i]` is re-synced
// from `result.row_subtotals[i].values`, and `result.grand_total` is
// re-synced from `result.grand_totals[0]`.

#ifndef FORMULON_PIVOT_SHOW_VALUES_AS_H_
#define FORMULON_PIVOT_SHOW_VALUES_AS_H_

#include <cstddef>
#include <vector>

#include "pivot/pivot_cache.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"

namespace formulon::pivot {

/// Applies the `show_as` transform of every data field in `table` to
/// `result`. `cache` is the table's bound cache, which labels a specific
/// `baseItem`. The subtotal leaf-set vectors are not consulted.
void apply_show_values_as_transforms(const PivotTable& table, const PivotCache& cache, PivotResult& result,
                                     const std::vector<std::vector<std::size_t>>& row_subtotal_leaf_sets,
                                     const std::vector<std::vector<std::size_t>>& col_subtotal_leaf_sets);

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_SHOW_VALUES_AS_H_
