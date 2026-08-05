//
// "Show values as" transforms applied after the raw aggregation has
// finished. Each data field with `show_as != Normal` triggers a walk
// of `result.values` that replaces every cell with the derived value.
//
// Modes implemented:
//
//   * PercentOfRow / PercentOfCol: cell / row-or-col sum (Div0 when 0).
//   * PercentOfTotal: cell / table grand total.
//   * RunningTotalInRow / RunningTotalInCol: cumulative sum along the
//     axis (errors short-circuit later cells in that row / column).
//   * Index: (cell * total) / (row_sum * col_sum); Div0 if either
//     partial is 0.
//   * DifferenceFrom / PercentDifferenceFrom: per the base-field axis
//     selection.
//   * PercentOfParent / PercentOfParentRow / PercentOfParentCol: per
//     the parent subtotal at the selected hierarchy depth.
//
// Subtotal / grand-total propagation policy (partial):
//   * The three Percent* ratio modes (PercentOfRow / PercentOfCol /
//     PercentOfTotal) propagate the transform to `row_subtotals`,
//     `col_subtotals`, and `grand_totals` so the rendered subtotal /
//     grand-total cells display the same ratio Excel would emit at
//     those positions.
//   * RunningTotalInRow / RunningTotalInCol leave subtotals and grand
//     totals at their raw aggregate. A running total at a subtotal
//     break is semantically the cumulative position at that point, but
//     subtotal rows are aggregated independently from the leaf cells;
//     the running cumulative position is no longer recoverable
//     post-aggregation, so we surface the raw subtotal aggregate
//     instead of synthesising a misleading running value.
//   * Index uses partials (row_sum * col_sum / total) that have no
//     meaningful analogue at a subtotal break, so subtotals and grand
//     totals stay raw for the same reason.
//
// After mutation, the legacy mirror `result.subtotals[i]` is re-synced
// from `result.row_subtotals[i].values`, and `result.grand_total` is
// re-synced from `result.grand_totals[0]`.

#ifndef FORMULON_PIVOT_LAYOUT_GENERATOR_H_
#define FORMULON_PIVOT_LAYOUT_GENERATOR_H_

#include <cstddef>
#include <vector>

#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"

namespace formulon::pivot {

/// Applies the `show_as` transform of every data field in `table` to
/// `result`. `row_subtotal_leaf_sets` / `col_subtotal_leaf_sets` are
/// the per-subtotal leaf-set vectors the evaluator built when emitting
/// `result.row_subtotals` / `result.col_subtotals`; PercentOfParent*
/// uses them to locate the right subtotal for each leaf.
void apply_show_values_as_transforms(const PivotTable& table, PivotResult& result,
                                     const std::vector<std::vector<std::size_t>>& row_subtotal_leaf_sets,
                                     const std::vector<std::vector<std::size_t>>& col_subtotal_leaf_sets);

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_LAYOUT_GENERATOR_H_
