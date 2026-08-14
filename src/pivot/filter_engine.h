//
// Pivot filter engine.
//
// Two filter passes live here:
//
//   * Pre-aggregation, per-record: `record_passes_manual_filter` runs
//     the manual `items[]` visibility check on every field that declares
//     one, then applies axis-level label / date filters from
//     `active_filters`.
//
//   * Post-aggregation, per-leaf: `score_*_axis` + `build_value_filter_keep`
//     score each axis leaf for the active value filter and produce a
//     boolean keep mask aligned to `result.values`'s DFS pre-order leaf
//     enumeration; `compact_leaf_axis` then re-expresses the whole
//     result in the surviving-leaf index space.

#ifndef FORMULON_PIVOT_FILTER_ENGINE_H_
#define FORMULON_PIVOT_FILTER_ENGINE_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <vector>

#include "pivot/pivot_cache.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "value.h"

namespace formulon::pivot {

/// True iff `record` survives the manual `items[]` filter on every
/// field that declares one, AND the axis-level label filters in
/// `active_filters`. Empty `items` lists match all values (Excel default
/// — items[] is only authored when the user has hidden at least one
/// value).
///
/// A hidden item normally matches records by its label. The blank item is the
/// exception: it has no label of its own, so it is matched by the cache value
/// it binds to (`shared_items[cache_index]`, the same binding
/// `resolve_pivot_names` reads) being blank. Identifying it by the placeholder
/// the grid draws instead would also hide any genuine text value spelled the
/// same way. An item that is unlabelled *and* binds to nothing resolvable is
/// malformed and filters nothing.
///
/// An axis filter names its field under the shared resolution rule
/// (`resolve_field_by_any_name`); a name that resolves to nothing is skipped
/// here because the public mutators already reject one on entry.
bool record_passes_manual_filter(const PivotTable& table, const PivotCache& cache, const PivotCacheRecord& record);

/// Per-axis scoring intermediate used by `build_value_filter_keep`.
/// `scores[i]` is the numeric aggregate sum at leaf `i`; `all_blank[i]`
/// is true when leaf `i` had no numeric content (such leaves can be
/// excluded even by `ValueGreaterThan`-style filters).
struct AxisScores {
  std::vector<double> scores;
  std::vector<bool> all_blank;
};

/// Sums leaf scores along the row axis (one entry per row leaf,
/// reducing across the column axis). Used to back row-axis value
/// filters such as `ValueTop10` and `ValueGreaterThan`.
AxisScores score_row_axis(const PivotResult& result, std::size_t row_count, std::size_t col_count,
                          std::size_t data_field_index = 0);

/// Mirror of `score_row_axis` for the column axis.
AxisScores score_col_axis(const PivotResult& result, std::size_t col_count, std::size_t row_count,
                          std::size_t data_field_index = 0);

/// Builds a per-leaf keep mask for `f`. Returns `nullopt` for filter
/// shapes that should degrade to a no-op (e.g. unbounded `ValueBetween`).
std::optional<std::vector<bool>> build_value_filter_keep(const PivotFilter& f, const AxisScores& axis);

/// Which leaf axis a value filter pruned.
enum class LeafAxis : std::uint8_t {
  Row = 0,
  Col = 1,
};

/// Re-expresses every leaf-indexed structure of `result` in the surviving-leaf
/// index space of `axis` after a value filter produced `keep` (one entry per
/// pre-filter leaf, in the DFS pre-order `finalize_hierarchy` assigned).
///
/// Compacting only part of that state would leave the remainder indexed by the
/// pre-filter leaf space, so a subtotal would silently read a filter-excluded
/// leaf. Rewritten here for a row-axis prune: `values` rows, `row_leaf_totals`,
/// every `ColSubtotal::values` row slot, and the row-subtotal leaf sets; for a
/// column-axis prune: every `values` row's column slice, `col_leaf_totals`,
/// every `RowSubtotal::col_values` slot, and the column-subtotal leaf sets.
///
/// A subtotal whose leaf set becomes empty describes a label path that no
/// longer exists in the pruned hierarchy and is dropped, together with its
/// leaf set, its `PivotResult::subtotals` compatibility entry (row axis), and
/// its slot in every `RowSubtotal::col_subtotal_values` (column axis).
///
/// The caller keeps ownership of the leaf-set vectors because they outlive the
/// result only as evaluator locals handed to the show-values-as pass.
void compact_leaf_axis(PivotResult& result, const std::vector<bool>& keep, LeafAxis axis,
                       std::vector<std::vector<std::size_t>>& row_subtotal_leaf_sets,
                       std::vector<std::vector<std::size_t>>& col_subtotal_leaf_sets);

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_FILTER_ENGINE_H_
