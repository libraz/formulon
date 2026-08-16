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

#include "eval/date_time.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "value.h"

namespace formulon::pivot {

/// Evaluation-time inputs a filter can need beyond the pivot model itself.
///
/// Only the relative-period family reads any of this, but it travels with
/// every filter pass so the two entry points keep one signature. The
/// default value reproduces the pre-seam behaviour exactly: follow the host
/// clock, 1900 epoch.
struct PivotFilterEnv {
  /// Pinned wall-clock reading (`Workbook::pinned_now()`), or empty to read
  /// the host clock. Pinning is what makes a pivot carrying a
  /// relative-period filter reproducible.
  std::optional<eval::date_time::CivilTime> pinned_now;
  /// The workbook date epoch, so a window resolved from the clock lands on
  /// the same serial scale as the dates stored in the cache.
  bool date1904 = false;
};

/// An inclusive date window, in Excel serials under the workbook epoch.
struct DateWindow {
  double low = 0.0;
  double high = 0.0;
};

/// Resolves a relative period against a wall-clock reading.
///
/// Exposed rather than kept internal because the calendar boundaries are
/// the whole substance of the family — month lengths, quarter starts, the
/// year-to-date half-open end — and they are worth asserting directly
/// instead of only through a filtered pivot.
DateWindow resolve_relative_period(RelativePeriod period, const eval::date_time::CivilTime& now, bool date1904);

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
bool record_passes_manual_filter(const PivotTable& table, const PivotCache& cache, const PivotCacheRecord& record,
                                 const PivotFilterEnv& env = PivotFilterEnv{});

/// Projects an authored `<filters>` value entry onto the `PivotFilter`
/// shape the post-aggregation pass already consumes.
///
/// A file names only the field, so the axis is recovered here from that
/// field's membership in `row_field_order()` / `col_field_order()`.
/// Returns `nullopt` when the entry is not a value family (a date entry
/// is applied pre-aggregation by `record_passes_manual_filter`) or when
/// the field sits on neither axis, leaving nothing to prune.
std::optional<PivotFilter> authored_value_filter_as_pivot_filter(const PivotTable& table,
                                                                 const AuthoredValueFilter& authored);

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

/// Keep mask for the two running-total flavours of the top-N dialog.
///
/// Both walk the scoring leaves in descending order and stop once the
/// running total first reaches the target, so the last kept leaf is the
/// one that crosses it. `Percent` reads `target` as a share of the axis
/// total; `Sum` reads it as an absolute amount. Returns `nullopt` for
/// `Items`, which counts leaves rather than accumulating them and is
/// served by `build_value_filter_keep`.
std::optional<std::vector<bool>> build_running_total_keep(TopNBasis basis, double target, const AxisScores& axis);

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
