//
// Pivot filter engine.
//
// Two filter passes live here:
//
//   * Pre-aggregation, per-record: `PreparedRecordFilter` runs the manual
//     `items[]` visibility check on every field that hides something, then
//     applies axis-level label / date filters from `active_filters`.
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
#include <string_view>
#include <unordered_set>
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

/// The pre-aggregation filter pass with everything that depends only on the
/// table hoisted out of the per-record decision.
///
/// Which items a field hides, which field a named axis filter designates, and
/// which date window a relative period spans are all properties of the table,
/// not of a record. Deriving them inside the record loop costs
/// O(records x fields x items), and Excel authors an `items[]` list for every
/// field it places on an axis, so that product is paid whether or not anything
/// is actually hidden. Worse, matching an item by label re-renders the record's
/// value once per item, so a field with a few thousand distinct values turns
/// the pass into tens of millions of number formatting calls.
///
/// Built once, a record costs one hash lookup per field that hides something,
/// and renders its label at most once per such field.
///
/// Borrows `table` and `cache`: neither may be destroyed, and `table` may not
/// be mutated, while the filter is alive — the hidden-label set holds views
/// into `PivotItem::name`.
class PreparedRecordFilter {
 public:
  PreparedRecordFilter(const PivotTable& table, const PivotCache& cache, const PivotFilterEnv& env);

  /// True iff `record` survives the manual `items[]` filter on every field
  /// that hides something, AND the axis-level label / date filters in
  /// `active_filters`, AND the authored `<filters>` entries decided per
  /// record. An `items[]` list with nothing hidden matches every value (Excel
  /// authors the list for any field it places on an axis, hidden items or
  /// not).
  ///
  /// A hidden item normally matches records by its label. The blank item is
  /// the exception: it has no label of its own, so it is matched by the cache
  /// value it binds to (`shared_items[cache_index]`, the same binding
  /// `resolve_pivot_names` reads) being blank. Identifying it by the
  /// placeholder the grid draws instead would also hide any genuine text value
  /// spelled the same way. An item that is unlabelled *and* binds to nothing
  /// resolvable is malformed and filters nothing.
  ///
  /// An axis filter names its field under the shared resolution rule
  /// (`resolve_field_by_any_name`); a name that resolves to nothing filters
  /// nothing, because the public mutators already reject one on entry.
  bool passes(const PivotCacheRecord& record) const;

 private:
  /// The labels one field hides, keyed for O(1) membership. `hides_blank` is
  /// the unlabelled item's rule, resolved through its cache binding at build
  /// time; an unlabelled item that binds to nothing resolvable is malformed
  /// and contributes neither.
  struct HiddenItems {
    std::size_t field_index = 0;
    std::unordered_set<std::string_view> labels;
    bool hides_blank = false;
  };

  /// An `active_filters` entry whose field name already resolved. Entries
  /// naming nothing are dropped at build time rather than re-resolved — and
  /// re-failing — per record.
  struct ResolvedFilter {
    std::size_t field_index = 0;
    const PivotFilter* filter = nullptr;
  };

  /// A relative-period entry with its window already resolved against the
  /// clock reading the pass was built with.
  struct ResolvedPeriod {
    std::size_t field_index = 0;
    DateWindow window;
  };

  /// A recurring-period entry with its field index validated. Unlike
  /// `ResolvedPeriod` there is nothing to resolve against a clock: the
  /// criterion is which calendar month a record's date falls in, so the
  /// bounds are carried through from the file as-is.
  struct ResolvedRecurring {
    std::size_t field_index = 0;
    unsigned month_low = 1;
    unsigned month_high = 1;
  };

  const PivotCache* cache_ = nullptr;
  const PivotTable* table_ = nullptr;
  /// The workbook epoch, kept because a recurring selector reads a
  /// record's month back out of its serial at match time rather than
  /// resolving to serials up front the way a window does.
  bool date1904_ = false;
  std::vector<HiddenItems> hidden_items_;
  std::vector<ResolvedFilter> label_filters_;
  std::vector<ResolvedPeriod> period_windows_;
  std::vector<ResolvedRecurring> recurring_months_;
};

/// Projects an authored `<filters>` value entry onto the `PivotFilter`
/// shape the post-aggregation pass already consumes.
///
/// A file names only the field, so the axis is recovered here from that
/// field's membership in `row_field_order()` / `col_field_order()`.
/// Returns `nullopt` when the entry is not a value family (a date entry
/// is applied pre-aggregation by `PreparedRecordFilter`) or when
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
