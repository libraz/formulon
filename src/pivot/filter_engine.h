//
// Pivot filter engine.
//
// Two filter passes live here:
//
//   * Pre-aggregation, per-record: `PreparedRecordFilter` runs the manual
//     `items[]` visibility check on every field that hides something, then
//     applies the caption and date filters from `active_filters`.
//
//   * Post-aggregation, per-item: `build_grouped_value_filter_keep` ranks
//     the filtered field's items within each parent group and produces a
//     boolean keep mask aligned to `result.values`'s DFS pre-order leaf
//     enumeration; `compact_leaf_axis` then re-expresses the whole
//     result in the surviving-leaf index space.

#ifndef FORMULON_PIVOT_FILTER_ENGINE_H_
#define FORMULON_PIVOT_FILTER_ENGINE_H_

#include <cstddef>
#include <cstdint>
#include <deque>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_set>
#include <variant>
#include <vector>

#include "eval/collated_wildcard.h"
#include "pivot/hierarchy_builder.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "utils/date_time.h"
#include "value.h"

namespace formulon {
struct StylesTable;
}  // namespace formulon

namespace formulon::pivot {
class PivotFormulaHost;

/// True for the types ranked on an aggregate after aggregation (the top /
/// bottom families and the value comparisons), as opposed to the caption and
/// date types applied per record.
inline bool is_value_filter_type(FilterType type) {
  switch (type) {
    case FilterType::ValueTop10:
    case FilterType::ValueGreaterThan:
    case FilterType::ValueBetween:
    case FilterType::ValueEqual:
    case FilterType::ValueNotEqual:
    case FilterType::ValueGreaterThanOrEqual:
    case FilterType::ValueLessThan:
    case FilterType::ValueLessThanOrEqual:
    case FilterType::ValueNotBetween:
    case FilterType::ValueBottomCount:
    case FilterType::ValueTopPercent:
    case FilterType::ValueBottomPercent:
    case FilterType::ValueTopSum:
    case FilterType::ValueBottomSum:
      return true;
    default:
      return false;
  }
}

/// True for the caption types, matched against an item's label.
inline bool is_caption_filter_type(FilterType type) {
  return type == FilterType::LabelContains || type == FilterType::LabelBeginsWith ||
         (type >= FilterType::LabelEqual && type <= FilterType::LabelNotBetween);
}

/// True for the relative-period and calendar-position types, which carry no
/// operand.
inline bool is_operandless_filter_type(FilterType type) {
  return type >= FilterType::DateTomorrow;
}

/// True for the types that read both `value` and `value_high`.
inline bool is_range_filter_type(FilterType type) {
  switch (type) {
    case FilterType::ValueBetween:
    case FilterType::ValueNotBetween:
    case FilterType::LabelDate:
    case FilterType::LabelBetween:
    case FilterType::LabelNotBetween:
    case FilterType::DateNotBetween:
      return true;
    default:
      return false;
  }
}

/// Evaluation-time inputs a filter can need beyond the pivot model itself.
///
/// The filter's relative-period family reads the clock and epoch; the
/// evaluator also reads the epoch and styles to label axis items. The
/// default value reproduces the pre-seam behaviour exactly: follow the host
/// clock, 1900 epoch.
struct PivotFilterEnv {
  /// Pinned wall-clock reading (`Workbook::pinned_now()`), or empty to read
  /// the host clock. Pinning is what makes a pivot carrying a
  /// relative-period filter reproducible.
  std::optional<date_time::CivilTime> pinned_now;
  /// The workbook date epoch, so a window resolved from the clock lands on
  /// the same serial scale as the dates stored in the cache.
  bool date1904 = false;
  /// The workbook styles, resolving a custom `numFmtId` an axis item is
  /// labelled with. Null renders custom ids as General.
  const StylesTable* styles = nullptr;
  /// Evaluates calculated-field and calculated-item formulas. Null leaves
  /// calculated cells empty.
  PivotFormulaHost* formulas = nullptr;
};

/// An inclusive date window, in Excel serials under the workbook epoch.
struct DateWindow {
  double low = 0.0;
  double high = 0.0;
};

/// Resolves a relative-period type (`DateTomorrow` .. `DateYearToDate`)
/// against a wall-clock reading.
///
/// Exposed rather than kept internal because the calendar boundaries are
/// the whole substance of the family — month lengths, quarter starts, the
/// year-to-date half-open end — and they are worth asserting directly
/// instead of only through a filtered pivot.
DateWindow resolve_relative_period(FilterType type, const date_time::CivilTime& now, bool date1904);

/// The wildcard pattern a begins / ends / contains caption type (or its
/// negation) matches the whole label against: `v*`, `*v` or `*v*`. Any
/// `*` / `?` / `~` in `value` stays live. Empty for the other types, which
/// do not match by an affixed pattern.
std::optional<std::string> caption_wildcard_pattern(FilterType type, std::string_view value);

/// The numeric payload of a filter operand slot; empty when unset, text, or
/// not finite, which no file can spell.
std::optional<double> payload_number(const std::variant<std::monostate, int, double, std::string>& slot);

/// The index into `table.fields()` a filter applies to: the field its name
/// resolves to, or, for a decoded entry whose name is still empty, its
/// source `fld`. Empty when neither designates a field.
std::optional<std::size_t> filter_field_index(const PivotTable& table, const PivotFilter& f);

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
/// Borrows `cache`, which may not be destroyed while the filter is alive.
/// The hidden-label set views labels the filter owns, so it is neither
/// copyable nor movable.
class PreparedRecordFilter {
 public:
  PreparedRecordFilter(const PivotTable& table, const PivotCache& cache, const PivotFilterEnv& env);
  PreparedRecordFilter(const PreparedRecordFilter&) = delete;
  PreparedRecordFilter& operator=(const PreparedRecordFilter&) = delete;
  PreparedRecordFilter(PreparedRecordFilter&&) = delete;
  PreparedRecordFilter& operator=(PreparedRecordFilter&&) = delete;

  /// True iff `record` survives the manual `items[]` filter on every field
  /// that hides something, AND every caption and date filter in
  /// `active_filters`. An `items[]` list with nothing hidden matches every
  /// value (Excel authors the list for any field it places on an axis,
  /// hidden items or not).
  ///
  /// A hidden item normally matches records by its label (`pivot_item_label`,
  /// so an item built by cache index matches its bound value's rendering).
  /// A group field's item is matched by its group-item number. The blank
  /// item is the exception: it has no label of its own, so it is
  /// matched by the cache value it binds to (`pivot_item_cache_value`) being
  /// blank. Identifying it by the placeholder the grid draws instead would
  /// also hide any genuine text value spelled the same way. An item that is
  /// unlabelled *and* binds to nothing resolvable is malformed and filters
  /// nothing.
  ///
  /// A filter designates its field through `filter_field_index`; one that
  /// designates nothing filters nothing, because the public mutators already
  /// reject one on entry.
  bool passes(const PivotCacheRecord& record) const;

 private:
  /// The labels one field hides, keyed for O(1) membership. `hides_blank` is
  /// the unlabelled item's rule, resolved through its cache binding at build
  /// time; an unlabelled item that binds to nothing resolvable is malformed
  /// and contributes neither. A group field hides by group-item number
  /// instead, which is what its records carry.
  struct HiddenItems {
    std::size_t field_index = 0;
    std::unordered_set<std::string_view> labels;
    bool hides_blank = false;
    std::unordered_set<std::uint32_t> group_items;
  };

  /// A caption or date entry of `active_filters` with its field resolved and
  /// its payload decoded. Entries that designate nothing or carry an unusable
  /// payload are dropped at build time rather than re-failing per record.
  struct ResolvedFilter {
    std::size_t field_index = 0;
    FilterType type = FilterType::LabelEqual;
    /// Caption operands as matched: the affixed pattern for the begins / ends
    /// / contains family, otherwise the operand as entered.
    std::string low;
    std::string high;
    /// The caption operands read as numbers, compared with a numeric label.
    std::optional<double> low_number;
    std::optional<double> high_number;
    /// Whether a caption operand carries `*`, `?` or `~` and so matches as a
    /// pattern.
    bool low_pattern = false;
    bool high_pattern = false;
    /// Date bounds: the operands of a comparison, the inclusive window of a
    /// range or relative period, or the inclusive month range of a calendar
    /// position.
    double date_low = 0.0;
    double date_high = 0.0;
    /// Caption matchers, built on the first label; `passes` is const, so they
    /// mutate under it and the filter is single-threaded.
    mutable std::optional<eval::CollatedWildcard> matcher;
    mutable std::optional<eval::CollatedWildcard> matcher_high;
  };

  bool caption_passes(const ResolvedFilter& f, const Value& v) const;
  bool date_passes(const ResolvedFilter& f, const Value& v) const;

  const PivotCache* cache_ = nullptr;
  /// The workbook epoch, kept because a calendar-position selector reads a
  /// record's month back out of its serial at match time.
  bool date1904_ = false;
  /// Owns every hidden label `HiddenItems::labels` views. A deque, so a
  /// later label never relocates an earlier one.
  std::deque<std::string> hidden_labels_;
  std::vector<HiddenItems> hidden_items_;
  std::vector<ResolvedFilter> filters_;
};

/// Per-item scoring intermediate used by `build_value_filter_keep`.
/// `scores[i]` is the numeric aggregate of item `i`; `all_blank[i]` is
/// true when item `i` had no numeric content (such items can be excluded
/// even by `ValueGreaterThan`-style filters).
struct AxisScores {
  std::vector<double> scores;
  std::vector<bool> all_blank;
};

/// Builds a per-leaf keep mask for the value filter `f`. Returns `nullopt`
/// for filter shapes that should degrade to a no-op (e.g. an unbounded
/// `ValueBetween`).
///
/// The count types keep the N highest (lowest) scoring leaves. The percent
/// and sum types walk the leaves in ranked order and keep each one while the
/// running total is still short of the target, so the last kept leaf is the
/// one that reaches it; percent reads the target as a share of the total of
/// the leaves scored.
std::optional<std::vector<bool>> build_value_filter_keep(const PivotFilter& f, const AxisScores& axis);

/// Leaf keep mask for a value filter on the field whose axis nodes are
/// `groups` (see `axis_leaf_groups_at_depth`), with `group_scores[g]`
/// scoring group `g`.
///
/// Each run of groups sharing a parent is ranked or thresholded on its own,
/// so a Top-N on an inner field keeps N items under every outer item; a
/// kept group keeps all `leaf_count` of its leaves. Returns `nullopt` when
/// the filter shape is a no-op.
std::optional<std::vector<bool>> build_grouped_value_filter_keep(const PivotFilter& f,
                                                                 const std::vector<AxisLeafGroup>& groups,
                                                                 const AxisScores& group_scores,
                                                                 std::size_t leaf_count);

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

/// Numeric payload of a filter's `value` (`int` or `double`); 0 for any other
/// variant, which callers treat as a no-op.
double filter_number_value(const PivotFilter& f);

/// Numeric payload of a range filter's `value_high`; `nullopt` when unset, so
/// the range is unbounded above and the filter no-ops.
std::optional<double> filter_number_value_high(const PivotFilter& f);

/// Drops the entries of `items` whose index is marked false in `keep`,
/// preserving the relative order of the survivors. An index past the end of
/// `keep` is kept: the mask only describes the leaves the filter scored.
template <typename T>
void keep_marked(std::vector<T>& items, const std::vector<bool>& keep) {
  std::vector<T> kept;
  kept.reserve(items.size());
  for (std::size_t i = 0; i < items.size(); ++i) {
    if (i < keep.size() && !keep[i]) {
      continue;
    }
    kept.push_back(std::move(items[i]));
  }
  items = std::move(kept);
}

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_FILTER_ENGINE_H_
