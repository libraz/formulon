
#include "pivot/filter_engine.h"

#include <algorithm>
#include <cstddef>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <variant>
#include <vector>

#include "pivot/aggregator.h"
#include "pivot/field_lookup.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/record_access.h"
#include "pivot/value_order.h"
#include "utils/checked_index.h"
#include "value.h"

namespace formulon::pivot {
namespace {

// Pulls the string payload of a filter when one is expected. Returns
// an empty `string_view` when the variant carries a non-string value,
// which effectively makes the filter a no-op rather than a runtime
// error.
std::string_view filter_string_value(const PivotFilter& f) {
  if (const auto* s = std::get_if<std::string>(&f.value)) {
    return *s;
  }
  return {};
}

// Pulls the numeric payload of a filter when one is expected. The
// variant carries either `int` or `double`; both coerce to double here.
// Returns 0 for non-numeric variants (caller treats as no-op).
double filter_number_value(const PivotFilter& f) {
  if (const auto* d = std::get_if<double>(&f.value)) {
    return *d;
  }
  if (const auto* i = std::get_if<int>(&f.value)) {
    return static_cast<double>(*i);
  }
  return 0.0;
}

// Invokes `consume` with the display label for `v`.
//
// Rendering goes through the shared `display_string`, which is the single
// place a cache `Value` is turned into a pivot label. A private renderer
// here would let a record's label drift from the `PivotItem::name` /
// hierarchy label produced elsewhere, and a filter can only ever match a
// label that is spelled the same way on both sides.
//
// Text is the one kind whose label is the payload itself, so it is handed
// to `consume` directly instead of through a copy.
template <typename Consume>
bool with_display_string(const Value& v, Consume&& consume) {
  if (v.is_text()) {
    return consume(v.as_text());
  }
  const std::string rendered = display_string(v);
  return consume(std::string_view{rendered});
}

// Pulls the numeric upper-bound payload of a range filter. Returns
// `nullopt` when `value_high` is `monostate` (i.e. caller didn't set
// it), signalling the range is unbounded above; the calling filter
// then no-ops.
std::optional<double> filter_number_value_high(const PivotFilter& f) {
  if (const auto* d = std::get_if<double>(&f.value_high)) {
    return *d;
  }
  if (const auto* i = std::get_if<int>(&f.value_high)) {
    return static_cast<double>(*i);
  }
  return std::nullopt;
}

// Evaluates a `LabelDate` filter against the underlying numeric value
// of a record's field. The `field_name` ostensibly identifies a date
// column, so the cache value should be a date serial; non-numeric cells
// (text labels, blanks, errors) are not in the date domain and skip the
// filter rather than being dropped.
bool label_date_filter_passes(const PivotFilter& f, const Value& v) {
  if (!v.is_number()) {
    return true;  // Non-numeric cells skip the date filter.
  }
  const double serial = v.as_number();
  const double lo = filter_number_value(f);
  const auto hi_or = filter_number_value_high(f);
  if (!hi_or) {
    return true;  // Half-open / no upper bound: treat as no-op.
  }
  return serial >= lo && serial <= *hi_or;
}

// Evaluates a single label-flavoured filter against `label`.
// Value-filter types short-circuit to true: those are applied
// post-aggregation.
bool label_filter_passes(const PivotFilter& f, std::string_view label) {
  switch (f.type) {
    case FilterType::LabelContains: {
      const std::string_view needle = filter_string_value(f);
      if (needle.empty()) {
        return true;
      }
      return label.find(needle) != std::string::npos;
    }
    case FilterType::LabelBeginsWith: {
      const std::string_view needle = filter_string_value(f);
      if (needle.empty()) {
        return true;
      }
      return label.size() >= needle.size() && label.compare(0, needle.size(), needle) == 0;
    }
    case FilterType::ValueTop10:
    case FilterType::ValueGreaterThan:
    case FilterType::ValueBetween:
    case FilterType::LabelDate:
      // Value-axis filters: handled later by the post-aggregation pass.
      // `LabelDate` is dispatched through `label_date_filter_passes` above
      // by `record_passes_manual_filter`; this branch is unreachable for
      // it but kept exhaustive for the switch.
      return true;
  }
  return true;
}

// ASCII-only case folding for authored caption comparisons.
//
// Excel matches a pivot caption filter the way it matches an AutoFilter
// criterion: without regard to case. Only the ASCII range is folded here
// — a full Unicode case mapping is a table this build does not carry, and
// the pivot corpus that reaches these filters is field labels rather than
// arbitrary text.
constexpr char AsciiLower(char c) {
  return (c >= 'A' && c <= 'Z') ? static_cast<char>(c - 'A' + 'a') : c;
}

// Three-way ASCII-case-insensitive comparison, ordering by folded bytes
// and breaking ties on length.
int CaseInsensitiveCompare(std::string_view lhs, std::string_view rhs) {
  const std::size_t common = lhs.size() < rhs.size() ? lhs.size() : rhs.size();
  for (std::size_t i = 0; i < common; ++i) {
    const char a = AsciiLower(lhs[i]);
    const char b = AsciiLower(rhs[i]);
    if (a != b) {
      return a < b ? -1 : 1;
    }
  }
  if (lhs.size() == rhs.size()) {
    return 0;
  }
  return lhs.size() < rhs.size() ? -1 : 1;
}

bool CaseInsensitiveEquals(std::string_view lhs, std::string_view rhs) {
  return lhs.size() == rhs.size() && CaseInsensitiveCompare(lhs, rhs) == 0;
}

bool CaseInsensitiveContains(std::string_view haystack, std::string_view needle) {
  if (needle.empty()) {
    return true;
  }
  if (needle.size() > haystack.size()) {
    return false;
  }
  const std::size_t last = haystack.size() - needle.size();
  for (std::size_t i = 0; i <= last; ++i) {
    if (CaseInsensitiveEquals(haystack.substr(i, needle.size()), needle)) {
      return true;
    }
  }
  return false;
}

bool CaseInsensitiveBeginsWith(std::string_view label, std::string_view value) {
  return label.size() >= value.size() && CaseInsensitiveEquals(label.substr(0, value.size()), value);
}

bool CaseInsensitiveEndsWith(std::string_view label, std::string_view value) {
  return label.size() >= value.size() && CaseInsensitiveEquals(label.substr(label.size() - value.size()), value);
}

// Evaluates one decoded `<filters>` caption entry against `label`.
//
// The ordering comparisons treat the label as text even when it renders
// as digits, which is what a caption filter means: it filters what the
// grid draws, not the underlying cache value.
bool caption_filter_passes(const AuthoredCaptionFilter& f, std::string_view label) {
  const std::string_view value = f.value;
  const auto between = [&]() {
    return CaseInsensitiveCompare(label, value) >= 0 && CaseInsensitiveCompare(label, f.value_high) <= 0;
  };
  switch (f.predicate) {
    case CaptionPredicate::Equal:
      return CaseInsensitiveEquals(label, value);
    case CaptionPredicate::NotEqual:
      return !CaseInsensitiveEquals(label, value);
    case CaptionPredicate::BeginsWith:
      return CaseInsensitiveBeginsWith(label, value);
    case CaptionPredicate::NotBeginsWith:
      return !CaseInsensitiveBeginsWith(label, value);
    case CaptionPredicate::EndsWith:
      return CaseInsensitiveEndsWith(label, value);
    case CaptionPredicate::NotEndsWith:
      return !CaseInsensitiveEndsWith(label, value);
    case CaptionPredicate::Contains:
      return CaseInsensitiveContains(label, value);
    case CaptionPredicate::NotContains:
      return !CaseInsensitiveContains(label, value);
    case CaptionPredicate::GreaterThan:
      return CaseInsensitiveCompare(label, value) > 0;
    case CaptionPredicate::GreaterThanOrEqual:
      return CaseInsensitiveCompare(label, value) >= 0;
    case CaptionPredicate::LessThan:
      return CaseInsensitiveCompare(label, value) < 0;
    case CaptionPredicate::LessThanOrEqual:
      return CaseInsensitiveCompare(label, value) <= 0;
    case CaptionPredicate::Between:
      return between();
    case CaptionPredicate::NotBetween:
      return !between();
  }
  return true;
}

// Resolves the cache value a manual-filter item binds to, using the same
// index `resolve_pivot_names` reads when it derives `PivotItem::name`, so the
// two cannot disagree about which shared item an entry denotes. Returns
// `nullptr` when the cache does not cover the field or the index falls outside
// `shared_items`: such an item has no trustworthy binding.
const Value* item_cache_value(const PivotCache& cache, std::size_t field_index, const PivotItem& item) {
  if (field_index >= cache.fields().size()) {
    return nullptr;
  }
  const std::vector<Value>& shared = cache.fields()[field_index].shared_items;
  if (item.cache_index >= shared.size()) {
    return nullptr;
  }
  return &shared[item.cache_index];
}

enum class ScoreAxis { Row, Col };

Value leaf_score(const PivotResult& result, std::size_t r, std::size_t c, std::size_t data_field_index) {
  if (r >= result.values.size() || c >= result.values[r].size() || data_field_index >= result.values[r][c].size()) {
    return Value::blank();
  }
  return result.values[r][c][data_field_index];
}

// Drops the entries of `items` whose index is marked false in `keep`,
// preserving the relative order of the survivors. An index past the end of
// `keep` is kept: the mask only describes the leaves the filter scored.
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

// Rewrites each leaf set in `leaf_sets` from the pre-filter leaf index space
// into the surviving-leaf one, dropping pruned leaves. Returns a per-set mask
// marking the sets that still cover at least one surviving leaf; a set that
// empties out belongs to a subtotal whose whole group was filtered away.
std::vector<bool> remap_leaf_sets(std::vector<std::vector<std::size_t>>& leaf_sets, const std::vector<bool>& keep) {
  std::vector<std::size_t> new_index(keep.size(), 0);
  std::size_t surviving = 0;
  for (std::size_t i = 0; i < keep.size(); ++i) {
    new_index[i] = surviving;
    if (keep[i]) {
      ++surviving;
    }
  }
  std::vector<bool> keep_set(leaf_sets.size(), false);
  for (std::size_t s = 0; s < leaf_sets.size(); ++s) {
    std::vector<std::size_t>& set = leaf_sets[s];
    std::vector<std::size_t> remapped;
    remapped.reserve(set.size());
    for (const std::size_t leaf : set) {
      if (leaf < keep.size() && keep[leaf]) {
        remapped.push_back(new_index[leaf]);
      }
    }
    keep_set[s] = !remapped.empty();
    set = std::move(remapped);
  }
  return keep_set;
}

AxisScores score_axis(const PivotResult& result, ScoreAxis score_axis, std::size_t axis_count,
                      std::size_t cross_axis_count, std::size_t data_field_index) {
  AxisScores axis{{}, {}};
  axis.scores.assign(axis_count, 0.0);
  axis.all_blank.assign(axis_count, true);
  for (std::size_t i = 0; i < axis_count; ++i) {
    for (std::size_t j = 0; j < cross_axis_count; ++j) {
      const std::size_t r = score_axis == ScoreAxis::Row ? i : j;
      const std::size_t c = score_axis == ScoreAxis::Row ? j : i;
      if (auto n = numeric_aggregate_value(leaf_score(result, r, c, data_field_index))) {
        axis.scores[i] += *n;
        axis.all_blank[i] = false;
      }
    }
  }
  return axis;
}

}  // namespace

bool record_passes_manual_filter(const PivotTable& table, const PivotCache& cache, const PivotCacheRecord& record) {
  for (std::size_t fi = 0; fi < table.fields().size(); ++fi) {
    const PivotField& field = table.fields()[fi];
    if (field.items.empty()) {
      continue;
    }
    const Value v = cell_value(cache, record, fi);
    for (const PivotItem& item : field.items) {
      if (item.visible) {
        continue;
      }
      if (item.name.empty()) {
        // The one axis item with no label of its own is the blank: it binds to
        // a cache value that renders to nothing, so `resolve_pivot_names`
        // leaves the name empty. Identify it by that binding rather than by the
        // label it is drawn with — the placeholder is an ordinary string a
        // genuine text value is free to spell identically, and matching on it
        // would hide that value too. An item whose binding does not resolve is
        // merely malformed and must not filter anything.
        const Value* bound = item_cache_value(cache, fi, item);
        if (bound != nullptr && bound->is_blank() && v.is_blank()) {
          return false;
        }
        continue;
      }
      if (with_display_string(v, [&](std::string_view name) { return item.name == name; })) {
        return false;
      }
    }
  }
  for (const PivotFilter& f : table.active_filters()) {
    auto fi_or = resolve_field_by_any_name(table, f.field_name);
    if (!fi_or) {
      continue;
    }
    const Value v = cell_value(cache, record, *fi_or);
    if (f.type == FilterType::LabelDate) {
      // Date-range filters need the underlying numeric serial; the
      // rendered label string would lose precision and locale.
      if (!label_date_filter_passes(f, v)) {
        return false;
      }
      continue;
    }
    if (!with_display_string(v, [&](std::string_view label) { return label_filter_passes(f, label); })) {
      return false;
    }
  }
  // Authored `<filters>` caption entries, decoded by the OOXML reader.
  // `field_index` is the source `fld` attribute, and OOXML keeps
  // `<pivotFields>` parallel to the bound cache's fields, so it indexes
  // `table.fields()` directly instead of resolving through a name. An
  // index past the end is a malformed definition and filters nothing.
  for (const AuthoredCaptionFilter& f : table.authored_caption_filters()) {
    if (f.field_index >= table.fields().size()) {
      continue;
    }
    const Value v = cell_value(cache, record, f.field_index);
    if (!with_display_string(v, [&](std::string_view label) { return caption_filter_passes(f, label); })) {
      return false;
    }
  }
  return true;
}

AxisScores score_row_axis(const PivotResult& result, std::size_t row_count, std::size_t col_count,
                          std::size_t data_field_index) {
  return score_axis(result, ScoreAxis::Row, row_count, col_count, data_field_index);
}

AxisScores score_col_axis(const PivotResult& result, std::size_t col_count, std::size_t row_count,
                          std::size_t data_field_index) {
  return score_axis(result, ScoreAxis::Col, col_count, row_count, data_field_index);
}

std::optional<std::vector<bool>> build_value_filter_keep(const PivotFilter& f, const AxisScores& axis) {
  const std::size_t n = axis.scores.size();
  std::vector<bool> keep(n, false);
  if (f.type == FilterType::ValueTop10) {
    // The filter value is embedder-supplied, so saturate it instead of
    // narrowing it blind: a request larger than the axis keeps every leaf,
    // while a negative one — and NaN, which loses every comparison — keeps
    // none.
    const double requested = filter_number_value(f);
    const std::size_t top_n =
        index_from_double(requested, n + 1U).value_or(requested > static_cast<double>(n) ? n : 0U);
    // Sort indices by score descending; all-blank leaves sink to the
    // bottom regardless of N.
    std::vector<std::size_t> order(n);
    for (std::size_t i = 0; i < n; ++i) {
      order[i] = i;
    }
    std::sort(order.begin(), order.end(), [&](std::size_t a, std::size_t b) {
      if (axis.all_blank[a] != axis.all_blank[b]) {
        return !axis.all_blank[a];
      }
      return axis.scores[a] > axis.scores[b];
    });
    const std::size_t k = std::min(top_n, n);
    for (std::size_t i = 0; i < k; ++i) {
      if (!axis.all_blank[order[i]]) {
        keep[order[i]] = true;
      }
    }
    return keep;
  }
  if (f.type == FilterType::ValueGreaterThan) {
    const double threshold = filter_number_value(f);
    for (std::size_t i = 0; i < n; ++i) {
      if (!axis.all_blank[i] && axis.scores[i] > threshold) {
        keep[i] = true;
      }
    }
    return keep;
  }
  if (f.type == FilterType::ValueBetween) {
    const double lo = filter_number_value(f);
    const auto hi_or = filter_number_value_high(f);
    if (!hi_or) {
      return std::nullopt;  // Unbounded above -> no-op for this filter.
    }
    for (std::size_t i = 0; i < n; ++i) {
      if (!axis.all_blank[i] && axis.scores[i] >= lo && axis.scores[i] <= *hi_or) {
        keep[i] = true;
      }
    }
    return keep;
  }
  return std::nullopt;
}

void compact_leaf_axis(PivotResult& result, const std::vector<bool>& keep, LeafAxis axis,
                       std::vector<std::vector<std::size_t>>& row_subtotal_leaf_sets,
                       std::vector<std::vector<std::size_t>>& col_subtotal_leaf_sets) {
  if (axis == LeafAxis::Row) {
    keep_marked(result.values, keep);
    keep_marked(result.row_leaf_totals, keep);
    // A column subtotal renders one value per row leaf, so its matrix is
    // indexed by the axis being pruned.
    for (ColSubtotal& col_subtotal : result.col_subtotals) {
      keep_marked(col_subtotal.values, keep);
    }
    const std::vector<bool> keep_subtotal = remap_leaf_sets(row_subtotal_leaf_sets, keep);
    keep_marked(result.row_subtotals, keep_subtotal);
    keep_marked(result.subtotals, keep_subtotal);
    keep_marked(row_subtotal_leaf_sets, keep_subtotal);
    return;
  }
  for (auto& row_slot : result.values) {
    keep_marked(row_slot, keep);
  }
  keep_marked(result.col_leaf_totals, keep);
  // A row subtotal renders one value per column leaf.
  for (RowSubtotal& row_subtotal : result.row_subtotals) {
    keep_marked(row_subtotal.col_values, keep);
  }
  const std::vector<bool> keep_subtotal = remap_leaf_sets(col_subtotal_leaf_sets, keep);
  keep_marked(result.col_subtotals, keep_subtotal);
  keep_marked(col_subtotal_leaf_sets, keep_subtotal);
  // The row x column subtotal intersections are indexed by column subtotal.
  for (RowSubtotal& row_subtotal : result.row_subtotals) {
    keep_marked(row_subtotal.col_subtotal_values, keep_subtotal);
  }
}

}  // namespace formulon::pivot
