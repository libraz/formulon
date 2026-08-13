
#include "pivot/filter_engine.h"

#include <algorithm>
#include <array>
#include <charconv>
#include <cstddef>
#include <cstdio>
#include <limits>
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

// Invokes `consume` with the display label for `v`, avoiding the temporary
// std::string that the per-record filter path previously allocated. Numeric
// formatting intentionally follows `display_string`: integral values omit
// `.0`, while other values use the `%f` spelling used by std::to_string.
template <typename Consume>
bool with_display_string(const Value& v, Consume&& consume) {
  switch (v.kind()) {
    case ValueKind::Blank:
      return consume({});
    case ValueKind::Text:
      return consume(v.as_text());
    case ValueKind::Bool:
      return consume(v.as_boolean() ? std::string_view{"TRUE"} : std::string_view{"FALSE"});
    case ValueKind::Error:
      return consume(display_name(v.as_error()));
    case ValueKind::Number: {
      const double d = v.as_number();
      std::array<char, 512> buffer{};
      if (d >= static_cast<double>(std::numeric_limits<long long>::min()) &&
          d <= static_cast<double>(std::numeric_limits<long long>::max())) {
        const auto i = static_cast<long long>(d);
        if (static_cast<double>(i) == d) {
          const auto [end, ec] = std::to_chars(buffer.data(), buffer.data() + buffer.size(), i);
          if (ec != std::errc{}) {
            return false;
          }
          return consume(std::string_view(buffer.data(), static_cast<std::size_t>(end - buffer.data())));
        }
      }
      const int written = std::snprintf(buffer.data(), buffer.size(), "%f", d);
      if (written < 0 || static_cast<std::size_t>(written) >= buffer.size()) {
        return false;
      }
      return consume(std::string_view(buffer.data(), static_cast<std::size_t>(written)));
    }
    default:
      return consume({});
  }
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
      // An item whose cache index did not resolve has no trustworthy label.
      // Treating its empty name as a hidden label accidentally filters every
      // genuine blank source value instead of only the malformed item.
      if (item.name.empty()) {
        continue;
      }
      if (!item.visible && with_display_string(v, [&](std::string_view name) { return item.name == name; })) {
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
