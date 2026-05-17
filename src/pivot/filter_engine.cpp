// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.

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
#include "pivot/pivot_cache.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/record_access.h"
#include "pivot/value_order.h"
#include "value.h"

namespace formulon::pivot {
namespace {

// Resolves a filter `field_name` against `table.fields()`. Returns the
// field's index on match, `nullopt` if no field has that name. Cache
// fields and pivot fields share the same index in MVP, so this also
// indexes into the cache record.
std::optional<std::size_t> resolve_filter_field(const PivotTable& table, const std::string& name) {
  for (std::size_t i = 0; i < table.fields().size(); ++i) {
    if (table.fields()[i].source_name == name) {
      return i;
    }
  }
  return std::nullopt;
}

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
bool label_filter_passes(const PivotFilter& f, const std::string& label) {
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

Value leaf_score(const PivotResult& result, std::size_t r, std::size_t c) {
  if (r >= result.values.size() || c >= result.values[r].size() || result.values[r][c].empty()) {
    return Value::blank();
  }
  return result.values[r][c][0];
}

AxisScores score_axis(const PivotResult& result, ScoreAxis score_axis, std::size_t axis_count,
                      std::size_t cross_axis_count) {
  AxisScores axis{{}, {}};
  axis.scores.assign(axis_count, 0.0);
  axis.all_blank.assign(axis_count, true);
  for (std::size_t i = 0; i < axis_count; ++i) {
    for (std::size_t j = 0; j < cross_axis_count; ++j) {
      const std::size_t r = score_axis == ScoreAxis::Row ? i : j;
      const std::size_t c = score_axis == ScoreAxis::Row ? j : i;
      if (auto n = numeric_aggregate_value(leaf_score(result, r, c))) {
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
    const std::string name = display_string(v);
    for (const PivotItem& item : field.items) {
      if (!item.visible && item.name == name) {
        return false;
      }
    }
  }
  for (const PivotFilter& f : table.active_filters()) {
    auto fi_or = resolve_filter_field(table, f.field_name);
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
    const std::string label = display_string(v);
    if (!label_filter_passes(f, label)) {
      return false;
    }
  }
  return true;
}

AxisScores score_row_axis(const PivotResult& result, std::size_t row_count, std::size_t col_count) {
  return score_axis(result, ScoreAxis::Row, row_count, col_count);
}

AxisScores score_col_axis(const PivotResult& result, std::size_t col_count, std::size_t row_count) {
  return score_axis(result, ScoreAxis::Col, col_count, row_count);
}

std::optional<std::vector<bool>> build_value_filter_keep(const PivotFilter& f, const AxisScores& axis) {
  const std::size_t n = axis.scores.size();
  std::vector<bool> keep(n, false);
  if (f.type == FilterType::ValueTop10) {
    const auto top_n = static_cast<std::size_t>(filter_number_value(f));
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

void compact_row_axis_values(std::vector<std::vector<std::vector<Value>>>& values, const std::vector<bool>& keep) {
  std::vector<std::vector<std::vector<Value>>> new_values;
  new_values.reserve(values.size());
  for (std::size_t i = 0; i < values.size() && i < keep.size(); ++i) {
    if (keep[i]) {
      new_values.push_back(std::move(values[i]));
    }
  }
  values = std::move(new_values);
}

void compact_col_axis_values(std::vector<std::vector<std::vector<Value>>>& values, const std::vector<bool>& keep) {
  for (auto& row_slot : values) {
    std::vector<std::vector<Value>> new_row;
    new_row.reserve(row_slot.size());
    for (std::size_t c = 0; c < row_slot.size() && c < keep.size(); ++c) {
      if (keep[c]) {
        new_row.push_back(std::move(row_slot[c]));
      }
    }
    row_slot = std::move(new_row);
  }
}

}  // namespace formulon::pivot
