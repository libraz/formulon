// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Cache-record value extraction. Pulled out so the filter engine, the
// hierarchy builder, and the aggregator can all index a record's field
// through the same code path.

#ifndef FORMULON_PIVOT_RECORD_ACCESS_H_
#define FORMULON_PIVOT_RECORD_ACCESS_H_

#include <cstddef>

#include "pivot/pivot_cache.h"
#include "value.h"

namespace formulon::pivot {

/// Pulls the effective `Value` for `(record, field)`. When the field is
/// shared (`shared_items` non-empty), the record stores a `Number` index
/// into `shared_items`; otherwise the record stores the value inline.
/// Out-of-range record/field/index references collapse to `Blank` so the
/// rest of the evaluator can stay branch-free; cache-record corruption
/// is the OOXML reader's responsibility to surface.
inline Value cell_value(const PivotCache& cache, const PivotCacheRecord& record, std::size_t field_index) {
  if (field_index >= cache.fields().size() || field_index >= record.cells.size()) {
    return Value::blank();
  }
  const auto& field = cache.fields()[field_index];
  const Value& cell = record.cells[field_index];
  // Prefer the explicit per-cell encoding flag when the reader populated
  // it: an inline cell (`cell_is_index == false`) is returned verbatim
  // even inside a shared field, and an index cell is always resolved
  // against `shared_items`. When the flag vector is empty (hand-built
  // caches), fall back to inferring from the field being shared.
  const bool have_flag = field_index < record.cell_is_index.size();
  const bool is_index =
      have_flag ? record.cell_is_index[field_index] : (!field.shared_items.empty() && cell.is_number());
  if (!is_index) {
    return cell;
  }
  if (!cell.is_number()) {
    return cell;  // Defensive: an index flag on a non-numeric cell.
  }
  const double idx = cell.as_number();
  if (idx < 0.0) {
    return Value::blank();
  }
  const auto i = static_cast<std::size_t>(idx);
  if (i >= field.shared_items.size()) {
    return Value::blank();
  }
  return field.shared_items[i];
}

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_RECORD_ACCESS_H_
