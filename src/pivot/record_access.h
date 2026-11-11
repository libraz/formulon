//
// Cache-record value extraction. Pulled out so the filter engine, the
// hierarchy builder, and the aggregator can all index a record's field
// through the same code path.

#ifndef FORMULON_PIVOT_RECORD_ACCESS_H_
#define FORMULON_PIVOT_RECORD_ACCESS_H_

#include <cstddef>
#include <optional>

#include "pivot/pivot_cache.h"
#include "utils/checked_index.h"
#include "value.h"

namespace formulon::pivot {

/// True when cell `i` of `record` holds an index into `field.shared_items`.
/// The captured per-cell flag decides when present; otherwise a numeric cell of
/// a shared field is an index unless it is marked a date.
inline bool cell_is_shared_index(const PivotCacheRecord& record, const PivotCacheField& field, std::size_t i) {
  if (i < record.cell_is_index.size()) {
    return record.cell_is_index[i];
  }
  return i < record.cells.size() && !field.shared_items.empty() && record.cells[i].is_number() &&
         !(i < record.cell_is_date.size() && record.cell_is_date[i]);
}

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
  // An inline cell is returned verbatim even inside a shared field; an index
  // cell is always resolved against `shared_items`.
  if (!cell_is_shared_index(record, field, field_index)) {
    return cell;
  }
  if (!cell.is_number()) {
    return cell;  // Defensive: an index flag on a non-numeric cell.
  }
  // A hand-built cache carries whatever number the embedder stored, so the
  // index is validated against `shared_items` before it is narrowed rather
  // than after. Negative, NaN, infinite and oversized indices all collapse to
  // Blank like any other out-of-range reference here.
  const std::optional<std::size_t> index = index_from_double(cell.as_number(), field.shared_items.size());
  if (!index.has_value()) {
    return Value::blank();
  }
  return field.shared_items[*index];
}

/// True when `(record, field_index)` holds a number its source spelled as a date: the shared item's
/// mark for an index cell, the cell's own mark otherwise.
inline bool cell_is_date_value(const PivotCache& cache, const PivotCacheRecord& record, std::size_t field_index) {
  if (field_index >= cache.fields().size() || field_index >= record.cells.size()) {
    return false;
  }
  const auto& field = cache.fields()[field_index];
  if (cell_is_shared_index(record, field, field_index)) {
    const Value& cell = record.cells[field_index];
    const std::optional<std::size_t> index =
        cell.is_number() ? index_from_double(cell.as_number(), field.shared_items.size()) : std::nullopt;
    return index.has_value() && *index < field.shared_item_is_date.size() && field.shared_item_is_date[*index];
  }
  return field_index < record.cell_is_date.size() && record.cell_is_date[field_index];
}

/// True when every present (non-blank) value of `field_index` is a date-marked number, which is what
/// makes a field date-typed. A field holding no value is not contradicted.
inline bool field_holds_only_dates(const PivotCache& cache, std::size_t field_index) {
  for (const PivotCacheRecord& record : cache.records()) {
    if (!cell_value(cache, record, field_index).is_blank() && !cell_is_date_value(cache, record, field_index)) {
      return false;
    }
  }
  return true;
}

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_RECORD_ACCESS_H_
