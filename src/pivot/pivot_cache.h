// Copyright 2026 libraz. Licensed under the MIT License.
//
// In-memory representation of an OOXML pivot cache (pair of
// `xl/pivotCache/cacheDefinition*.xml` and `cacheRecords*.xml`). The cache
// is a snapshot of the source data taken at refresh time; the pivot
// evaluator consumes it without reaching back into the source workbook.
//
// See backup/plans/15-pivot-and-advanced.md §15.1.6.

#ifndef FORMULON_PIVOT_PIVOT_CACHE_H_
#define FORMULON_PIVOT_PIVOT_CACHE_H_

#include <cstdint>
#include <string>
#include <vector>

#include "value.h"

namespace formulon::pivot {

/// One column of the pivot cache.
///
/// `shared_items` enumerates the distinct values for a discrete (string /
/// boolean) column. Range-typed (numeric / date) columns leave it empty
/// and store inline values on each `PivotCacheRecord`.
struct PivotCacheField {
  std::string name;
  std::vector<Value> shared_items;
};

/// One row of the pivot cache.
///
/// Each `cells` entry is either a `Number` index into the corresponding
/// field's `shared_items` (when that field is shared) or an inline `Value`
/// (when the field is range-typed). Encoding the index as a `Number` keeps
/// the `Value` shape uniform; consumers branch on whether the matching
/// field's `shared_items` is empty.
struct PivotCacheRecord {
  std::vector<Value> cells;
};

/// Owning container for the cache definition + records. Held by the
/// workbook (one cache may back multiple pivot tables) and accessed by
/// the pivot evaluator and OOXML round-trip path.
class PivotCache {
 public:
  PivotCache() = default;

  PivotCache(const PivotCache&) = delete;
  PivotCache& operator=(const PivotCache&) = delete;
  PivotCache(PivotCache&&) = default;
  PivotCache& operator=(PivotCache&&) = default;

  std::uint32_t cache_id() const { return cache_id_; }
  void set_cache_id(std::uint32_t id) { cache_id_ = id; }

  const std::vector<PivotCacheField>& fields() const { return fields_; }
  std::vector<PivotCacheField>& mutable_fields() { return fields_; }

  const std::vector<PivotCacheRecord>& records() const { return records_; }
  std::vector<PivotCacheRecord>& mutable_records() { return records_; }

 private:
  std::uint32_t cache_id_ = 0;
  std::vector<PivotCacheField> fields_;
  std::vector<PivotCacheRecord> records_;
};

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_PIVOT_CACHE_H_
