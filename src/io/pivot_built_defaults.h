//
// What the xlsx writers and the .xlsb pivot bridges write for a PivotTable or
// cache built in memory, where no root attribute was captured from a file:
// the values Excel keeps when it re-saves such a part. Not part of the public API.

#ifndef FORMULON_IO_PIVOT_BUILT_DEFAULTS_H_
#define FORMULON_IO_PIVOT_BUILT_DEFAULTS_H_

#include <cstdint>

#include "io/pivot_table_writer_internal.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"

namespace formulon::io {

/// The cache version triple Excel writes on a cache it did not create.
inline constexpr std::uint8_t kBuiltCacheCreatedVersion = 8;
inline constexpr std::uint8_t kBuiltCacheRefreshedVersion = 8;
inline constexpr std::uint8_t kBuiltCacheMinRefreshableVersion = 3;

/// A table's `updatedVersion` when none was captured; without it Excel adds a
/// stray caption row above the pivot on refresh.
inline constexpr std::uint8_t kBuiltTableUpdatedVersion = 8;

/// True for a cache built in memory: no root attribute was captured on read.
inline bool cache_built_in_memory(const pivot::PivotCache& cache) {
  return cache.passthrough_attrs().empty();
}

/// True when Excel must rebuild `cache` on open: built in memory and its
/// records filled by hand, so they miss what a refresh reads from the source.
inline bool built_cache_refreshes_on_load(const pivot::PivotCache& cache) {
  return cache_built_in_memory(cache) && !cache.records_from_source();
}

/// `hideValuesRow` for a table with no captured extension list: its own value,
/// else on for a table with no captured `updatedVersion`.
inline bool table_hides_values_row(const pivot::PivotTable& table) {
  return table.hide_values_row().value_or(!detail::HasRawAttr(table.passthrough_attrs(), "updatedVersion"));
}

}  // namespace formulon::io

#endif  // FORMULON_IO_PIVOT_BUILT_DEFAULTS_H_
