//
// What the xlsx writers and the .xlsb pivot bridges write for a PivotTable or
// cache built in memory, where no root attribute was captured from a file:
// the values Excel keeps when it re-saves such a part. Not part of the public API.

#ifndef FORMULON_IO_PIVOT_BUILT_DEFAULTS_H_
#define FORMULON_IO_PIVOT_BUILT_DEFAULTS_H_

#include <cstdint>
#include <string>
#include <string_view>

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

/// The pivotTableDefinition16 extension Excel's save of a built table adds after the x14 one.
inline constexpr std::string_view kBuiltTableDefinition16Ext =
    "<ext uri=\"{747A6164-185A-40DC-8AA5-F01512510D54}\" "
    "xmlns:xpdl=\"http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout\">"
    "<xpdl:pivotTableDefinition16/></ext>";

/// The cache extension Excel writes while its source range is unedited in the session.
inline constexpr std::string_view kDataSourceCleanUri = "{F4336242-56EE-4DD3-A5F7-E900642BF6E9}";
inline constexpr std::string_view kDataSourceCleanExt =
    "<ext uri=\"{F4336242-56EE-4DD3-A5F7-E900642BF6E9}\" "
    "xmlns:xlpdsc=\"http://schemas.microsoft.com/office/spreadsheetml/2026/pivotDataSourceClean\">"
    "<xlpdsc:dataSourceClean>1</xlpdsc:dataSourceClean></ext>";

/// A cache's trailing children `xml` with the dataSourceClean extension ending
/// its `<extLst>`, which is added when there is none.
inline std::string with_data_source_clean(std::string_view xml) {
  constexpr std::string_view kClose = "</extLst>";
  std::string out(xml);
  if (xml.size() >= kClose.size() && xml.substr(xml.size() - kClose.size()) == kClose) {
    out.insert(out.size() - kClose.size(), kDataSourceCleanExt);
  } else {
    out.append("<extLst>").append(kDataSourceCleanExt).append(kClose);
  }
  return out;
}

/// The cache extension naming the richPivotRecords part Excel saves a cache holding rich values with.
inline constexpr std::string_view kRichInfoUri = "{2C874A73-7782-4A18-856F-96AC7E287872}";

/// True when `cache` was read with its records in a richPivotRecords part: its captured extension list carries
/// richInfo, whose `r:id` names that part as rId1 of the definition's rels, and the root has no `r:id`.
inline bool cache_has_rich_records(const pivot::PivotCache& cache) {
  return cache.trailing_children_after_calculated_xml().find(kRichInfoUri) != std::string::npos;
}

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
