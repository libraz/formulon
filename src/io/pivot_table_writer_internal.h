//
// Internal seam between the pivot-table-definition writer and its
// `<filters>` emission (`pivot_filter_writer.cpp`). Not part of the public API.

#ifndef FORMULON_IO_PIVOT_TABLE_WRITER_INTERNAL_H_
#define FORMULON_IO_PIVOT_TABLE_WRITER_INTERNAL_H_

#include <cstdint>
#include <optional>
#include <string>

#include "pivot/pivot_table.h"

namespace formulon::io::detail {

/// The `<pivotFields>` index a public filter is written against, or
/// `nullopt` when it designates nothing a file can name.
std::optional<std::uint32_t> FilterFieldIndex(const pivot::PivotTable& table, const pivot::PivotFilter& filter);

/// True when a decoded filter is written back as its source bytes (see the
/// definition).
bool ReEmitsSource(const pivot::PivotTable& table, const pivot::PivotFilter& filter);

/// Emits `<filters>` (see the definition).
void AppendFilters(std::string& out, const pivot::PivotTable& table);

}  // namespace formulon::io::detail

#endif  // FORMULON_IO_PIVOT_TABLE_WRITER_INTERNAL_H_
