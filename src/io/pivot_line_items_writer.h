//
// `<rowItems>` / `<colItems>` of an edited pivot table, serialized from the
// line items `pivot::layout` projects. A line names its items by identity; this
// writer turns each into its position among the `<items>` the table writer
// emits for that field, which is what `<x v>` addresses.

#ifndef FORMULON_IO_PIVOT_LINE_ITEMS_WRITER_H_
#define FORMULON_IO_PIVOT_LINE_ITEMS_WRITER_H_

#include <optional>

#include "io/pivot_cache_writer.h"
#include "io/pivot_table_writer.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_table.h"

namespace formulon::io {

/// Serializes `lines` into the `<rowItems>` and `<colItems>` Excel writes for
/// the projected report of `table`.
///
/// Each `<x v>` is the item's position in the `<items>`
/// `write_pivot_table_definition` emits for its field under `plan` (the cache
/// write plan of `cache`), `blank` markers included. Returns nullopt when an
/// axis of `lines` was refused, or when a line names an item with no identity
/// or one no written item holds; the table is then saved without lines.
///
/// `lines` point into the `PivotResult` they were projected from, which must
/// still be alive. The caller must hold an `eval::ScopedEvalProfile`: items
/// are matched with `pivot::value_less`.
std::optional<PivotItemsXml> write_pivot_line_items(const pivot::PivotTable& table, const pivot::PivotCache& cache,
                                                    const PivotCacheWritePlan* plan,
                                                    const pivot::PivotLineItems& lines);

}  // namespace formulon::io

#endif  // FORMULON_IO_PIVOT_LINE_ITEMS_WRITER_H_
