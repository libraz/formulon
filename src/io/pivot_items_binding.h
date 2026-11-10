//
// The `<items>` the pivot-table writer emits for a field, shared by the
// `<pivotField>` writer and the `<rowItems>` / `<colItems>` serializer, which
// addresses those items by position.

#ifndef FORMULON_IO_PIVOT_ITEMS_BINDING_H_
#define FORMULON_IO_PIVOT_ITEMS_BINDING_H_

#include <vector>

#include "io/pivot_cache_writer.h"
#include "pivot/pivot_table.h"

namespace formulon::io::detail {

/// The items of a field the plan derives as discrete, bound to the derived
/// shared items; `field.items` verbatim for any other field.
///
/// An axis field with no authored items gets one per shared item, in its sort
/// order. Authored items resolve to a shared item by their cache index, else
/// by label (as `fm_sheet_add_pivot_field_item` binds), else by position;
/// shared items none of them names follow, visible.
///
/// `bound_by_position`, when non-null, receives one flag per returned item:
/// true for an authored item that neither a cache index nor a label bound.
std::vector<pivot::PivotItem> BoundItems(const pivot::PivotField& field, const PivotCacheFieldWritePlan* fp,
                                         std::vector<bool>* bound_by_position = nullptr);

}  // namespace formulon::io::detail

#endif  // FORMULON_IO_PIVOT_ITEMS_BINDING_H_
