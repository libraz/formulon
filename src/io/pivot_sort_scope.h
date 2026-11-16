#ifndef FORMULON_IO_PIVOT_SORT_SCOPE_H_
#define FORMULON_IO_PIVOT_SORT_SCOPE_H_

// A pivot field sorted by a data field's values carries Excel's `<autoSortScope>`: a pivot area whose one reference
// names the data field by its position in `<dataFields>` (field 4294967294). The model holds the sort as the data
// field's name in `SortSpec::by_field`; these helpers move between that name and the scope both formats store.

#include <string>

#include "pivot/pivot_table.h"

namespace formulon {
namespace io {

/// Moves each field's stored scope that names one data field into `sort.by_field`. A scope of any other shape stays
/// in `raw_children_xml` as read.
void lift_pivot_sort_scopes(pivot::PivotTable& table);

/// The children written after a field's `<items>`: the scope of its value sort, when `sort.by_field` names a data
/// field of `table`, then the children kept as read.
std::string pivot_field_trailing_children(const pivot::PivotTable& table, const pivot::PivotField& field);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_PIVOT_SORT_SCOPE_H_
