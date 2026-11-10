//
// Moving a pivot table's definition onto a cache whose fields or shared items
// changed: a refresh rebuilds both, a removed calculated field drops one.
//
// Index-keyed state follows the field and item maps: the field list, the axis
// orders, page fields and their selected item, data fields and their show-as
// base, filters, pivot items. Name-keyed state (`source_name`, a sort's
// `by_field`, a filter's `field_name`) follows a field's new spelling. State
// on a dropped field goes with it.

#ifndef FORMULON_PIVOT_FIELD_REMAP_H_
#define FORMULON_PIVOT_FIELD_REMAP_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"

namespace formulon::pivot {

/// Where each field and item of a cache went.
struct FieldRemap {
  /// Per old field, its new index; empty when the field is gone.
  std::vector<std::optional<std::uint32_t>> fields;
  /// Per old field, per old item (a shared item, or a group item of a group
  /// field), its new index; empty when the item is gone. A field without an
  /// entry, or with an empty one, keeps its items as they are.
  std::vector<std::vector<std::optional<std::uint32_t>>> items;
  /// Per new field, the spelling it now has. Empty keeps every name.
  std::vector<std::string> names;
};

/// Rewrites `table` for `cache`, the cache `remap` leads to.
///
/// Fields keep their placement at their new index; a new index no old field
/// reaches gets an unplaced field. A dropped field leaves the axes, its data
/// fields and the filters on it go, and a show-as base on it is cleared.
///
/// A field's items keep their order with indices remapped; an item whose
/// shared item is gone is removed, one kept as missing is marked `m="1"`, and
/// every shared item no item names (missing and calculated ones aside) is
/// appended, hidden when the field still hides an item. A field sorted
/// ascending or descending is re-sorted, missing items included. Positions
/// into a field's items (a page field's item, a show-as base item) follow.
///
/// Sorting collates under the caller's `eval::ScopedEvalProfile`.
void remap_table(PivotTable& table, const PivotCache& cache, const FieldRemap& remap);

/// The table entry of cache field `field` on no axis. A calculated field's
/// entry cannot be dragged to an axis.
PivotField unplaced_pivot_field(const PivotCacheField& field);

/// Position of `field` on its axis in `table` (rows, columns, then pages), or
/// empty when it is on none.
std::optional<std::uint32_t> axis_position(const PivotTable& table, std::size_t field);

/// Sets each calculated item's `<pivotArea fieldPosition>` to its field's
/// place on `table`'s axes, where the table places it.
void place_calculated_item_areas(PivotCache& cache, const PivotTable& table);

/// Forgets what was derived from the table's definition: its last result,
/// the span read from a file, and the report items written with it.
void clear_derived_state(PivotTable& table);

/// Drops the location offsets read with a span Excel wrote once the span is no
/// longer authored and the report has been projected; the writer then derives
/// them from the model.
void drop_derived_location(PivotTable& table);

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_FIELD_REMAP_H_
