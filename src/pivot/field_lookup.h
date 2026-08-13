//
// Pivot field name resolution. Pulled out so the filter engine, the
// value-sort resolver, and the public mutators all accept the same set of
// spellings for a field: its source name, its display (custom) name, or the
// display name of a data field aggregating it.

#ifndef FORMULON_PIVOT_FIELD_LOOKUP_H_
#define FORMULON_PIVOT_FIELD_LOOKUP_H_

#include <cstddef>
#include <optional>
#include <string_view>

#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"

namespace formulon::pivot {

/// True when `name` designates `field` directly: either its `source_name` or,
/// when the field was renamed, its non-empty `custom_name` (the display name
/// layout renders).
inline bool pivot_field_has_name(const PivotField& field, std::string_view name) {
  return field.source_name == name || (!field.custom_name.empty() && field.custom_name == name);
}

/// Resolves `name` to an index into `table.fields()`. Direct field names win
/// over data-field display names, so a data field named after another pivot
/// field cannot shadow it. Returns `nullopt` when nothing matches.
inline std::optional<std::size_t> resolve_field_by_any_name(const PivotTable& table, std::string_view name) {
  for (std::size_t i = 0; i < table.fields().size(); ++i) {
    if (pivot_field_has_name(table.fields()[i], name)) {
      return i;
    }
  }
  for (const PivotDataField& data_field : table.data_fields()) {
    if (data_field.name == name && data_field.field_index < table.fields().size()) {
      return static_cast<std::size_t>(data_field.field_index);
    }
  }
  return std::nullopt;
}

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_FIELD_LOOKUP_H_
