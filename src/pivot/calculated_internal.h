//
// Helpers shared by the calculated-field / calculated-item edit operations
// (`calculated_edit.cpp`) and the evaluation plan (`calculated.cpp`).

#ifndef FORMULON_PIVOT_CALCULATED_INTERNAL_H_
#define FORMULON_PIVOT_CALCULATED_INTERNAL_H_

#include <cstdint>
#include <optional>
#include <string_view>

#include "pivot/pivot_types.h"

namespace formulon::pivot::detail {

inline std::string_view without_equals(std::string_view formula) {
  if (!formula.empty() && formula.front() == '=') {
    formula.remove_prefix(1);
  }
  return formula;
}

/// The shared item a `<calculatedItem>` entry defines: the first item its area
/// selects on its own field.
inline std::optional<std::uint32_t> entry_item(const PivotCalculatedItem& entry) {
  if (!entry.has_area) {
    return std::nullopt;
  }
  for (const PivotAreaReference& reference : entry.area.references) {
    if (reference.field == entry.cache_field_index && !reference.items.empty()) {
      return reference.items.front();
    }
  }
  return std::nullopt;
}

}  // namespace formulon::pivot::detail

#endif  // FORMULON_PIVOT_CALCULATED_INTERNAL_H_
