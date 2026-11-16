#include "io/pivot_sort_scope.h"

#include <cstddef>
#include <optional>
#include <string>

#include "pivot/field_lookup.h"

namespace formulon {
namespace io {

namespace {

/// Excel's scope for a sort by data field `index`, as it writes it.
std::string ScopeXml(std::size_t index) {
  return "<autoSortScope><pivotArea dataOnly=\"0\" outline=\"0\" fieldPosition=\"0\"><references count=\"1\">"
         "<reference field=\"4294967294\" count=\"1\" selected=\"0\"><x v=\"" +
         std::to_string(index) + "\"/></reference></references></pivotArea></autoSortScope>";
}

}  // namespace

void lift_pivot_sort_scopes(pivot::PivotTable& table) {
  for (pivot::PivotField& field : table.mutable_fields()) {
    if (field.sort.manual || field.raw_children_xml.empty()) {
      continue;
    }
    for (std::size_t i = 0; i < table.data_fields().size(); ++i) {
      const std::string scope = ScopeXml(i);
      const std::string& name = table.data_fields()[i].name;
      // Kept as read unless the name leads back to this data field when the scope is written.
      if (!name.empty() && field.raw_children_xml.compare(0, scope.size(), scope) == 0 &&
          pivot::find_value_sort_data_field(table, name) == i) {
        field.sort.by_field = name;
        field.raw_children_xml.erase(0, scope.size());
        break;
      }
    }
  }
}

std::string pivot_field_trailing_children(const pivot::PivotTable& table, const pivot::PivotField& field) {
  std::string out;
  if (!field.sort.manual && !field.sort.by_field.empty()) {
    if (const std::optional<std::size_t> index = pivot::find_value_sort_data_field(table, field.sort.by_field)) {
      out = ScopeXml(*index);
    }
  }
  out.append(field.raw_children_xml);
  return out;
}

}  // namespace io
}  // namespace formulon
