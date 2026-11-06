//
// Internal seam between the pivot-table-definition writer and its
// `<filters>` emission (`pivot_filter_writer.cpp`) and `<pivotField>` emission
// (`pivot_field_writer.cpp`). Not part of the public API.

#ifndef FORMULON_IO_PIVOT_TABLE_WRITER_INTERNAL_H_
#define FORMULON_IO_PIVOT_TABLE_WRITER_INTERNAL_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/xml_utils.h"
#include "pivot/pivot_table.h"

namespace formulon::io::detail {

using AttrList = std::vector<std::pair<std::string, std::string>>;

/// Appends `attrs` in the order of `order`; names outside it follow in their
/// given order.
template <std::size_t N>
inline void AppendAttrsInOrder(std::string& out, const AttrList& attrs, const std::string_view (&order)[N]) {
  AttrList sorted;
  sorted.reserve(attrs.size());
  std::vector<bool> taken(attrs.size(), false);
  for (const std::string_view name : order) {
    for (std::size_t i = 0; i < attrs.size(); ++i) {
      if (!taken[i] && attrs[i].first == name) {
        taken[i] = true;
        sorted.push_back(attrs[i]);
      }
    }
  }
  for (std::size_t i = 0; i < attrs.size(); ++i) {
    if (!taken[i]) {
      sorted.push_back(attrs[i]);
    }
  }
  append_raw_attrs(out, sorted);
}

/// True when `attrs` already carries `name` (an attribute captured on read).
inline bool HasRawAttr(const std::vector<std::pair<std::string, std::string>>& attrs, std::string_view name) {
  for (const auto& attr : attrs) {
    if (attr.first == name) {
      return true;
    }
  }
  return false;
}

/// True when `attrs` carries any of the report-form attributes the reader
/// keeps only as an override of the form `layout` derives.
inline bool HasFormOverride(const std::vector<std::pair<std::string, std::string>>& attrs) {
  return HasRawAttr(attrs, "compact") || HasRawAttr(attrs, "outline") || HasRawAttr(attrs, "compactData") ||
         HasRawAttr(attrs, "outlineData");
}

/// The `<pivotFields>` index a public filter is written against, or
/// `nullopt` when it designates nothing a file can name.
std::optional<std::uint32_t> FilterFieldIndex(const pivot::PivotTable& table, const pivot::PivotFilter& filter);

/// True when a decoded filter is written back as its source bytes (see the
/// definition).
bool ReEmitsSource(const pivot::PivotTable& table, const pivot::PivotFilter& filter);

/// Emits `<filters>` (see the definition).
void AppendFilters(std::string& out, const pivot::PivotTable& table);

/// Emits one `<pivotField>` (see the definition).
void AppendPivotField(std::string& out, const pivot::PivotField& field, const std::vector<pivot::PivotItem>& items,
                      pivot::PivotLayout layout, bool used_as_data, bool measure_filter);

}  // namespace formulon::io::detail

#endif  // FORMULON_IO_PIVOT_TABLE_WRITER_INTERNAL_H_
