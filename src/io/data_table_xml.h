//
// The `<f t="dataTable">` element: the attributes of a What-If data table
// and their A1 text form, shared by the DOM and SAX sheet readers and the
// sheet writer.

#ifndef FORMULON_IO_DATA_TABLE_XML_H_
#define FORMULON_IO_DATA_TABLE_XML_H_

#include <string>
#include <string_view>

#include "data_table.h"

namespace formulon::io {

/// The raw attribute values of one `<f t="dataTable">`; an absent attribute
/// is the empty view.
struct DataTableAttrs {
  std::string_view ref;
  std::string_view r1;
  std::string_view r2;
  std::string_view dt2d;
  std::string_view dtr;
  std::string_view del1;
  std::string_view del2;
  std::string_view ca;
  std::string_view aca;
};

/// Builds the table from `attrs`. False when `ref` is missing, not a single
/// range, or covers more than `kMaxDataTableCells`, in which case nothing is
/// modelled.
bool parse_data_table(const DataTableAttrs& attrs, DataTable* out);

/// Appends `<f t="dataTable" .../>` for `table`.
void append_data_table_formula(std::string& out, const DataTable& table);

}  // namespace formulon::io

#endif  // FORMULON_IO_DATA_TABLE_XML_H_
