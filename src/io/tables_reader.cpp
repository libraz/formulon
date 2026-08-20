//
// Implementation of the tables-part reader. See tables_reader.h for the
// public contract.

#include "io/tables_reader.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace {

std::string CaptureExtraAttrs(const pugi::xml_node& node, bool root) {
  std::string out;
  for (pugi::xml_attribute attr : node.attributes()) {
    const std::string_view name(attr.name());
    const bool is_known_column =
        name == "id" || name == "name" || name == "totalsRowLabel" || name == "totalsRowFunction";
    const bool is_known_root = name == "id" || name == "name" || name == "displayName" || name == "ref" ||
                               name == "headerRowCount" || name == "totalsRowCount" || name == "xmlns";
    const bool is_namespace = name.rfind("xmlns:", 0U) == 0U || name == "mc:Ignorable";
    if ((root && is_known_root) || (!root && (is_known_column || is_namespace))) {
      continue;
    }
    // Verbatim-retained attributes go through the same escape rule as
    // modelled ones, so a captured value cannot be written in a weaker form
    // than the tag it is spliced into.
    append_xml_attr(out, name, attr.value());
  }
  return out;
}

}  // namespace

Expected<TableMetadata, Error> read_table(const std::vector<std::uint8_t>& table_bytes, std::size_t sheet_index) {
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_xml_buffer(doc, table_bytes, "tables_reader", "table*.xml"));
  pugi::xml_node root = doc.child("table");
  if (!root) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid, "table*.xml: missing <table> root",
                      "context=tables_reader");
  }

  TableMetadata table;
  table.sheet_index = sheet_index;
  table.id = attr_u32(root, "id");
  table.name = attr_str(root, "name");
  table.display_name = attr_str(root, "displayName");
  table.root_extra_attrs = CaptureExtraAttrs(root, /*root=*/true);

  // `ref` is required; an absent / empty value rejects the part.
  const std::string_view ref_v = attr_str(root, "ref");
  if (ref_v.empty()) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "table*.xml: missing required ref attribute",
                      "context=tables_reader");
  }
  table.ref = ref_v;

  // headerRowCount defaults to 1 in the OOXML spec — i.e. tables have a
  // header row by default. Only an explicit "0" disables it.
  if (pugi::xml_attribute hrc = root.attribute("headerRowCount"); hrc) {
    if (attr_u32(root, "headerRowCount", 1U) == 0U) {
      table.header_row = false;
    }
  }
  // totalsRowCount defaults to 0 (no totals row).
  if (attr_u32(root, "totalsRowCount") >= 1U) {
    table.totals_row = true;
  }

  if (pugi::xml_node cols_node = root.child("tableColumns"); cols_node) {
    for (pugi::xml_node col = cols_node.child("tableColumn"); col; col = col.next_sibling("tableColumn")) {
      TableColumn entry;
      entry.id = attr_u32(col, "id");
      entry.name = attr_str(col, "name");
      entry.totals_label = attr_str(col, "totalsRowLabel");
      entry.totals_function = attr_str(col, "totalsRowFunction");
      entry.extra_attrs = CaptureExtraAttrs(col, /*root=*/false);
      // Preserve <calculatedColumnFormula> verbatim. Children come
      // after attributes per the OOXML schema, so this lookup runs
      // last for the column. pugixml's `text().as_string()` already
      // returns the unescaped PCDATA payload.
      if (pugi::xml_node fn_node = col.child("calculatedColumnFormula"); fn_node) {
        entry.calculated_column_formula = fn_node.text().as_string();
      }
      table.columns.push_back(std::move(entry));
    }
  }

  // `<tableStyleInfo>` (style name + banded-row flags) follows
  // `<tableColumns>`. Captured verbatim; the engine does not model table
  // styles but must not drop them on save.
  if (pugi::xml_node style_info = root.child("tableStyleInfo"); style_info) {
    table.table_style_info_xml = raw_xml(style_info);
  }
  if (pugi::xml_node auto_filter = root.child("autoFilter"); auto_filter) {
    table.auto_filter_xml = raw_xml(auto_filter);
  }
  if (pugi::xml_node sort_state = root.child("sortState"); sort_state) {
    table.sort_state_xml = raw_xml(sort_state);
  }
  if (pugi::xml_node ext_lst = root.child("extLst"); ext_lst) {
    table.ext_lst_xml = raw_xml(ext_lst);
  }

  return table;
}

}  // namespace io
}  // namespace formulon
