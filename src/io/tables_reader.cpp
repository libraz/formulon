// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the tables-part reader. See tables_reader.h for the
// public contract.

#include "io/tables_reader.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {

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
  table.id = parse_xml_u32_attr(root.attribute("id"), 0U);
  if (pugi::xml_attribute name_attr = root.attribute("name"); name_attr) {
    table.name = name_attr.value();
  }
  if (pugi::xml_attribute disp_attr = root.attribute("displayName"); disp_attr) {
    table.display_name = disp_attr.value();
  }

  pugi::xml_attribute ref_attr = root.attribute("ref");
  if (!ref_attr || ref_attr.value() == nullptr || *ref_attr.value() == '\0') {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "table*.xml: missing required ref attribute",
                      "context=tables_reader");
  }
  table.ref = ref_attr.value();

  // headerRowCount defaults to 1 in the OOXML spec — i.e. tables have a
  // header row by default. Only an explicit "0" disables it.
  if (pugi::xml_attribute hrc = root.attribute("headerRowCount"); hrc) {
    if (parse_xml_u32_attr(hrc, 1U) == 0U) {
      table.header_row = false;
    }
  }
  // totalsRowCount defaults to 0 (no totals row).
  if (pugi::xml_attribute trc = root.attribute("totalsRowCount"); trc) {
    if (parse_xml_u32_attr(trc, 0U) >= 1U) {
      table.totals_row = true;
    }
  }

  if (pugi::xml_node cols_node = root.child("tableColumns"); cols_node) {
    for (pugi::xml_node col = cols_node.child("tableColumn"); col; col = col.next_sibling("tableColumn")) {
      TableColumn entry;
      entry.id = parse_xml_u32_attr(col.attribute("id"), 0U);
      if (pugi::xml_attribute n = col.attribute("name"); n) {
        entry.name = n.value();
      }
      if (pugi::xml_attribute lbl = col.attribute("totalsRowLabel"); lbl) {
        entry.totals_label = lbl.value();
      }
      if (pugi::xml_attribute fn = col.attribute("totalsRowFunction"); fn) {
        entry.totals_function = fn.value();
      }
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

  return table;
}

}  // namespace io
}  // namespace formulon
