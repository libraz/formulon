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
namespace {

/// Serialises a pugixml node to a raw (unindented) XML string. Local to
/// this TU; used to capture `<tableStyleInfo>` verbatim.
struct StringXmlWriter final : pugi::xml_writer {
  std::string* dst = nullptr;
  void write(const void* data, std::size_t size) override {
    if (dst != nullptr) {
      dst->append(static_cast<const char*>(data), size);
    }
  }
};

std::string RawXml(const pugi::xml_node& node) {
  std::string out;
  StringXmlWriter sink;
  sink.dst = &out;
  node.print(sink, /*indent=*/"", pugi::format_raw);
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
    table.table_style_info_xml = RawXml(style_info);
  }

  return table;
}

}  // namespace io
}  // namespace formulon
