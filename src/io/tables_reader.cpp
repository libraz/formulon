// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the tables-part reader. See tables_reader.h for the
// public contract.

#include "io/tables_reader.h"

#include <cerrno>
#include <climits>
#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <string>
#include <utility>
#include <vector>

#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace {

/// Parses an unsigned-int attribute, returning `default_value` on any
/// malformed / negative / out-of-range input. Defensive parsing keeps a
/// stray attribute from rejecting an otherwise-valid table.
///
/// Uses `strtoul` with explicit `errno` and `> UINT32_MAX` checks: a
/// 32-bit `long` would clamp legitimate IDs (e.g. tableId="3000000000")
/// to LONG_MAX without surfacing the overflow.
std::uint32_t ParseU32Attr(const pugi::xml_attribute& attr, std::uint32_t default_value) {
  if (!attr) {
    return default_value;
  }
  const char* raw = attr.value();
  if (raw == nullptr || *raw == '\0') {
    return default_value;
  }
  // Reject leading sign characters explicitly: strtoul accepts a leading
  // '-' and silently wraps the result, so "-1" would otherwise parse as
  // UINT32_MAX rather than being rejected.
  if (*raw == '-' || *raw == '+') {
    return default_value;
  }
  errno = 0;
  char* end = nullptr;
  const unsigned long parsed = std::strtoul(raw, &end, 10);
  if (end == raw || *end != '\0' || errno != 0 || parsed > static_cast<unsigned long>(UINT32_MAX)) {
    return default_value;
  }
  return static_cast<std::uint32_t>(parsed);
}

}  // namespace

Expected<TableMetadata, Error> read_table(const std::vector<std::uint8_t>& table_bytes, std::size_t sheet_index) {
  pugi::xml_document doc;
  pugi::xml_parse_result parse =
      doc.load_buffer(table_bytes.data(), table_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    std::string ctx("context=tables_reader desc=");
    ctx.append(parse.description());
    return make_error(FormulonErrorCode::kIoXmlParse, "table*.xml: pugixml parse failed", std::move(ctx));
  }
  pugi::xml_node root = doc.child("table");
  if (!root) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid, "table*.xml: missing <table> root",
                      "context=tables_reader");
  }

  TableMetadata table;
  table.sheet_index = sheet_index;
  table.id = ParseU32Attr(root.attribute("id"), 0U);
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
    if (ParseU32Attr(hrc, 1U) == 0U) {
      table.header_row = false;
    }
  }
  // totalsRowCount defaults to 0 (no totals row).
  if (pugi::xml_attribute trc = root.attribute("totalsRowCount"); trc) {
    if (ParseU32Attr(trc, 0U) >= 1U) {
      table.totals_row = true;
    }
  }

  if (pugi::xml_node cols_node = root.child("tableColumns"); cols_node) {
    for (pugi::xml_node col = cols_node.child("tableColumn"); col; col = col.next_sibling("tableColumn")) {
      TableColumn entry;
      entry.id = ParseU32Attr(col.attribute("id"), 0U);
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
