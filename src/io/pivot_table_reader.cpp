// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the pivot-table-definition reader. See
// pivot_table_reader.h for the public contract.

#include "io/pivot_table_reader.h"

#include <cerrno>
#include <cstdint>
#include <cstdlib>
#include <cstring>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/cell_parser.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon::io {
namespace {

/// Parses a non-negative decimal integer attribute body, locale-independent.
/// Returns `default_value` on missing / empty / malformed / negative input;
/// caps at `std::uint32_t` range. Mirrors the defensive parsing in
/// `tables_reader.cpp` so a stray attribute does not reject an otherwise-
/// valid pivot table.
std::uint32_t ParseU32Attr(const pugi::xml_attribute& attr, std::uint32_t default_value) {
  if (!attr) {
    return default_value;
  }
  const char* raw = attr.value();
  if (raw == nullptr || *raw == '\0') {
    return default_value;
  }
  errno = 0;
  char* end = nullptr;
  const unsigned long parsed = std::strtoul(raw, &end, 10);
  if (end == raw || *end != '\0' || errno != 0) {
    return default_value;
  }
  if (parsed > 0xFFFFFFFFUL) {
    return default_value;
  }
  return static_cast<std::uint32_t>(parsed);
}

/// Maps an OOXML `axis="..."` attribute body to a `PivotAxis`. Unknown /
/// missing values fold to `Value` — both `axisValues` and the absent-
/// axis case mean "this field is exclusively used as a data field" in
/// the OOXML spec.
pivot::PivotAxis ParseAxis(std::string_view text) {
  if (text == "axisRow") {
    return pivot::PivotAxis::Row;
  }
  if (text == "axisCol") {
    return pivot::PivotAxis::Col;
  }
  if (text == "axisPage") {
    return pivot::PivotAxis::Page;
  }
  return pivot::PivotAxis::Value;
}

/// Maps an OOXML `subtotal="..."` attribute body on a `<dataField>` to an
/// `Aggregation`. Unknown spellings fold to `Sum` so the reader stays
/// forward-compatible with Excel additions; the writer round-trips the
/// original attribute through a separate passthrough path (future PR).
pivot::Aggregation ParseAggregation(std::string_view text) {
  if (text == "count") {
    return pivot::Aggregation::Count;
  }
  if (text == "average") {
    return pivot::Aggregation::Average;
  }
  if (text == "max") {
    return pivot::Aggregation::Max;
  }
  if (text == "min") {
    return pivot::Aggregation::Min;
  }
  if (text == "product") {
    return pivot::Aggregation::Product;
  }
  if (text == "countNums") {
    return pivot::Aggregation::CountNumbers;
  }
  if (text == "stdDev") {
    return pivot::Aggregation::StdDev;
  }
  if (text == "stdDevp") {
    return pivot::Aggregation::StdDevP;
  }
  if (text == "var") {
    return pivot::Aggregation::Var;
  }
  if (text == "varp") {
    return pivot::Aggregation::VarP;
  }
  // "sum" and anything unrecognised.
  return pivot::Aggregation::Sum;
}

/// Maps an OOXML `showDataAs="..."` attribute body to a `ShowValuesAs`.
/// Unknown spellings fold to `Normal` so the reader stays
/// forward-compatible with Excel additions; the writer's inverse helper
/// produces the same set of attribute names. The mapping for
/// `RunningTotalInCol` uses the Excel-private `runTotalInCol` spelling
/// since the standard OOXML schema only specifies a row-direction
/// `runTotal`; this is documented as a known scoping in the round-trip
/// table on the `ShowValuesAs` enum.
pivot::ShowValuesAs ParseShowDataAs(std::string_view text) {
  if (text == "percentOfRow") {
    return pivot::ShowValuesAs::PercentOfRow;
  }
  if (text == "percentOfCol") {
    return pivot::ShowValuesAs::PercentOfCol;
  }
  if (text == "percentOfTotal") {
    return pivot::ShowValuesAs::PercentOfTotal;
  }
  if (text == "runTotal") {
    return pivot::ShowValuesAs::RunningTotalInRow;
  }
  if (text == "runTotalInCol") {
    return pivot::ShowValuesAs::RunningTotalInCol;
  }
  if (text == "index") {
    return pivot::ShowValuesAs::Index;
  }
  if (text == "difference") {
    return pivot::ShowValuesAs::DifferenceFrom;
  }
  if (text == "percentDiff") {
    return pivot::ShowValuesAs::PercentDifferenceFrom;
  }
  if (text == "percentOfParentRow") {
    return pivot::ShowValuesAs::PercentOfParentRow;
  }
  if (text == "percentOfParentCol") {
    return pivot::ShowValuesAs::PercentOfParentCol;
  }
  if (text == "percentOfParent") {
    return pivot::ShowValuesAs::PercentOfParent;
  }
  return pivot::ShowValuesAs::Normal;
}

/// Returns true iff the OOXML boolean attribute body is the literal `"1"`
/// or `"true"`. Anything else (including missing) reads as false.
bool ParseBoolAttr(const pugi::xml_attribute& attr) {
  if (!attr) {
    return false;
  }
  const std::string_view text = attr.value();
  return text == "1" || text == "true";
}

/// Decodes one `<location ref="A3:D10"/>` (or single-cell `ref="A3"`)
/// into 0-based anchor + spans. Defers cell-coordinate decoding to the
/// shared `parse_a1` helper so the conventions match cell parsing
/// elsewhere.
Expected<void, Error> DecodeLocationRef(std::string_view ref, pivot::PivotTable* out) {
  if (ref.empty()) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt,
                      "pivotTableDefinition: <location> missing required ref attribute", "context=pivot_table_reader");
  }
  const std::size_t colon = ref.find(':');
  if (colon == std::string_view::npos) {
    // Single-cell anchor (Excel writes this for an empty pivot).
    auto rc = parse_a1(ref);
    if (!rc) {
      std::string ctx("context=pivot_table_reader ref=");
      ctx.append(ref);
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "pivotTableDefinition: <location> ref unparseable",
                        std::move(ctx));
    }
    out->set_anchor(rc.value().first, rc.value().second, 1U, 1U);
    return Expected<void, Error>::Ok();
  }
  const std::string_view a = ref.substr(0, colon);
  const std::string_view b = ref.substr(colon + 1);
  auto a_rc = parse_a1(a);
  auto b_rc = parse_a1(b);
  if (!a_rc || !b_rc) {
    std::string ctx("context=pivot_table_reader ref=");
    ctx.append(ref);
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "pivotTableDefinition: <location> ref unparseable",
                      std::move(ctx));
  }
  const std::uint32_t r0 = a_rc.value().first;
  const std::uint32_t c0 = a_rc.value().second;
  const std::uint32_t r1 = b_rc.value().first;
  const std::uint32_t c1 = b_rc.value().second;
  const std::uint32_t row_top = (r0 < r1) ? r0 : r1;
  const std::uint32_t row_bot = (r0 < r1) ? r1 : r0;
  const std::uint32_t col_left = (c0 < c1) ? c0 : c1;
  const std::uint32_t col_right = (c0 < c1) ? c1 : c0;
  out->set_anchor(row_top, col_left, row_bot - row_top + 1U, col_right - col_left + 1U);
  return Expected<void, Error>::Ok();
}

/// Walks the `<items>` child of a `<pivotField>` and populates
/// `field.items`. `<item t="default">` and `<item t="grand">` are
/// subtotal / grand-total markers in the OOXML schema, NOT real items;
/// they are skipped here. For real items we capture visibility from
/// `h="1"` (hidden -> `visible = false`) and leave `name` empty: the
/// `x` attribute references the cache's `shared_items` by index, which
/// the evaluator resolves later against the bound `PivotCache`.
void ParseItems(const pugi::xml_node& items_node, pivot::PivotField* field) {
  for (pugi::xml_node it = items_node.child("item"); it; it = it.next_sibling("item")) {
    const std::string_view t = it.attribute("t").as_string();
    if (t == "default" || t == "grand" || t == "blank" || t == "sum" || t == "count" || t == "avg" || t == "max" ||
        t == "min" || t == "product" || t == "countA" || t == "stdDev" || t == "stdDevP" || t == "var" || t == "varP") {
      // Subtotal / grand-total marker; not a real item.
      continue;
    }
    pivot::PivotItem entry;
    entry.visible = !ParseBoolAttr(it.attribute("h"));
    field->items.push_back(std::move(entry));
  }
}

/// Walks `<pivotFields>` in document order, materialising each
/// `<pivotField>` into the table.
void ParsePivotFields(const pugi::xml_node& fields_node, pivot::PivotTable* out) {
  for (pugi::xml_node f = fields_node.child("pivotField"); f; f = f.next_sibling("pivotField")) {
    pivot::PivotField field;
    field.axis = ParseAxis(f.attribute("axis").as_string());
    if (pugi::xml_attribute name_attr = f.attribute("name"); name_attr) {
      field.custom_name = name_attr.value();
    }
    field.subtotal_top = ParseBoolAttr(f.attribute("subtotalTop"));
    if (pugi::xml_node items_node = f.child("items"); items_node) {
      ParseItems(items_node, &field);
    }
    out->mutable_fields().push_back(std::move(field));
  }
}

/// Walks `<rowFields>` / `<colFields>` and pushes each `<field x="N">`
/// onto `out`. Missing `x` defaults to 0 (matches OOXML schema).
void ParseFieldOrder(const pugi::xml_node& parent, std::vector<std::uint32_t>* out) {
  for (pugi::xml_node f = parent.child("field"); f; f = f.next_sibling("field")) {
    out->push_back(ParseU32Attr(f.attribute("x"), 0U));
  }
}

/// Walks `<dataFields>` in document order. Returns `kIoSheetCorrupt` on
/// the first `<dataField>` that lacks a `name` attribute, since that
/// name is the GETPIVOTDATA lookup key.
Expected<void, Error> ParseDataFields(const pugi::xml_node& parent, pivot::PivotTable* out) {
  for (pugi::xml_node df = parent.child("dataField"); df; df = df.next_sibling("dataField")) {
    pivot::PivotDataField entry;
    pugi::xml_attribute name_attr = df.attribute("name");
    if (!name_attr || name_attr.value() == nullptr || *name_attr.value() == '\0') {
      return make_error(FormulonErrorCode::kIoSheetCorrupt,
                        "pivotTableDefinition: <dataField> missing required name attribute",
                        "context=pivot_table_reader");
    }
    entry.name = name_attr.value();
    entry.field_index = ParseU32Attr(df.attribute("fld"), 0U);
    entry.aggregation = ParseAggregation(df.attribute("subtotal").as_string());
    if (pugi::xml_attribute nf = df.attribute("numFmtId"); nf) {
      entry.number_format = nf.value();
    }
    if (pugi::xml_attribute sa = df.attribute("showDataAs"); sa) {
      entry.show_as = ParseShowDataAs(sa.as_string());
    }
    if (pugi::xml_attribute bf = df.attribute("baseField"); bf) {
      entry.show_as_base_field = ParseU32Attr(bf, 0U);
    }
    if (pugi::xml_attribute bi = df.attribute("baseItem"); bi) {
      entry.show_as_base_item = ParseU32Attr(bi, 0U);
    }
    out->mutable_data_fields().push_back(std::move(entry));
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<pivot::PivotTable, Error> read_pivot_table_definition(const std::vector<std::uint8_t>& definition_bytes) {
  pugi::xml_document doc;
  pugi::xml_parse_result parse =
      doc.load_buffer(definition_bytes.data(), definition_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    std::string ctx("context=pivot_table_reader desc=");
    ctx.append(parse.description());
    return make_error(FormulonErrorCode::kIoXmlParse, "pivotTable*.xml: pugixml parse failed", std::move(ctx));
  }
  pugi::xml_node root = doc.child("pivotTableDefinition");
  if (!root) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid, "pivotTable*.xml: missing <pivotTableDefinition> root",
                      "context=pivot_table_reader");
  }

  pivot::PivotTable table;
  if (pugi::xml_attribute name_attr = root.attribute("name"); name_attr) {
    table.set_name(name_attr.value());
  }
  table.set_pivot_cache_id(ParseU32Attr(root.attribute("cacheId"), 0U));

  pugi::xml_node loc = root.child("location");
  if (!loc) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "pivotTable*.xml: missing <location> element",
                      "context=pivot_table_reader");
  }
  pugi::xml_attribute ref_attr = loc.attribute("ref");
  if (!ref_attr) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "pivotTable*.xml: <location> missing required ref attribute",
                      "context=pivot_table_reader");
  }
  if (auto status = DecodeLocationRef(ref_attr.value(), &table); !status) {
    return status.error();
  }

  if (pugi::xml_node fields = root.child("pivotFields"); fields) {
    ParsePivotFields(fields, &table);
  }
  if (pugi::xml_node rows = root.child("rowFields"); rows) {
    ParseFieldOrder(rows, &table.mutable_row_field_order());
  }
  if (pugi::xml_node cols = root.child("colFields"); cols) {
    ParseFieldOrder(cols, &table.mutable_col_field_order());
  }
  if (pugi::xml_node data = root.child("dataFields"); data) {
    if (auto status = ParseDataFields(data, &table); !status) {
      return status.error();
    }
  }
  // Capture any remaining direct children of <pivotTableDefinition> that
  // we do not model structurally (`<pageFields>`, `<formats>`,
  // `<conditionalFormats>`, `<chartFormats>`, `<calculatedFields>`,
  // `<calculatedItems>`, `<pivotTableStyleInfo>`, `<extLst>`, ...) into
  // a single raw-XML buffer. The writer re-emits the buffer verbatim
  // between the structured tail and `</pivotTableDefinition>`, so
  // unmodelled features survive a read -> write round trip even when
  // v1.0 cannot evaluate them.
  static const std::string_view kRecognized[] = {"location",   "pivotFields", "rowFields", "colFields",
                                                 "dataFields", "rowItems",    "colItems"};
  struct StringXmlWriter : pugi::xml_writer {
    std::string* dst;
    void write(const void* data, std::size_t size) override { dst->append(static_cast<const char*>(data), size); }
  };
  StringXmlWriter sink{};
  sink.dst = &table.mutable_raw_passthrough_xml();
  for (pugi::xml_node child = root.first_child(); child; child = child.next_sibling()) {
    if (child.type() != pugi::node_element) {
      continue;
    }
    bool recognised = false;
    for (std::string_view name : kRecognized) {
      if (name == child.name()) {
        recognised = true;
        break;
      }
    }
    if (recognised) {
      continue;
    }
    child.print(sink, /*indent=*/"", pugi::format_raw);
  }

  return table;
}

}  // namespace formulon::io
