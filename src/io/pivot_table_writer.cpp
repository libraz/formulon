// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the pivot-table-definition writer. See
// pivot_table_writer.h for the public contract; see
// pivot_table_reader.cpp for the symmetric grammar definition.

#include "io/pivot_table_writer.h"

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "io/ooxml_writer_cell.h"  // EncodeA1
#include "io/xml_escape.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"

namespace formulon::io {
namespace {

constexpr std::string_view kXmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
constexpr std::string_view kPivotNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/// Emits an A1 range string `"<topLeft>:<bottomRight>"` for the pivot's
/// rectangular anchor. The reader accepts both single-cell and range
/// forms; the writer always emits the range form for simplicity (and so
/// the bytes mirror what real Excel files contain).
std::string EncodeA1Range(std::uint32_t row, std::uint32_t col, std::uint32_t span_rows, std::uint32_t span_cols) {
  // Guard against zero spans on a default-constructed table: an empty
  // pivot anchor (0,0,0,0) becomes "A1:A1" rather than emitting an
  // invalid one-past-the-end address. The reader treats single-cell
  // refs as a 1x1 anchor, which matches the empty-table semantics on
  // the round trip.
  const std::uint32_t bot = span_rows == 0U ? row : row + span_rows - 1U;
  const std::uint32_t right = span_cols == 0U ? col : col + span_cols - 1U;
  std::string out = EncodeA1(row, col);
  out.push_back(':');
  out.append(EncodeA1(bot, right));
  return out;
}

/// Appends ` name="value"` for a present optional `<location>` offset
/// attribute, or nothing when the optional is empty. The values are
/// non-negative integers, so no XML escaping is required.
void AppendOptionalLocationAttr(std::string& out, std::string_view name, std::optional<std::uint32_t> value) {
  if (!value.has_value()) {
    return;
  }
  out.push_back(' ');
  out.append(name);
  out.append("=\"");
  out.append(std::to_string(*value));
  out.push_back('"');
}

/// Maps a `PivotAxis` to the OOXML `axis="..."` attribute body for the
/// non-Value cases. `Value` is intentionally absent: Value-axis fields
/// emit `dataField="1"` instead, matching real Excel output and what
/// the reader's "no axis attribute + dataField=1 -> Value" branch
/// expects on the round trip.
std::string_view AxisAttrName(pivot::PivotAxis axis) {
  switch (axis) {
    case pivot::PivotAxis::Row:
      return "axisRow";
    case pivot::PivotAxis::Col:
      return "axisCol";
    case pivot::PivotAxis::Page:
      return "axisPage";
    case pivot::PivotAxis::Value:
      // Caller handles this branch via `dataField="1"`; returning an
      // empty view here means "do not emit an axis attribute".
      return {};
  }
  return {};
}

/// Maps a `ShowValuesAs` to the OOXML `showDataAs="..."` attribute body.
/// `Normal` returns an empty view; the caller skips emission in that
/// case so the default attribute stays absent. Mirrors the reader's
/// `ParseShowDataAs` table for round-trip parity.
std::string_view ShowDataAsAttrName(pivot::ShowValuesAs s) {
  switch (s) {
    case pivot::ShowValuesAs::Normal:
      return {};
    case pivot::ShowValuesAs::PercentOfRow:
      return "percentOfRow";
    case pivot::ShowValuesAs::PercentOfCol:
      return "percentOfCol";
    case pivot::ShowValuesAs::PercentOfTotal:
      return "percentOfTotal";
    case pivot::ShowValuesAs::RunningTotalInRow:
      return "runTotal";
    case pivot::ShowValuesAs::RunningTotalInCol:
      return "runTotalInCol";
    case pivot::ShowValuesAs::Index:
      return "index";
    case pivot::ShowValuesAs::DifferenceFrom:
      return "difference";
    case pivot::ShowValuesAs::PercentDifferenceFrom:
      return "percentDiff";
    case pivot::ShowValuesAs::PercentOfParentRow:
      return "percentOfParentRow";
    case pivot::ShowValuesAs::PercentOfParentCol:
      return "percentOfParentCol";
    case pivot::ShowValuesAs::PercentOfParent:
      return "percentOfParent";
  }
  return {};
}

/// Maps an `Aggregation` to the OOXML `subtotal="..."` attribute body.
/// Mirrors the reader's `ParseAggregation` table exactly so the round
/// trip is bit-stable.
std::string_view AggregationAttrName(pivot::Aggregation a) {
  switch (a) {
    case pivot::Aggregation::Sum:
      return "sum";
    case pivot::Aggregation::Count:
      return "count";
    case pivot::Aggregation::Average:
      return "average";
    case pivot::Aggregation::Max:
      return "max";
    case pivot::Aggregation::Min:
      return "min";
    case pivot::Aggregation::Product:
      return "product";
    case pivot::Aggregation::CountNumbers:
      return "countNums";
    case pivot::Aggregation::StdDev:
      return "stdDev";
    case pivot::Aggregation::StdDevP:
      return "stdDevp";
    case pivot::Aggregation::Var:
      return "var";
    case pivot::Aggregation::VarP:
      return "varp";
  }
  return "sum";
}

/// Maps a `SubtotalFn` to the OOXML `<pivotField>` `*Subtotal` boolean
/// attribute name. Mirrors the reader's `kSubtotalAttrs` table so a
/// custom subtotal selection round-trips bit-stably. `countA` /
/// `countNums` use the spec's distinct attribute names.
std::string_view SubtotalAttrName(pivot::SubtotalFn fn) {
  switch (fn) {
    case pivot::SubtotalFn::Sum:
      return "sumSubtotal";
    case pivot::SubtotalFn::Count:
      return "countASubtotal";
    case pivot::SubtotalFn::Average:
      return "avgSubtotal";
    case pivot::SubtotalFn::Max:
      return "maxSubtotal";
    case pivot::SubtotalFn::Min:
      return "minSubtotal";
    case pivot::SubtotalFn::Product:
      return "productSubtotal";
    case pivot::SubtotalFn::CountNumbers:
      return "countSubtotal";
    case pivot::SubtotalFn::StdDev:
      return "stdDevSubtotal";
    case pivot::SubtotalFn::StdDevP:
      return "stdDevPSubtotal";
    case pivot::SubtotalFn::Var:
      return "varSubtotal";
    case pivot::SubtotalFn::VarP:
      return "varPSubtotal";
  }
  return {};
}

/// Emits one `<pivotField>` element. Self-closing when there are no
/// items; open/close pair otherwise.
void AppendPivotField(std::string& out, const pivot::PivotField& field) {
  out.append("<pivotField");
  if (field.axis == pivot::PivotAxis::Value) {
    // Excel-saved files use `dataField="1"` for Value-axis fields; the
    // reader accepts both that and `axis="axisValues"`, but emitting
    // `dataField="1"` keeps the integration-test fixture grammar.
    out.append(" dataField=\"1\"");
  } else {
    out.append(" axis=\"");
    out.append(AxisAttrName(field.axis));
    out.append("\"");
  }
  if (!field.custom_name.empty()) {
    out.append(" name=\"");
    AppendXmlEscaped(out, field.custom_name);
    out.append("\"");
  }
  if (field.subtotal_top) {
    // Default in OOXML is "0"; only emit when true to keep the output
    // compact. The reader treats missing `subtotalTop` as false.
    out.append(" subtotalTop=\"1\"");
  }
  // `defaultSubtotal` defaults to true; only emit it when turned OFF so
  // an explicit suppression survives the round trip. The custom
  // `*Subtotal` attributes are emitted only for the functions actually
  // selected, matching the writer's "preserve only non-default" rule.
  if (!field.default_subtotal) {
    out.append(" defaultSubtotal=\"0\"");
  }
  for (const pivot::SubtotalFn fn : field.subtotal_fns) {
    const std::string_view attr = SubtotalAttrName(fn);
    if (attr.empty()) {
      continue;
    }
    out.push_back(' ');
    out.append(attr);
    out.append("=\"1\"");
  }
  if (field.items.empty()) {
    out.append("/>");
    return;
  }
  out.append("><items count=\"");
  out.append(std::to_string(field.items.size()));
  out.append("\">");
  for (std::size_t i = 0; i < field.items.size(); ++i) {
    const pivot::PivotItem& item = field.items[i];
    // The item's `x` attribute is the document-order index of this
    // item in `field.items`, which mirrors the cache `shared_items`
    // ordering on the reader's round trip. The reader does not store
    // item names today (they are resolved against the cache at eval
    // time), so we never emit a name attribute here.
    out.append("<item x=\"");
    out.append(std::to_string(i));
    out.append("\"");
    if (!item.visible) {
      out.append(" h=\"1\"");
    }
    out.append("/>");
  }
  out.append("</items></pivotField>");
}

/// Emits `<rowFields>` or `<colFields>` block. Only called when
/// `order` is non-empty.
void AppendFieldOrder(std::string& out, std::string_view tag, const std::vector<std::uint32_t>& order) {
  out.push_back('<');
  out.append(tag);
  out.append(" count=\"");
  out.append(std::to_string(order.size()));
  out.append("\">");
  for (const std::uint32_t idx : order) {
    out.append("<field x=\"");
    out.append(std::to_string(idx));
    out.append("\"/>");
  }
  out.append("</");
  out.append(tag);
  out.push_back('>');
}

/// Emits the `<dataFields>` block. Only called when
/// `table.data_fields()` is non-empty.
void AppendDataFields(std::string& out, const std::vector<pivot::PivotDataField>& data_fields) {
  out.append("<dataFields count=\"");
  out.append(std::to_string(data_fields.size()));
  out.append("\">");
  for (const pivot::PivotDataField& df : data_fields) {
    out.append("<dataField name=\"");
    AppendXmlEscaped(out, df.name);
    out.append("\" fld=\"");
    out.append(std::to_string(df.field_index));
    // Always emit the subtotal attribute, even for the implicit Sum
    // default. Real Excel files emit it explicitly and tests are
    // clearer when the round trip preserves the spelling.
    out.append("\" subtotal=\"");
    out.append(AggregationAttrName(df.aggregation));
    out.append("\"");
    if (!df.number_format.empty()) {
      // Pass through verbatim; the reader stored whatever string was
      // in the source attribute (typically a numFmtId integer in
      // string form, but we do not enforce that here).
      out.append(" numFmtId=\"");
      AppendXmlEscaped(out, df.number_format);
      out.append("\"");
    }
    if (df.show_as != pivot::ShowValuesAs::Normal) {
      out.append(" showDataAs=\"");
      out.append(ShowDataAsAttrName(df.show_as));
      out.append("\"");
    }
    if (df.show_as_base_field.has_value()) {
      out.append(" baseField=\"");
      out.append(std::to_string(*df.show_as_base_field));
      out.append("\"");
    }
    if (df.show_as_base_item.has_value()) {
      out.append(" baseItem=\"");
      out.append(std::to_string(*df.show_as_base_item));
      out.append("\"");
    }
    out.append("/>");
  }
  out.append("</dataFields>");
}

}  // namespace

std::string write_pivot_table_definition(const pivot::PivotTable& table) {
  std::string out;
  // Pre-reserve a rough lower bound. The XML declaration + root + location
  // run ~200B; per-field cost is ~80B with a few items each, per data-
  // field cost ~80B. This trims the first couple of growth-pass
  // reallocations without bloating tiny tables.
  std::size_t items_total = 0;
  for (const pivot::PivotField& f : table.fields()) {
    items_total += f.items.size();
  }
  out.reserve(256 + table.fields().size() * 80 + items_total * 16 + table.data_fields().size() * 80);

  out.append(kXmlDecl);
  out.append("<pivotTableDefinition xmlns=\"");
  out.append(kPivotNs);
  out.append("\" name=\"");
  AppendXmlEscaped(out, table.name());
  out.append("\" cacheId=\"");
  out.append(std::to_string(table.pivot_cache_id()));
  out.append("\"");
  // Grand-total flags default to true in OOXML and in the model. Emit them
  // only when turned OFF so an explicit OFF state survives the round trip
  // (omitting them would let the reader's default flip the pivot back ON).
  if (!table.grand_totals_rows()) {
    out.append(" rowGrandTotals=\"0\"");
  }
  if (!table.grand_totals_cols()) {
    out.append(" colGrandTotals=\"0\"");
  }
  out.append(">");

  out.append("<location ref=\"");
  // Anchor is unconditionally emitted as a range; an empty table
  // (span_rows == span_cols == 0) round-trips through the single-cell
  // form via EncodeA1Range's zero-span guard.
  AppendXmlEscaped(out, EncodeA1Range(table.anchor_row(), table.anchor_col(), table.span_rows(), table.span_cols()));
  out.append("\"");
  // Re-emit the `<location>` offset attributes captured at read time.
  // ECMA-376 requires firstHeaderRow / firstDataRow / firstDataCol;
  // rowPageCount / colPageCount are optional. Each is emitted only when it
  // was present in the source so a schema-valid `<location>` round-trips
  // without inventing values for absent optionals.
  AppendOptionalLocationAttr(out, "firstHeaderRow", table.location_first_header_row());
  AppendOptionalLocationAttr(out, "firstDataRow", table.location_first_data_row());
  AppendOptionalLocationAttr(out, "firstDataCol", table.location_first_data_col());
  AppendOptionalLocationAttr(out, "rowPageCount", table.location_row_page_count());
  AppendOptionalLocationAttr(out, "colPageCount", table.location_col_page_count());
  out.append("/>");

  // `<pivotFields>` is always emitted (even for an empty count) so the
  // structure stays self-describing and round-trips through the reader's
  // optional-block scan unambiguously.
  out.append("<pivotFields count=\"");
  out.append(std::to_string(table.fields().size()));
  out.append("\">");
  for (const pivot::PivotField& f : table.fields()) {
    AppendPivotField(out, f);
  }
  out.append("</pivotFields>");

  if (!table.row_field_order().empty()) {
    AppendFieldOrder(out, "rowFields", table.row_field_order());
  }
  if (!table.col_field_order().empty()) {
    AppendFieldOrder(out, "colFields", table.col_field_order());
  }
  if (!table.data_fields().empty()) {
    AppendDataFields(out, table.data_fields());
  }

  // Re-emit any unmodelled child elements the reader captured during the
  // last load (`<pageFields>`, `<formats>`, `<calculatedFields>`,
  // `<calculatedItems>`, `<pivotTableStyleInfo>`, `<extLst>`, ...). The
  // bytes are owned by the table and re-appended verbatim so a
  // read -> write round trip preserves features v1.0 does not model
  // structurally. Tables built from scratch leave the buffer empty.
  if (!table.raw_passthrough_xml().empty()) {
    out.append(table.raw_passthrough_xml());
  }

  out.append("</pivotTableDefinition>");
  return out;
}

}  // namespace formulon::io
