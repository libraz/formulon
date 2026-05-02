// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the pivot-table-definition writer. See
// pivot_table_writer.h for the public contract; see
// pivot_table_reader.cpp for the symmetric grammar definition.

#include "io/pivot_table_writer.h"

#include <cstddef>
#include <cstdint>
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
  out.append("\">");

  out.append("<location ref=\"");
  // Anchor is unconditionally emitted as a range; an empty table
  // (span_rows == span_cols == 0) round-trips through the single-cell
  // form via EncodeA1Range's zero-span guard.
  AppendXmlEscaped(out, EncodeA1Range(table.anchor_row(), table.anchor_col(), table.span_rows(), table.span_cols()));
  out.append("\"/>");

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

  // `<pageFields>`, `<formats>`, `<conditionalFormats>`, `<chartFormats>`,
  // `<pivotTableStyleInfo>`, and `<extLst>` are intentionally not emitted.
  // The reader silently skips them today; a future-passthrough writer
  // PR will preserve them as bytes.

  out.append("</pivotTableDefinition>");
  return out;
}

}  // namespace formulon::io
