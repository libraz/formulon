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

#include "io/pivot_attr_names.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "utils/a1_ref.h"

namespace formulon::io {
namespace {

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
  std::string out = a1::encode_a1(row, col);
  out.push_back(':');
  out.append(a1::encode_a1(bot, right));
  return out;
}

/// Appends ` name="value"` for a present optional `<location>` offset
/// attribute, or nothing when the optional is empty. The values are
/// non-negative integers, so no XML escaping is required.
void AppendOptionalLocationAttr(std::string& out, std::string_view name, std::optional<std::uint32_t> value) {
  if (!value.has_value()) {
    return;
  }
  append_xml_attr_uint(out, name, *value);
}

/// Emits one `<pivotField>` element. Self-closing when there are no
/// items; open/close pair otherwise.
void AppendPivotField(std::string& out, const pivot::PivotField& field, pivot::PivotLayout layout) {
  out.append("<pivotField");
  if (field.axis == pivot::PivotAxis::Value) {
    // Excel-saved files use `dataField="1"` for Value-axis fields; the
    // reader accepts both that and `axis="axisValues"`, but emitting
    // `dataField="1"` keeps the integration-test fixture grammar.
    out.append(" dataField=\"1\"");
  } else if (field.axis == pivot::PivotAxis::None) {
    // Unused ("available") field: emit neither an axis nor dataField.
  } else {
    append_xml_attr(out, "axis", enum_name(kPivotAxisNames, field.axis));
  }
  if (!field.custom_name.empty()) {
    append_xml_attr(out, "name", field.custom_name);
  }
  if (!field.number_format.empty()) {
    append_xml_attr(out, "numFmtId", field.number_format);
  }
  // `subtotalTop` defaults to true in OOXML; only emit it when turned OFF
  // so an explicit "subtotals at bottom" choice survives the round trip.
  if (!field.subtotal_top) {
    out.append(" subtotalTop=\"0\"");
  }
  // `defaultSubtotal` defaults to true; only emit it when turned OFF so
  // an explicit suppression survives the round trip. The custom
  // `*Subtotal` attributes are emitted only for the functions actually
  // selected, matching the writer's "preserve only non-default" rule.
  if (!field.default_subtotal) {
    out.append(" defaultSubtotal=\"0\"");
  }
  for (const pivot::SubtotalFn fn : field.subtotal_fns) {
    const PivotSubtotalName* names = find_pivot_subtotal_name(fn);
    if (names == nullptr) {
      continue;
    }
    out.push_back(' ');
    out.append(names->attr);
    out.append("=\"1\"");
  }
  // `sortType` defaults to ascending-by-label; only emit it when the
  // field's sort deviates from that (explicit descending, or manual
  // order from `<items>` document order), matching the reader's default.
  if (field.sort.manual) {
    out.append(" sortType=\"manual\"");
  } else if (!field.sort.ascending) {
    out.append(" sortType=\"descending\"");
  }
  // The report form is per field in Excel (`compact` / `outline` default to
  // true); every field carries the table's form, as Excel writes it.
  if (layout == pivot::PivotLayout::Tabular) {
    out.append(" compact=\"0\" outline=\"0\"");
  } else if (layout == pivot::PivotLayout::Outline) {
    out.append(" compact=\"0\"");
  }
  // Re-emit any unmodelled `<pivotField>` attributes captured on read.
  append_raw_attrs(out, field.passthrough_attrs);
  if (field.items.empty()) {
    out.append("/>");
    return;
  }
  // Excel closes a field's `<items>` with one entry per subtotal it shows:
  // `<item t="default"/>` for the implicit one, or a token per explicitly
  // selected function. The reader drops these on the way in because they
  // are not real items -- the selection lives in `default_subtotal` /
  // `subtotal_fns` -- so the writer has to put them back. Excel treats a
  // field whose items lack its subtotal entry as damaged and offers to
  // repair the file, and every part of that file is otherwise valid, so
  // nothing short of Excel's own verdict reports it.
  std::size_t subtotal_markers = field.default_subtotal ? 1U : 0U;
  for (const pivot::SubtotalFn fn : field.subtotal_fns) {
    if (find_pivot_subtotal_name(fn) != nullptr) {
      ++subtotal_markers;
    }
  }
  out.append("><items count=\"");
  out.append(std::to_string(field.items.size() + subtotal_markers));
  out.append("\">");
  for (std::size_t i = 0; i < field.items.size(); ++i) {
    const pivot::PivotItem& item = field.items[i];
    // Re-emit the cache index the reader captured from `<item x="N">` so
    // the item still points at the same `shared_items` entry after a
    // round trip. When the source omitted `x` (or the item was built
    // from scratch), fall back to the document-order position, matching
    // Excel's implicit ordering. Item names are resolved against the
    // cache, not stored on the element, so no name attribute is emitted.
    const std::uint32_t x = item.has_cache_index ? item.cache_index : static_cast<std::uint32_t>(i);
    out.append("<item x=\"");
    out.append(std::to_string(x));
    out.append("\"");
    if (!item.visible) {
      out.append(" h=\"1\"");
    }
    out.append("/>");
  }
  if (field.default_subtotal) {
    out.append("<item t=\"default\"/>");
  }
  for (const pivot::SubtotalFn fn : field.subtotal_fns) {
    const PivotSubtotalName* names = find_pivot_subtotal_name(fn);
    if (names == nullptr) {
      continue;
    }
    out.append("<item t=\"");
    out.append(names->item);
    out.append("\"/>");
  }
  out.append("</items></pivotField>");
}

/// Emits `<rowFields>` or `<colFields>` block. Only called when `order`
/// is non-empty or `values_position` is set.
///
/// `values_position`, when set, re-inserts `<field x="-2"/>` (the Values
/// pseudo-field marker `PivotTable::row_values_position()` /
/// `col_values_position()` records) at its original document position
/// among `order`'s real field indices.
void AppendFieldOrder(std::string& out, std::string_view tag, const std::vector<std::uint32_t>& order,
                      std::optional<std::size_t> values_position) {
  const std::size_t total = order.size() + (values_position.has_value() ? 1U : 0U);
  out.push_back('<');
  out.append(tag);
  out.append(" count=\"");
  out.append(std::to_string(total));
  out.append("\">");
  std::size_t real_idx = 0;
  for (std::size_t i = 0; i < total; ++i) {
    out.append("<field x=\"");
    if (values_position.has_value() && *values_position == i) {
      out.append(std::to_string(pivot::kValuesFieldPosition));
    } else {
      out.append(std::to_string(order[real_idx]));
      ++real_idx;
    }
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
    AppendXmlAttrEscaped(out, df.name);
    out.append("\" fld=\"");
    out.append(std::to_string(df.field_index));
    // Always emit the subtotal attribute, even for the implicit Sum
    // default. Real Excel files emit it explicitly and tests are
    // clearer when the round trip preserves the spelling.
    out.append("\" subtotal=\"");
    out.append(enum_name(kPivotAggregationNames, df.aggregation, "sum"));
    out.append("\"");
    if (!df.number_format.empty()) {
      // Pass through verbatim; the reader stored whatever string was
      // in the source attribute (typically a numFmtId integer in
      // string form, but we do not enforce that here).
      append_xml_attr(out, "numFmtId", df.number_format);
    }
    if (df.show_as != pivot::ShowValuesAs::Normal) {
      out.append(" showDataAs=\"");
      // OOXML has only `runTotal`; the direction is carried by `baseField`.
      const pivot::ShowValuesAs show_as =
          df.show_as == pivot::ShowValuesAs::RunningTotalInCol ? pivot::ShowValuesAs::RunningTotalInRow : df.show_as;
      out.append(enum_name(kPivotShowDataAsNames, show_as));
      out.append("\"");
    }
    if (df.show_as_base_field.has_value()) {
      append_xml_attr_uint(out, "baseField", *df.show_as_base_field);
    }
    if (df.show_as_base_item.has_value()) {
      append_xml_attr_uint(out, "baseItem", *df.show_as_base_item);
    }
    out.append("/>");
  }
  out.append("</dataFields>");
}

}  // namespace

std::string write_pivot_table_definition(const pivot::PivotTable& table,
                                         std::optional<PivotRenderedSpan> rendered_span) {
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
  AppendXmlAttrEscaped(out, table.name());
  out.append("\" cacheId=\"");
  out.append(std::to_string(table.pivot_cache_id()));
  out.append("\" dataCaption=\"");
  AppendXmlAttrEscaped(out, table.data_caption());
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
  // Report layout as Excel writes it at table level, where it is the form
  // new fields take (`compact` defaults to true, `outline` to false).
  switch (table.layout()) {
    case pivot::PivotLayout::Compact:
      out.append(" outline=\"1\" outlineData=\"1\"");
      break;
    case pivot::PivotLayout::Tabular:
      out.append(" compact=\"0\" compactData=\"0\"");
      break;
    case pivot::PivotLayout::Outline:
      out.append(" compact=\"0\" compactData=\"0\" outline=\"1\" outlineData=\"1\"");
      break;
  }
  // Re-emit any unmodelled root attributes captured on read.
  append_raw_attrs(out, table.passthrough_attrs());
  out.append(">");

  out.append("<location ref=\"");
  // Anchor is unconditionally emitted as a range; an empty table
  // (span_rows == span_cols == 0) round-trips through the single-cell
  // form via EncodeA1Range's zero-span guard.
  //
  // A span the reader decoded is Excel's own and is re-emitted verbatim.
  //
  // Otherwise the projection acts as a floor rather than a replacement.
  // Under-sizing is the failure that matters: a `ref` smaller than the grid
  // the pivot draws makes Excel terminate when the report is refreshed, and
  // that is what a model assembled in memory produces, because nothing
  // revises the span it was created with as fields are added. Over-sizing
  // is not a failure -- Excel reserves the range and recomputes it on
  // refresh -- so a caller that deliberately set a wider span through
  // `fm_workbook_pivot_set_anchor` keeps it.
  std::uint32_t span_rows = table.span_rows();
  std::uint32_t span_cols = table.span_cols();
  if (rendered_span.has_value() && !table.has_authored_span()) {
    span_rows = span_rows > rendered_span->rows ? span_rows : rendered_span->rows;
    span_cols = span_cols > rendered_span->cols ? span_cols : rendered_span->cols;
  }
  AppendXmlAttrEscaped(out, EncodeA1Range(table.anchor_row(), table.anchor_col(), span_rows, span_cols));
  out.append("\"");
  // ECMA-376 requires firstHeaderRow / firstDataRow / firstDataCol. Preserve
  // authored values, including explicit zero, while supplying the schema's
  // independent fallback for a model built without those attributes.
  AppendOptionalLocationAttr(out, "firstHeaderRow", table.location_first_header_row().value_or(1U));
  AppendOptionalLocationAttr(out, "firstDataRow", table.location_first_data_row().value_or(1U));
  AppendOptionalLocationAttr(out, "firstDataCol", table.location_first_data_col().value_or(1U));
  // rowPageCount / colPageCount are optional and stay presence-based.
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
    AppendPivotField(out, f, table.layout());
  }
  out.append("</pivotFields>");

  if (!table.row_field_order().empty() || table.row_values_position().has_value()) {
    AppendFieldOrder(out, "rowFields", table.row_field_order(), table.row_values_position());
  }
  // `<rowItems>` and any other post-rowFields unmodelled elements.
  out.append(table.raw_passthrough_after_row_fields());

  if (!table.col_field_order().empty() || table.col_values_position().has_value()) {
    AppendFieldOrder(out, "colFields", table.col_field_order(), table.col_values_position());
  }
  // `<colItems>` / `<pageFields>` and other post-colFields unmodelled
  // elements, which the schema places before `<dataFields>`.
  out.append(table.raw_passthrough_after_col_fields());

  // A table whose `<pageFields>` was decoded re-emits it from the bin
  // just flushed, so synthesise the element only for a table that never
  // carried one -- built through the C API, or read from a definition
  // that placed a field on the page axis without declaring it here.
  // `hier="-1"` is the non-OLAP marker Excel writes for a page field
  // bound to a cache field rather than to an OLAP hierarchy.
  if (table.page_fields().empty()) {
    const std::vector<std::uint32_t> page_order = table.page_field_order();
    if (!page_order.empty()) {
      out.append("<pageFields count=\"");
      out.append(std::to_string(page_order.size()));
      out.append("\">");
      for (const std::uint32_t idx : page_order) {
        out.append("<pageField fld=\"");
        out.append(std::to_string(idx));
        out.append("\" hier=\"-1\"/>");
      }
      out.append("</pageFields>");
    }
  }

  if (!table.data_fields().empty()) {
    AppendDataFields(out, table.data_fields());
  }

  // Re-emit the tail bin: unmodelled elements the reader captured after
  // `<dataFields>` (`<formats>`, `<conditionalFormats>`, `<chartFormats>`,
  // `<pivotTableStyleInfo>`, `<extLst>`, ...). The bytes are owned by the
  // table and re-appended verbatim so a read -> write round trip preserves
  // features v1.0 does not model structurally. Tables built from scratch
  // leave every bin empty.
  if (!table.raw_passthrough_xml().empty()) {
    out.append(table.raw_passthrough_xml());
  }

  out.append("</pivotTableDefinition>");
  return out;
}

}  // namespace formulon::io
