//
// Implementation of the pivot-table-definition writer. See
// pivot_table_writer.h for the public contract; see
// pivot_table_reader.cpp for the symmetric grammar definition.

#include "io/pivot_table_writer.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <variant>
#include <vector>

#include "io/pivot_attr_names.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "pivot/field_lookup.h"
#include "pivot/filter_engine.h"
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

/// Excel's x14 extension that hides the caption row of the Values field.
constexpr std::string_view kValuesRowExtLst =
    "<extLst><ext uri=\"{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}\" "
    "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">"
    "<x14:pivotTableDefinition hideValuesRow=\"1\" "
    "xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"/></ext></extLst>";

/// True when `attrs` already carries `name` (an attribute captured on read).
bool HasRawAttr(const std::vector<std::pair<std::string, std::string>>& attrs, std::string_view name) {
  for (const auto& attr : attrs) {
    if (attr.first == name) {
      return true;
    }
  }
  return false;
}

/// True when `attrs` carries any of the report-form attributes the reader
/// keeps only as an override of the form `layout` derives.
bool HasFormOverride(const std::vector<std::pair<std::string, std::string>>& attrs) {
  return HasRawAttr(attrs, "compact") || HasRawAttr(attrs, "outline") || HasRawAttr(attrs, "compactData") ||
         HasRawAttr(attrs, "outlineData");
}

/// Emits one `<pivotField>` element. Self-closing when there are no
/// items; open/close pair otherwise.
///
/// `used_as_data` says whether some `<dataField fld>` refers to this field.
/// Excel refuses a `dataField="1"` field that no data field refers to, so the
/// flag follows the data-field list rather than the field's axis.
void AppendPivotField(std::string& out, const pivot::PivotField& field, pivot::PivotLayout layout, bool used_as_data) {
  out.append("<pivotField");
  if (field.axis != pivot::PivotAxis::Value && field.axis != pivot::PivotAxis::None) {
    append_xml_attr(out, "axis", enum_name(kPivotAxisNames, field.axis));
  }
  if (used_as_data) {
    out.append(" dataField=\"1\"");
  }
  // Excel writes `showAll="0"` on every field; without it a refresh shows
  // items that have no data. A value captured on read is re-emitted below.
  if (!HasRawAttr(field.passthrough_attrs, "showAll")) {
    out.append(" showAll=\"0\"");
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
  // Custom functions replace the automatic subtotal, and Excel then writes
  // neither `defaultSubtotal` nor a `t="default"` item. Without custom
  // functions, `defaultSubtotal` is emitted only when turned OFF so an
  // explicit suppression survives the round trip.
  const std::vector<pivot::SubtotalFn> subtotal_fns = pivot::ordered_subtotal_fns(field);
  const bool automatic_subtotal = field.default_subtotal && subtotal_fns.empty();
  if (!field.default_subtotal && subtotal_fns.empty()) {
    out.append(" defaultSubtotal=\"0\"");
  }
  for (const pivot::SubtotalFn fn : subtotal_fns) {
    const PivotSubtotalName* names = find_pivot_subtotal_name(fn);
    if (names == nullptr) {
      continue;
    }
    out.push_back(' ');
    out.append(names->attr);
    out.append("=\"1\"");
  }
  // An absent `sortType` reads back as manual (`<items>` order), so a
  // label-sorted axis field spells its direction out, as does a field off
  // the axes whose order was stated rather than defaulted.
  const bool on_axis = field.axis == pivot::PivotAxis::Row || field.axis == pivot::PivotAxis::Col ||
                       field.axis == pivot::PivotAxis::Page;
  if ((on_axis || field.sort.stated) && !field.sort.manual) {
    out.append(field.sort.ascending ? " sortType=\"ascending\"" : " sortType=\"descending\"");
  }
  // The report form is per field in Excel (`compact` / `outline` default to
  // true); every field carries the table's form, as Excel writes it, unless
  // the field was read with a form of its own (kept in `passthrough_attrs`).
  if (!HasFormOverride(field.passthrough_attrs)) {
    if (layout == pivot::PivotLayout::Tabular) {
      out.append(" compact=\"0\" outline=\"0\"");
    } else if (layout == pivot::PivotLayout::Outline) {
      out.append(" compact=\"0\"");
    }
  }
  // Re-emit any unmodelled `<pivotField>` attributes captured on read.
  append_raw_attrs(out, field.passthrough_attrs);
  if (field.items.empty() && field.marker_items.empty()) {
    if (field.raw_children_xml.empty()) {
      out.append("/>");
    } else {
      out.push_back('>');
      out.append(field.raw_children_xml);
      out.append("</pivotField>");
    }
    return;
  }
  // Excel closes a field's `<items>` with one entry per subtotal it shows:
  // `<item t="default"/>` for the implicit one, or a token per explicitly
  // selected function. Excel treats a field whose items lack its subtotal
  // entry as damaged and offers to repair the file, so the markers follow
  // `default_subtotal` / `subtotal_fns`: the ones read from the file are
  // re-emitted while they still spell that selection, and regenerated once
  // it changed. `grand` and `blank` have no model and are always kept.
  std::vector<std::string_view> expected;
  if (automatic_subtotal) {
    expected.emplace_back("default");
  }
  for (const pivot::SubtotalFn fn : subtotal_fns) {
    if (const PivotSubtotalName* names = find_pivot_subtotal_name(fn); names != nullptr) {
      expected.push_back(names->item);
    }
  }
  const auto is_subtotal = [](std::string_view type) {
    if (type == "default") {
      return true;
    }
    for (const PivotSubtotalName& e : kPivotSubtotalNames) {
      if (type == e.item) {
        return true;
      }
    }
    return false;
  };
  std::vector<std::string_view> authored;
  for (const pivot::PivotMarkerItem& marker : field.marker_items) {
    if (is_subtotal(marker.type)) {
      authored.emplace_back(marker.type);
    }
  }
  const bool keep_subtotals = authored == expected;
  const auto closes_list = [&](const pivot::PivotMarkerItem& marker) {
    return is_subtotal(marker.type) || marker.type == "grand";
  };
  std::size_t count = field.items.size() + (keep_subtotals ? 0U : expected.size());
  for (const pivot::PivotMarkerItem& marker : field.marker_items) {
    count += keep_subtotals || !is_subtotal(marker.type) ? 1U : 0U;
  }
  out.append("><items count=\"");
  out.append(std::to_string(count));
  out.append("\">");
  const auto append_blanks_at = [&](std::size_t position) {
    for (const pivot::PivotMarkerItem& marker : field.marker_items) {
      const std::size_t anchor = std::min(marker.after_items, field.items.size());
      if (!closes_list(marker) && anchor == position) {
        out.append(marker.xml);
      }
    }
  };
  for (std::size_t i = 0; i < field.items.size(); ++i) {
    append_blanks_at(i);
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
    append_raw_attrs(out, item.passthrough_attrs);
    out.append("/>");
  }
  append_blanks_at(field.items.size());
  if (keep_subtotals) {
    for (const pivot::PivotMarkerItem& marker : field.marker_items) {
      if (is_subtotal(marker.type)) {
        out.append(marker.xml);
      }
    }
  } else {
    for (const std::string_view type : expected) {
      out.append("<item t=\"");
      out.append(type);
      out.append("\"/>");
    }
  }
  for (const pivot::PivotMarkerItem& marker : field.marker_items) {
    if (marker.type == "grand") {
      out.append(marker.xml);
    }
  }
  out.append("</items>");
  out.append(field.raw_children_xml);
  out.append("</pivotField>");
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

/// Emits `<pivotTableStyleInfo>`; the name attribute is omitted when empty.
void AppendStyleInfo(std::string& out, const pivot::PivotStyleInfo& info) {
  const auto flag = [](bool on) { return on ? "1" : "0"; };
  out.append("<pivotTableStyleInfo");
  if (!info.name.empty()) {
    out.append(" name=\"");
    AppendXmlAttrEscaped(out, info.name);
    out.push_back('"');
  }
  out.append(" showRowHeaders=\"");
  out.append(flag(info.show_row_headers));
  out.append("\" showColHeaders=\"");
  out.append(flag(info.show_col_headers));
  out.append("\" showRowStripes=\"");
  out.append(flag(info.show_row_stripes));
  out.append("\" showColStripes=\"");
  out.append(flag(info.show_col_stripes));
  out.append("\" showLastColumn=\"");
  out.append(flag(info.show_last_column));
  out.append("\"/>");
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
    const std::string_view x14_show_as = enum_name(kPivotShowAsX14Names, df.show_as);
    if (df.show_as != pivot::ShowValuesAs::Normal && x14_show_as.empty()) {
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
    // The show-as extension carries the x14 mode plus whatever else the
    // source's `<x14:dataField>` held.
    const bool x14_ext = !x14_show_as.empty() || !df.x14_attrs.empty() || !df.x14_children_xml.empty();
    if (!x14_ext && df.raw_ext_xml.empty()) {
      out.append("/>");
      continue;
    }
    out.append("><extLst>");
    if (x14_ext) {
      out.append("<ext uri=\"");
      out.append(kPivotDataFieldExtUri);
      out.append("\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">");
      out.append("<x14:dataField");
      if (!x14_show_as.empty()) {
        append_xml_attr(out, "pivotShowAs", x14_show_as);
      }
      append_raw_attrs(out, df.x14_attrs);
      if (df.x14_children_xml.empty()) {
        out.append("/></ext>");
      } else {
        out.push_back('>');
        out.append(df.x14_children_xml);
        out.append("</x14:dataField></ext>");
      }
    }
    out.append(df.raw_ext_xml);
    out.append("</extLst></dataField>");
  }
  out.append("</dataFields>");
}

/// The `<pivotFields>` index a public filter is written against, or
/// `nullopt` when it designates nothing a file can name.
///
/// A value filter prunes the leaves of its axis: the named field when it sits
/// on that axis, otherwise the axis's innermost field, which is how the value
/// pass reads it. Label and date filters apply to the field they name.
std::optional<std::uint32_t> FilterFieldIndex(const pivot::PivotTable& table, const pivot::PivotFilter& filter) {
  const std::optional<std::size_t> named = pivot::resolve_field_by_any_name(table, filter.field_name);
  if (!named.has_value()) {
    return std::nullopt;
  }
  if (!pivot::is_value_filter_type(filter.type)) {
    return static_cast<std::uint32_t>(*named);
  }
  if (filter.data_field_index >= table.data_fields().size()) {
    return std::nullopt;
  }
  const std::vector<std::uint32_t>& order =
      filter.axis == pivot::PivotAxis::Row ? table.row_field_order() : table.col_field_order();
  if (order.empty()) {
    return std::nullopt;
  }
  if (std::find(order.begin(), order.end(), *named) != order.end()) {
    return static_cast<std::uint32_t>(*named);
  }
  return order.back();
}

/// A filter payload as the value pass reads it: `int` / `double` as-is,
/// anything else as zero. Empty when the number is not finite, which no
/// file can spell.
std::optional<double> FilterNumber(const std::variant<int, double, std::string>& value) {
  double number = 0.0;
  if (const auto* d = std::get_if<double>(&value)) {
    number = *d;
  } else if (const auto* i = std::get_if<int>(&value)) {
    number = static_cast<double>(*i);
  }
  return std::isfinite(number) ? std::optional<double>(number) : std::nullopt;
}

/// The upper bound of a range filter; empty when unset or not finite.
std::optional<double> FilterNumberHigh(const std::variant<std::monostate, int, double>& value) {
  if (const auto* d = std::get_if<double>(&value)) {
    return std::isfinite(*d) ? std::optional<double>(*d) : std::nullopt;
  }
  if (const auto* i = std::get_if<int>(&value)) {
    return static_cast<double>(*i);
  }
  return std::nullopt;
}

/// Appends `<customFilter operator="..." val="..."/>`; an empty `op` is the
/// schema's default `equal` and is omitted, as Excel does for a pattern.
void AppendCustomFilter(std::string& out, std::string_view op, std::string_view text) {
  out.append("<customFilter");
  if (!op.empty()) {
    append_xml_attr(out, "operator", op);
  }
  append_xml_attr(out, "val", text);
  out.append("/>");
}

void AppendCustomFilter(std::string& out, std::string_view op, double number) {
  out.append("<customFilter");
  append_xml_attr(out, "operator", op);
  append_xml_attr_number(out, "val", number);
  out.append("/>");
}

/// Serialises one public filter as Excel writes it from the filter dialogs.
/// Returns false, emitting nothing, for an entry that filters nothing and
/// that the file has no spelling for: a range with no upper bound, a
/// non-finite bound, a date bound given as text.
bool AppendGeneratedFilter(std::string& out, const pivot::PivotFilter& filter, std::uint32_t field_index,
                           std::uint32_t id) {
  std::string_view type;
  switch (filter.type) {
    case pivot::FilterType::ValueTop10:
      type = "count";
      break;
    case pivot::FilterType::ValueGreaterThan:
      type = "valueGreaterThan";
      break;
    case pivot::FilterType::ValueBetween:
      type = "valueBetween";
      break;
    case pivot::FilterType::LabelContains:
      type = "captionContains";
      break;
    case pivot::FilterType::LabelBeginsWith:
      type = "captionBeginsWith";
      break;
    case pivot::FilterType::LabelDate:
      type = "dateBetween";
      break;
    default:
      return false;  // Outside the public vocabulary; never added through the API.
  }
  const bool ranged = filter.type == pivot::FilterType::ValueBetween || filter.type == pivot::FilterType::LabelDate;
  const bool caption =
      filter.type == pivot::FilterType::LabelContains || filter.type == pivot::FilterType::LabelBeginsWith;
  // A date bound has no zero fallback: a text payload leaves the date filter
  // inert, and a written zero would not.
  const bool text_payload = std::holds_alternative<std::string>(filter.value);
  const std::optional<double> low = caption ? std::optional<double>(0.0)
                                    : filter.type == pivot::FilterType::LabelDate && text_payload
                                        ? std::nullopt
                                        : FilterNumber(filter.value);
  const std::optional<double> high = ranged ? FilterNumberHigh(filter.value_high) : std::optional<double>(0.0);
  if (!low || !high) {
    return false;
  }
  out.append("<filter");
  append_xml_attr_uint(out, "fld", field_index);
  append_xml_attr(out, "type", type);
  out.append(" evalOrder=\"-1\"");
  append_xml_attr_uint(out, "id", id);
  if (pivot::is_value_filter_type(filter.type)) {
    append_xml_attr_uint(out, "iMeasureFld", filter.data_field_index);
  }
  std::string_view operand;
  if (const auto* text = std::get_if<std::string>(&filter.value); text != nullptr && caption) {
    operand = *text;
  }
  if (caption) {
    append_xml_attr(out, "stringValue1", operand);
  }
  out.append("><autoFilter ref=\"A1\"><filterColumn colId=\"0\">");
  if (filter.type == pivot::FilterType::ValueTop10) {
    out.append("<top10");
    append_xml_attr_number(out, "val", *low);
    append_xml_attr_number(out, "filterVal", *low);
    out.append("/>");
  } else if (caption) {
    const pivot::CaptionPredicate predicate = filter.type == pivot::FilterType::LabelContains
                                                  ? pivot::CaptionPredicate::Contains
                                                  : pivot::CaptionPredicate::BeginsWith;
    out.append("<customFilters>");
    AppendCustomFilter(out, {}, *pivot::caption_wildcard_pattern(predicate, operand));
    out.append("</customFilters>");
  } else if (ranged) {
    out.append("<customFilters and=\"1\">");
    AppendCustomFilter(out, "greaterThanOrEqual", *low);
    AppendCustomFilter(out, "lessThanOrEqual", *high);
    out.append("</customFilters>");
  } else {
    out.append("<customFilters>");
    AppendCustomFilter(out, "greaterThan", *low);
    out.append("</customFilters>");
  }
  out.append("</filterColumn></autoFilter></filter>");
  return true;
}

/// Emits `<filters>`: the public entries and the retained ones, interleaved
/// back into the order they were read in. A decoded entry that still points
/// at the field and data field it was read with re-emits its source bytes;
/// so does one whose field name never resolved (the pivot had no cache to
/// name it from). Every other public entry is serialised from the model under
/// an `id` no sibling uses. Nothing is written when no entry survives.
void AppendFilters(std::string& out, const pivot::PivotTable& table) {
  const std::vector<pivot::PivotFilter>& active = table.active_filters();
  const std::vector<pivot::RetainedPivotFilter>& retained = table.retained_filters();
  if (active.empty() && retained.empty()) {
    return;
  }
  std::uint32_t next_id = 1;
  for (const pivot::PivotFilter& filter : active) {
    if (filter.source.has_value()) {
      next_id = std::max(next_id, filter.source->id + 1U);
    }
  }
  for (const pivot::RetainedPivotFilter& entry : retained) {
    next_id = std::max(next_id, entry.id + 1U);
  }
  std::string body;
  std::size_t count = 0;
  std::size_t next_retained = 0;
  const auto flush_retained = [&](std::size_t active_index) {
    for (; next_retained < retained.size() && retained[next_retained].active_before <= active_index; ++next_retained) {
      body.append(retained[next_retained].xml);
      ++count;
    }
  };
  for (std::size_t i = 0; i < active.size(); ++i) {
    flush_retained(i);
    const pivot::PivotFilter& filter = active[i];
    const std::optional<std::uint32_t> field_index = FilterFieldIndex(table, filter);
    if (filter.source.has_value() &&
        (filter.field_name.empty() ||
         (field_index == filter.source->field_index && filter.data_field_index == filter.source->data_field_index))) {
      body.append(filter.source->xml);
      ++count;
      continue;
    }
    if (field_index.has_value() && AppendGeneratedFilter(body, filter, *field_index, next_id)) {
      ++next_id;
      ++count;
    }
  }
  flush_retained(retained.size() + active.size());
  if (count == 0) {
    return;
  }
  out.append("<filters count=\"");
  out.append(std::to_string(count));
  out.append("\">");
  out.append(body);
  out.append("</filters>");
}

/// Location attributes Excel derives from the report shape, used when the
/// model carries no authored value. `col_entries` / `row_entries` count the
/// real fields plus the Values marker on that axis.
struct LocationDefaults {
  std::uint32_t first_header_row = 1;
  std::uint32_t first_data_row = 1;
  std::uint32_t first_data_col = 1;
};

LocationDefaults ComputeLocationDefaults(pivot::PivotLayout layout, std::size_t row_entries, std::size_t col_entries,
                                         bool col_values_only) {
  LocationDefaults d;
  if (col_values_only) {
    // Only the Values marker on columns: no "Column Labels" header row.
    d.first_header_row = 0;
    d.first_data_row = 1;
  } else {
    d.first_data_row = static_cast<std::uint32_t>(1U + col_entries);
  }
  if (row_entries > 0) {
    d.first_data_col = layout == pivot::PivotLayout::Compact ? 1U : static_cast<std::uint32_t>(row_entries);
  } else if (col_entries > 0) {
    d.first_data_col = 0;
  }
  return d;
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
  if (table.subtotal_hidden_items()) {
    out.append(" subtotalHiddenItems=\"1\"");
  }
  // Report layout as Excel writes it at table level, where it is the form
  // new fields take (`compact` defaults to true, `outline` to false). A form
  // read back that departs from it is re-emitted from `passthrough_attrs`.
  if (!HasFormOverride(table.passthrough_attrs())) {
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
  }
  // Without `updatedVersion` Excel adds a stray caption row above the
  // pivot on refresh; a captured value is re-emitted below.
  if (!HasRawAttr(table.passthrough_attrs(), "updatedVersion")) {
    out.append(" updatedVersion=\"8\"");
  }
  // Re-emit any unmodelled root attributes captured on read.
  append_raw_attrs(out, table.passthrough_attrs());
  out.append(">");

  // Excel refuses a definition with two or more data fields unless the
  // Values pseudo-field sits on an axis; default it to the end of the
  // column axis when the model records no position.
  std::optional<std::size_t> col_values_position = table.col_values_position();
  if (table.data_fields().size() >= 2U && !table.row_values_position().has_value() &&
      !col_values_position.has_value()) {
    col_values_position = table.col_field_order().size();
  }
  const std::size_t row_entries = table.row_field_order().size() + (table.row_values_position().has_value() ? 1U : 0U);
  const std::size_t col_entries = table.col_field_order().size() + (col_values_position.has_value() ? 1U : 0U);
  const LocationDefaults loc_defaults = ComputeLocationDefaults(
      table.layout(), row_entries, col_entries, col_values_position.has_value() && table.col_field_order().empty());

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
  AppendOptionalLocationAttr(out, "firstHeaderRow",
                             table.location_first_header_row().value_or(loc_defaults.first_header_row));
  AppendOptionalLocationAttr(out, "firstDataRow",
                             table.location_first_data_row().value_or(loc_defaults.first_data_row));
  AppendOptionalLocationAttr(out, "firstDataCol",
                             table.location_first_data_col().value_or(loc_defaults.first_data_col));
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
  for (std::size_t i = 0; i < table.fields().size(); ++i) {
    bool used_as_data = false;
    for (const pivot::PivotDataField& df : table.data_fields()) {
      used_as_data = used_as_data || df.field_index == i;
    }
    AppendPivotField(out, table.fields()[i], table.layout(), used_as_data);
  }
  out.append("</pivotFields>");

  if (!table.row_field_order().empty() || table.row_values_position().has_value()) {
    AppendFieldOrder(out, "rowFields", table.row_field_order(), table.row_values_position());
  }
  // `<rowItems>` and any other post-rowFields unmodelled elements.
  out.append(table.raw_passthrough_after_row_fields());

  if (!table.col_field_order().empty() || col_values_position.has_value()) {
    AppendFieldOrder(out, "colFields", table.col_field_order(), col_values_position);
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

  // Re-emit the post-dataFields bins around `<pivotTableStyleInfo>`: the
  // unmodelled elements the schema places before it (`<formats>`,
  // `<conditionalFormats>`, `<chartFormats>`, `<pivotHierarchies>`), then
  // the style, the modelled `<filters>`, then the rest (hierarchy usage,
  // `<extLst>`, ...). The bin bytes are owned by the table and re-appended
  // verbatim so a read -> write round trip preserves features this version
  // does not model structurally. Tables built from scratch leave every bin
  // empty.
  out.append(table.raw_passthrough_before_style_info());
  if (table.style_info().has_value()) {
    AppendStyleInfo(out, *table.style_info());
  }
  AppendFilters(out, table);
  if (!table.raw_passthrough_xml().empty()) {
    out.append(table.raw_passthrough_xml());
  }
  // Without this extension Excel adds a caption row above the pivot on
  // refresh. A table read from a file keeps its own `<extLst>`.
  if (!HasRawAttr(table.passthrough_attrs(), "updatedVersion") &&
      table.raw_passthrough_xml().find("<extLst") == std::string::npos) {
    out.append(kValuesRowExtLst);
  }

  out.append("</pivotTableDefinition>");
  return out;
}

}  // namespace formulon::io
