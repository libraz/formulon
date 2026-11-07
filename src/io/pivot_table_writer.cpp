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

#include "io/ooxml_defs.h"
#include "io/pivot_attr_names.h"
#include "io/pivot_table_writer_internal.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "io/xsd_double.h"
#include "pivot/field_lookup.h"
#include "pivot/filter_engine.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/value_order.h"
#include "utils/a1_ref.h"
#include "utils/strings.h"

namespace formulon::io {
namespace {

using detail::AppendAttrsInOrder;
using detail::AppendFilters;
using detail::AppendPivotField;
using detail::AttrList;
using detail::FilterFieldIndex;
using detail::HasFormOverride;
using detail::HasRawAttr;
using detail::ReEmitsSource;

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

/// Appends the `<ext>` holding the x14 `<pivotTableDefinition>`: the table's
/// alternative text and, when `hide_values_row`, the hidden caption row of the
/// Values field.
void AppendTableExt(std::string& out, const pivot::PivotTable& table, bool hide_values_row) {
  out.append("<ext uri=\"{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}\"");
  append_xml_ns_decl(out, "x14", kNsX14);
  out.append("><x14:pivotTableDefinition");
  if (!table.alt_text().empty()) {
    append_xml_attr(out, "altText", table.alt_text());
  }
  if (!table.alt_text_summary().empty()) {
    append_xml_attr(out, "altTextSummary", table.alt_text_summary());
  }
  if (hide_values_row) {
    out.append(" hideValuesRow=\"1\"");
  }
  append_xml_ns_decl(out, "xm", kNsExcelMain);
  out.append("/></ext>");
}

/// Appends the tail bin with the table's alternative text put on its x14
/// `<pivotTableDefinition>` extension, adding the extension when the bin
/// carries an `<extLst>` without one.
void AppendTailWithAltText(std::string& out, const pivot::PivotTable& table) {
  const std::string& tail = table.raw_passthrough_xml();
  const bool has_alt = !table.alt_text().empty() || !table.alt_text_summary().empty();
  if (!has_alt) {
    out.append(tail);
    return;
  }
  constexpr std::string_view kElement = "<x14:pivotTableDefinition";
  if (const std::size_t at = tail.find(kElement); at != std::string::npos) {
    const std::size_t name_end = at + kElement.size();
    out.append(tail, 0, name_end);
    if (!table.alt_text().empty()) {
      append_xml_attr(out, "altText", table.alt_text());
    }
    if (!table.alt_text_summary().empty()) {
      append_xml_attr(out, "altTextSummary", table.alt_text_summary());
    }
    out.append(tail, name_end, std::string::npos);
    return;
  }
  if (const std::size_t close = tail.rfind("</extLst>"); close != std::string::npos) {
    out.append(tail, 0, close);
    AppendTableExt(out, table, /*hide_values_row=*/false);
    out.append(tail, close, std::string::npos);
    return;
  }
  out.append(tail);
}

/// `<pivotTableDefinition>` attribute names in CT_pivotTableDefinition order.
constexpr std::string_view kPivotTableAttrOrder[] = {
    "name",
    "cacheId",
    "dataOnRows",
    "dataPosition",
    "authKey",
    "applyNumberFormats",
    "applyBorderFormats",
    "applyFontFormats",
    "applyPatternFormats",
    "applyAlignmentFormats",
    "applyWidthHeightFormats",
    "dataCaption",
    "grandTotalCaption",
    "errorCaption",
    "showError",
    "missingCaption",
    "showMissing",
    "pageStyle",
    "pivotTableStyle",
    "vacatedStyle",
    "tag",
    "updatedVersion",
    "minRefreshableVersion",
    "asteriskTotals",
    "showItems",
    "editData",
    "disableFieldList",
    "showCalcMbrs",
    "visualTotals",
    "showMultipleLabel",
    "showDataDropDown",
    "showDrill",
    "printDrill",
    "showMemberPropertyTips",
    "showDataTips",
    "enableWizard",
    "enableDrill",
    "enableFieldProperties",
    "preserveFormatting",
    "useAutoFormatting",
    "pageWrap",
    "pageOverThenDown",
    "subtotalHiddenItems",
    "rowGrandTotals",
    "colGrandTotals",
    "fieldPrintTitles",
    "itemPrintTitles",
    "mergeItem",
    "showDropZones",
    "createdVersion",
    "indent",
    "showEmptyRow",
    "showEmptyCol",
    "showHeaders",
    "compact",
    "outline",
    "outlineData",
    "compactData",
    "published",
    "gridDropZones",
    "immersive",
    "multipleFieldFilters",
    "chartFormat",
    "rowHeaderCaption",
    "colHeaderCaption",
    "fieldListSortAscending",
    "mdxSubqueries",
    "customListSort",
};

/// The items of a field the plan derives as discrete, bound to the derived
/// shared items; `field.items` verbatim for any other field.
///
/// An axis field with no authored items gets one per shared item, in its sort
/// order. Authored items resolve to a shared item by their cache index, else
/// by label (as `fm_sheet_add_pivot_field_item` binds), else by position;
/// shared items none of them names follow, visible.
std::vector<pivot::PivotItem> BoundItems(const pivot::PivotField& field, const PivotCacheFieldWritePlan* fp) {
  if (fp == nullptr || !fp->derived) {
    return field.items;
  }
  const std::size_t shared = fp->items.size();
  const bool on_axis = field.axis == pivot::PivotAxis::Row || field.axis == pivot::PivotAxis::Col ||
                       field.axis == pivot::PivotAxis::Page;
  std::vector<pivot::PivotItem> bound;
  std::vector<bool> used(shared, false);
  if (field.items.empty()) {
    if (!on_axis || !field.marker_items.empty()) {
      return bound;
    }
    std::vector<std::uint32_t> order = fp->asc_order;
    if (field.sort.manual || !field.sort.by_field.empty()) {
      for (std::size_t k = 0; k < order.size(); ++k) {
        order[k] = static_cast<std::uint32_t>(k);
      }
    } else if (!field.sort.ascending) {
      std::reverse(order.begin(), order.end());
    }
    for (const std::uint32_t k : order) {
      pivot::PivotItem item;
      item.has_cache_index = true;
      item.cache_index = k;
      bound.push_back(std::move(item));
    }
    return bound;
  }
  for (std::size_t pos = 0; pos < field.items.size(); ++pos) {
    pivot::PivotItem item = field.items[pos];
    if (!item.has_cache_index) {
      std::size_t target = pos;
      if (!item.name.empty()) {
        for (std::size_t k = 0; k < shared; ++k) {
          if (pivot::display_string(fp->items[k]) == item.name) {
            target = k;
            break;
          }
        }
      }
      item.has_cache_index = true;
      item.cache_index = static_cast<std::uint32_t>(target);
    }
    if (item.cache_index < shared) {
      used[item.cache_index] = true;
    }
    bound.push_back(std::move(item));
  }
  for (std::size_t k = 0; k < shared; ++k) {
    if (!used[k]) {
      pivot::PivotItem item;
      item.has_cache_index = true;
      item.cache_index = static_cast<std::uint32_t>(k);
      bound.push_back(std::move(item));
    }
  }
  return bound;
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
    append_xml_attr(out, "name", info.name);
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
      out.append("\"");
      append_xml_ns_decl(out, "x14", kNsX14);
      out.append(">");
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

/// Location attributes Excel derives from the report shape, used when the
/// model carries no authored value. `col_entries` / `row_entries` count the
/// real fields plus the Values marker on that axis.
struct LocationDefaults {
  std::uint32_t first_header_row = 1;
  std::uint32_t first_data_row = 1;
  std::uint32_t first_data_col = 1;
  std::optional<std::uint32_t> row_page_count;
  std::optional<std::uint32_t> col_page_count;
};

LocationDefaults ComputeLocationDefaults(const pivot::PivotTable& table, std::size_t row_entries,
                                         std::size_t col_entries, bool col_values_only) {
  LocationDefaults d;
  if (col_values_only) {
    // Only the Values marker on columns: no "Column Labels" header row.
    d.first_header_row = 0;
    d.first_data_row = 1;
  } else {
    d.first_data_row = static_cast<std::uint32_t>(1U + col_entries);
  }
  const bool col_fields = !table.col_field_order().empty();
  // `gridDropZones` keeps a row-label column even with an empty row axis.
  const bool grid_without_rows = table.grid_drop_zones() && row_entries == 0 && !table.data_fields().empty();
  if (col_fields && !table.show_headers() && !(table.data_fields().size() == 1U && !table.row_field_order().empty())) {
    // Hidden field captions drop the "Column Labels" row, unless a lone data
    // field's caption holds its corner.
    d.first_header_row -= 1U;
    d.first_data_row -= 1U;
  } else if (!col_fields && table.grid_drop_zones() && !(grid_without_rows && table.data_fields().size() == 1U)) {
    // The classic layout's drop-zone row sits above the header, except over
    // a report holding a lone data field and nothing else.
    d.first_header_row += 1U;
    d.first_data_row += 1U;
  }
  if (row_entries > 0) {
    d.first_data_col = table.layout() == pivot::PivotLayout::Compact ? 1U : static_cast<std::uint32_t>(row_entries);
  } else if (grid_without_rows) {
    d.first_data_col = 1U;
  } else if (col_entries > 0) {
    d.first_data_col = 0;
  }
  // Page fields fold into lines of `pageWrap` (0 is one line), down then over
  // unless `pageOverThenDown`.
  const std::size_t pages = table.page_field_order().size();
  if (pages > 0) {
    const std::size_t per_line = table.page_wrap() == 0U ? pages : std::min<std::size_t>(table.page_wrap(), pages);
    const auto lines = static_cast<std::uint32_t>((pages + per_line - 1U) / per_line);
    const auto across = static_cast<std::uint32_t>(per_line);
    d.row_page_count = table.page_over_then_down() ? lines : across;
    d.col_page_count = table.page_over_then_down() ? across : lines;
  }
  return d;
}

}  // namespace

std::string write_pivot_table_definition(const pivot::PivotTable& table, std::optional<PivotRenderedSpan> rendered_span,
                                         const PivotCacheWritePlan* plan) {
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
  out.append(kNsSpreadsheetMl);
  out.append("\"");
  // Modelled and captured attributes alike are written in
  // CT_pivotTableDefinition order, as Excel does; names the schema does not
  // list follow in captured order.
  AttrList attrs;
  const auto flag = [&attrs](bool on, const char* name, const char* value) {
    if (on) {
      attrs.emplace_back(name, value);
    }
  };
  attrs.emplace_back("name", table.name());
  attrs.emplace_back("cacheId", std::to_string(table.pivot_cache_id()));
  attrs.emplace_back("dataCaption", table.data_caption());
  if (table.grand_total_caption().has_value()) {
    attrs.emplace_back("grandTotalCaption", *table.grand_total_caption());
  }
  if (table.error_caption().has_value()) {
    attrs.emplace_back("errorCaption", *table.error_caption());
  }
  flag(table.show_error(), "showError", "1");
  if (table.missing_caption().has_value()) {
    attrs.emplace_back("missingCaption", *table.missing_caption());
  }
  flag(!table.show_missing(), "showMissing", "0");
  // Grand-total flags default to true in OOXML and in the model. Emit them
  // only when turned OFF so an explicit OFF state survives the round trip
  // (omitting them would let the reader's default flip the pivot back ON).
  // `rowGrandTotals` is the bottom grand total row, i.e. the model's column
  // totals; `colGrandTotals` the right-hand column, the model's row totals.
  flag(!table.grand_totals_cols(), "rowGrandTotals", "0");
  flag(!table.grand_totals_rows(), "colGrandTotals", "0");
  // The Values level sitting among the row fields is what `dataOnRows` states;
  // it is written from the field orders, never from a stored copy.
  flag(table.data_fields().size() >= 2U && table.row_values_position().has_value(), "dataOnRows", "1");
  flag(table.subtotal_hidden_items(), "subtotalHiddenItems", "1");
  flag(!table.show_data_drop_down(), "showDataDropDown", "0");
  flag(!table.show_drill(), "showDrill", "0");
  flag(table.print_drill(), "printDrill", "1");
  flag(!table.show_data_tips(), "showDataTips", "0");
  flag(!table.enable_drill(), "enableDrill", "0");
  flag(!table.preserve_formatting(), "preserveFormatting", "0");
  flag(table.auto_formatting(), "useAutoFormatting", "1");
  if (table.page_wrap() != 0U) {
    attrs.emplace_back("pageWrap", std::to_string(table.page_wrap()));
  }
  flag(table.page_over_then_down(), "pageOverThenDown", "1");
  flag(table.field_print_titles(), "fieldPrintTitles", "1");
  flag(table.item_print_titles(), "itemPrintTitles", "1");
  flag(table.merge_item(), "mergeItem", "1");
  if (table.indent() != 1U) {
    attrs.emplace_back("indent", std::to_string(table.indent()));
  }
  flag(table.show_empty_row(), "showEmptyRow", "1");
  flag(table.show_empty_col(), "showEmptyCol", "1");
  flag(!table.show_headers(), "showHeaders", "0");
  // Report layout as Excel writes it at table level, where it is the form
  // new fields take (`compact` defaults to true, `outline` to false). A form
  // read back that departs from it is re-emitted from `passthrough_attrs`.
  if (!HasFormOverride(table.passthrough_attrs())) {
    const bool outline = table.layout() != pivot::PivotLayout::Tabular;
    const bool compact = table.layout() == pivot::PivotLayout::Compact;
    flag(!compact, "compact", "0");
    flag(!compact, "compactData", "0");
    flag(outline, "outline", "1");
    flag(outline, "outlineData", "1");
  }
  flag(table.grid_drop_zones(), "gridDropZones", "1");
  flag(!table.multiple_field_filters(), "multipleFieldFilters", "0");
  flag(!table.custom_list_sort(), "customListSort", "0");
  // Without `updatedVersion` Excel adds a stray caption row above the
  // pivot on refresh; a captured value is re-emitted with the rest.
  flag(!HasRawAttr(table.passthrough_attrs(), "updatedVersion"), "updatedVersion", "8");
  for (const auto& attr : table.passthrough_attrs()) {
    if (attr.first != "dataOnRows" && attr.first != "dataPosition") {
      attrs.push_back(attr);
    }
  }
  AppendAttrsInOrder(out, attrs, kPivotTableAttrOrder);
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
      table, row_entries, col_entries, col_values_position.has_value() && table.col_field_order().empty());

  out.append("<location ref=\"");
  // Anchor is unconditionally emitted as a range; an empty table
  // (span_rows == span_cols == 0) round-trips through the single-cell
  // form via EncodeA1Range's zero-span guard.
  //
  // A span the reader decoded is Excel's own and is re-emitted verbatim.
  //
  // A span read from a file that an edit made stale is replaced by the
  // projection. Otherwise the projection acts as a floor rather than a replacement.
  // Under-sizing is the failure that matters: a `ref` smaller than the grid
  // the pivot draws makes Excel terminate when the report is refreshed, and
  // that is what a model assembled in memory produces, because nothing
  // revises the span it was created with as fields are added. Over-sizing
  // is not a failure -- Excel reserves the range and recomputes it on
  // refresh -- so a caller that deliberately set a wider span through
  // `fm_sheet_set_pivot_anchor` keeps it.
  std::uint32_t span_rows = table.span_rows();
  std::uint32_t span_cols = table.span_cols();
  std::uint32_t ref_top = table.anchor_row();
  if (rendered_span.has_value() && !table.has_authored_span()) {
    // A read span gone stale describes the report before the edit; the projection replaces it.
    const bool stale = table.span_read();
    span_rows = !stale && span_rows > rendered_span->rows ? span_rows : rendered_span->rows;
    span_cols = !stale && span_cols > rendered_span->cols ? span_cols : rendered_span->cols;
    // A page-field block taller than the anchor's headroom pushes the body down.
    ref_top = ref_top > rendered_span->top ? ref_top : rendered_span->top;
  }
  AppendXmlAttrEscaped(out, EncodeA1Range(ref_top, table.anchor_col(), span_rows, span_cols));
  out.append("\"");
  // ECMA-376 requires firstHeaderRow / firstDataRow / firstDataCol. Preserve
  // authored values, including explicit zero, while supplying the schema's
  // independent fallback for a model built without those attributes.
  append_xml_attr_uint_if(out, "firstHeaderRow",
                          table.location_first_header_row().value_or(loc_defaults.first_header_row));
  append_xml_attr_uint_if(out, "firstDataRow", table.location_first_data_row().value_or(loc_defaults.first_data_row));
  append_xml_attr_uint_if(out, "firstDataCol", table.location_first_data_col().value_or(loc_defaults.first_data_col));
  // rowPageCount / colPageCount are optional: a table without page fields omits them.
  append_xml_attr_uint_if(
      out, "rowPageCount",
      table.location_row_page_count() ? table.location_row_page_count() : loc_defaults.row_page_count);
  append_xml_attr_uint_if(
      out, "colPageCount",
      table.location_col_page_count() ? table.location_col_page_count() : loc_defaults.col_page_count);
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
    const PivotCacheFieldWritePlan* field_plan =
        plan != nullptr && i < plan->fields.size() ? &plan->fields[i] : nullptr;
    // A re-emitted filter leaves the field's `measureFilter` as it was read.
    const bool measure_filter =
        std::any_of(table.active_filters().begin(), table.active_filters().end(), [&](const pivot::PivotFilter& f) {
          return pivot::is_value_filter_type(f.type) && !ReEmitsSource(table, f) && FilterFieldIndex(table, f) == i;
        });
    AppendPivotField(out, table.fields()[i], BoundItems(table.fields()[i], field_plan), table.layout(), used_as_data,
                     measure_filter);
  }
  out.append("</pivotFields>");

  if (!table.row_field_order().empty() || table.row_values_position().has_value()) {
    AppendFieldOrder(out, "rowFields", table.row_field_order(), table.row_values_position());
  }
  // `<rowItems>` and `<colItems>` render the report as it was read. An edited
  // table leaves them out only when its cache asks Excel to rebuild the report
  // on open; otherwise nothing would ever redraw it.
  const bool rebuilt_on_open = table.layout_dirty() && plan != nullptr && plan->refresh_on_load;
  if (!rebuilt_on_open) {
    out.append(table.raw_passthrough_after_row_fields());
  }

  if (!table.col_field_order().empty() || col_values_position.has_value()) {
    AppendFieldOrder(out, "colFields", table.col_field_order(), col_values_position);
  }
  if (!rebuilt_on_open) {
    out.append(table.raw_col_items());
  }
  // `<pageFields>`, which the schema places before `<dataFields>`.
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
  AppendTailWithAltText(out, table);
  // Without the hidden-row extension Excel adds a caption row above the pivot
  // on refresh. A table read from a file keeps its own `<extLst>`.
  if (table.raw_passthrough_xml().find("<extLst") == std::string::npos) {
    const bool hide_values_row = !HasRawAttr(table.passthrough_attrs(), "updatedVersion");
    if (hide_values_row || !table.alt_text().empty() || !table.alt_text_summary().empty()) {
      out.append("<extLst>");
      AppendTableExt(out, table, hide_values_row);
      out.append("</extLst>");
    }
  }

  out.append("</pivotTableDefinition>");
  return out;
}

}  // namespace formulon::io
