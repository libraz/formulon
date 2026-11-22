#include "io/xlsb/pivot_table_bridge.h"

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <iterator>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "io/auto_filter_xml.h"
#include "io/ooxml_defs.h"
#include "io/pivot_attr_names.h"
#include "io/pivot_built_defaults.h"
#include "io/pivot_sort_scope.h"
#include "io/pivot_table_reader.h"
#include "io/pivot_table_writer_internal.h"
#include "io/xlsb/auto_filter_records.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/guid_text.h"
#include "io/xlsb/pivot_area_records.h"
#include "io/xlsb/pivot_line_item_records.h"
#include "io/xlsb/pivot_records_common.h"
#include "io/xlsb/pivot_table_bridge_internal.h"
#include "io/xlsb/pivot_table_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xml_utils.h"
#include "pivot/filter_engine.h"
#include "pivot/pivot_types.h"
#include "utils/a1_ref.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using namespace pivot_table_bridge_detail;

Expected<void, Error> CheckView(const PivotTableDoc& doc) {
  const SxView& v = doc.view;
  const auto paired = [&v](std::uint8_t pair) {
    const std::uint8_t bits = static_cast<std::uint8_t>(v.flags8 & pair);
    return bits == 0U || bits == pair;
  };
  if (v.flags1 != kViewFlags1 || v.flags2 != kViewFlags2 || (v.flags4 & ~kViewFlags4Named) != kViewFlags4 ||
      (v.flags7 & ~kViewApplyWidthHeight) != 0U || (v.flags5 & ~kViewFlags5Named) != kViewFlags5Fixed ||
      (v.flags6 & ~kViewFlags6Named) != kViewFlags6Fixed || (v.flags8 & ~kViewFlags8Named) != 0U ||
      (v.flags9 & ~kSxViewSingleFieldFilter) != 0U || v.flags10 != 0U || v.flags11 != 0U ||
      v.reserved20 != kViewReserved20 || !paired(kViewCompactPair) || !paired(kViewOutlinePair)) {
    return Unmeasured("view header bits");
  }
  if (v.data_position != kNoDataPosition || (v.data_axis != kDataOnRows && v.data_axis != kDataOnCols)) {
    return Unmeasured("data axis or position");
  }
  if ((doc.view14 &&
       (!SameBytes(doc.view14->frt_payload, kView14Frt) || !SameBytes(doc.view14->payload, kView14Payload))) ||
      !SameBytes(doc.view16.frt_payload, kView16Frt) || !SameBytes(doc.view16.payload, kView16Payload) ||
      (doc.version_info && (!SameBytes(doc.version_info->frt_payload, pivot_records::kVersionInfoFrtVersion) ||
                            !SameBytes(doc.version_info->payload, pivot_records::kRichDataFeature)))) {
    return Unmeasured("table future records");
  }
  return Expected<void, Error>::Ok();
}

/// Root attributes as the xlsx reader captures them from Excel's part: the
/// namespace bookkeeping of `xr:uid`, then document order.
void ReadRootAttrs(const PivotTableDoc& doc, AttrList& attrs) {
  if (doc.uid) {
    attrs.emplace_back("xmlns:mc", std::string(kNsMarkupCompat));
    attrs.emplace_back("xmlns:xr", std::string(kNsXr));
    attrs.emplace_back("mc:Ignorable", std::string(kXrPrefix));
    attrs.emplace_back("xr:uid", FormatGuid(doc.uid->data()));
  }
  for (std::size_t i = 0; i < std::size(kApplyFormats); ++i) {
    const bool on = i == kApplyWidthHeight && (doc.view.flags7 & kViewApplyWidthHeight) != 0U;
    attrs.emplace_back(std::string(kApplyFormats[i]), Flag(on));
  }
  attrs.emplace_back("updatedVersion", std::to_string(doc.view.updated_version));
  // Excel leaves out a version of 0, the schema default.
  if (doc.view.min_refreshable_version != 0U) {
    attrs.emplace_back("minRefreshableVersion", std::to_string(doc.view.min_refreshable_version));
  }
  if ((doc.view.flags4 & kViewAsteriskTotals) != 0U) {
    attrs.emplace_back("asteriskTotals", "1");
  }
  if (doc.view.created_version != 0U) {
    attrs.emplace_back("createdVersion", std::to_string(doc.view.created_version));
  }
}

/// Appends the form attributes that depart from `layout`, as the xlsx reader keeps them.
void CaptureForm(bool compact, bool outline, pivot::PivotLayout layout, bool root, AttrList& out) {
  if (compact == (layout == pivot::PivotLayout::Compact) && outline == (layout != pivot::PivotLayout::Tabular)) {
    return;
  }
  out.emplace_back("compact", Flag(compact));
  if (root) {
    out.emplace_back("compactData", Flag(compact));
  }
  out.emplace_back("outline", Flag(outline));
  if (root) {
    out.emplace_back("outlineData", Flag(outline));
  }
}

Expected<void, Error> ReadLocation(const SxLocation& loc, pivot::PivotTable& table) {
  if (loc.last_row < loc.first_row || loc.last_col < loc.first_col || loc.last_row >= a1::kMaxRows ||
      loc.last_col >= a1::kMaxCols) {
    return Corrupt("location rectangle is inverted or past the grid");
  }
  if (loc.first_header_row < loc.first_row || loc.first_data_row < loc.first_row ||
      loc.first_data_col < loc.first_col || loc.first_header_row > loc.last_row || loc.first_data_row > loc.last_row ||
      loc.first_data_col > loc.last_col) {
    return Corrupt("location offsets fall outside the table");
  }
  table.set_anchor(loc.first_row, loc.first_col, loc.last_row - loc.first_row + 1U, loc.last_col - loc.first_col + 1U);
  table.mark_span_authored();
  const auto count = [](std::uint32_t v) { return v == 0U ? std::nullopt : std::optional<std::uint32_t>(v); };
  table.set_location_attributes(loc.first_header_row - loc.first_row, loc.first_data_row - loc.first_row,
                                loc.first_data_col - loc.first_col, count(loc.row_page_count),
                                count(loc.col_page_count));
  return Expected<void, Error>::Ok();
}

Expected<void, Error> ReadItems(const SxField& in, pivot::PivotField& field) {
  if (!in.items) {
    return Expected<void, Error>::Ok();
  }
  if (in.items->empty()) {
    return Unmeasured("empty field item list");
  }
  for (const SxItem& item : *in.items) {
    if (item.kind == kItemData) {
      pivot::PivotItem entry;
      entry.visible = (item.flags & kSxItemHidden) == 0U;
      entry.caption = item.caption.value_or(std::string());
      entry.blank_caption = item.caption && item.caption->empty();
      entry.show_detail = (item.flags & kSxItemCollapsed) == 0U;
      entry.calculated = (item.flags & kSxItemCalculated) != 0U;
      if ((item.flags & kSxItemMissing) != 0U) {
        entry.passthrough_attrs.emplace_back("m", "1");
      }
      entry.has_cache_index = true;
      entry.cache_index = item.index;
      field.items.push_back(std::move(entry));
      continue;
    }
    const bool collapsed = (item.flags & kSxItemCollapsed) != 0U;
    if (item.index != kNoItemIndex || item.caption || (item.flags & ~kSxItemCollapsed) != 0U ||
        (collapsed && item.kind != kItemDefault)) {
      return Unmeasured("subtotal item flags");
    }
    const std::string_view type = SubtotalItemType(item.kind);
    std::string xml("<item t=\"");
    xml.append(type);
    xml.append(collapsed ? "\" sd=\"0\"/>" : "\"/>");
    field.marker_items.push_back(
        pivot::PivotMarkerItem{std::move(xml), std::string(type), field.items.size(), !collapsed});
  }
  return Expected<void, Error>::Ok();
}

/// The field's `<autoSortScope>` as the xlsx reader keeps it in `raw_children_xml`.
Expected<void, Error> ReadAutoSortScope(const RawRecords& block, std::string& out) {
  if (block.empty()) {
    return Expected<void, Error>::Ok();
  }
  if (block.size() < 3U || !block.front().payload.empty() || !block.back().payload.empty() ||
      block.back().type != ids::kEndAutoSortScope) {
    return Unmeasured("sort scope framing");
  }
  std::size_t pos = 1;
  ASSIGN_OR_RETURN(const pivot_area::PivotArea area, pivot_area::decode_pivot_area(block, &pos));
  if (pos + 1U != block.size() || area.field != pivot_area::kNoField || area.type != pivot_area::kAreaTypeNormal ||
      (area.flags & ~pivot_area::kAreaDataOnly) != 0U || !area.filters || area.filters->empty()) {
    return Unmeasured("sort scope pivot area shape");
  }
  out.append("<autoSortScope><pivotArea");
  if ((area.flags & pivot_area::kAreaDataOnly) == 0U) {
    out.append(" dataOnly=\"0\"");
  }
  // Excel's one sorted-by-value scope carries outline="0" fieldPosition="0" with no bit for either.
  out.append(" outline=\"0\" fieldPosition=\"0\"><references");
  append_xml_attr_uint(out, "count", area.filters->size());
  out.push_back('>');
  for (const pivot_area::PivotAreaFilter& ref : *area.filters) {
    if ((ref.flags & kAreaFilterAlways) == 0U ||
        (ref.flags & ~(kAreaFilterAlways | pivot_area::kFilterSelected)) != 0U) {
      return Unmeasured("sort scope reference flags");
    }
    out.append("<reference");
    append_xml_attr_uint(out, "field", ref.field);
    append_xml_attr_uint(out, "count", ref.items.size());
    if ((ref.flags & pivot_area::kFilterSelected) == 0U) {
      out.append(" selected=\"0\"");
    }
    out.push_back('>');
    for (const std::uint32_t x : ref.items) {
      out.append("<x");
      append_xml_attr_uint(out, "v", x);
      out.append("/>");
    }
    out.append("</reference>");
  }
  out.append("</references></pivotArea></autoSortScope>");
  return Expected<void, Error>::Ok();
}

/// A format's pivot area as the xlsx reader types Excel's `<pivotArea>`. The
/// codec has already refused every shape no Excel part shows.
pivot::PivotTableArea ReadTableArea(const pivot_area::PivotArea& in) {
  pivot::PivotTableArea area;
  if (in.field != pivot_area::kNoField) {
    area.field = in.field;
  }
  for (const AreaTypeCode& entry : kAreaTypes) {
    if (entry.code == in.type) {
      area.type = entry.type;
    }
  }
  const auto has = [&in](std::uint32_t bit) { return (in.flags & bit) != 0U; };
  area.data_only = has(pivot_area::kAreaDataOnly);
  area.label_only = has(pivot_area::kAreaLabelOnly);
  area.grand_row = has(pivot_area::kAreaGrandRow);
  area.grand_col = has(pivot_area::kAreaGrandCol);
  area.collapsed_levels_are_subtotals = has(pivot_area::kAreaCollapsedLevelsAreSubtotals);
  if ((in.flags & pivot_area::kAreaAxisMask) == pivot_area::kAreaAxisRow) {
    area.axis = pivot::PivotAxis::Row;
  }
  // Every Excel-saved area carries outline="0" fieldPosition="0" with no bit for either.
  area.outline = false;
  area.field_position = 0U;
  if (in.filters) {
    for (const pivot_area::PivotAreaFilter& filter : *in.filters) {
      pivot::PivotTableAreaReference ref;
      ref.field = filter.field;
      ref.items = filter.items;
      ref.selected = (filter.flags & pivot_area::kFilterSelected) != 0U;
      area.references.push_back(std::move(ref));
    }
  }
  return area;
}

void ReadFormats(const PivotTableDoc& doc, pivot::PivotTable& table) {
  for (const SxFormat& in : doc.formats) {
    pivot::PivotFormat format;
    format.dxf_id = in.dxf_id;
    format.area = ReadTableArea(in.area);
    table.mutable_formats().push_back(std::move(format));
  }
  for (const SxChartFormat& in : doc.chart_formats) {
    pivot::PivotChartFormat format;
    format.format = in.format;
    format.series = true;
    format.area = ReadTableArea(in.area);
    table.mutable_chart_formats().push_back(std::move(format));
  }
}

/// The table extLst the xlsx reader keeps for Excel's part.
std::string TableExtLst(bool hide_values_row, bool rich_version_info) {
  std::string out("<extLst>");
  if (hide_values_row) {
    out.append(kView14Ext);
  }
  out.append(kView16Ext);
  if (rich_version_info) {
    out.append(kVersionInfoExt);
  }
  out.append("</extLst>");
  return out;
}

Expected<pivot::PivotField, Error> ReadField(const SxField& in) {
  constexpr std::uint8_t kAxes = kSxFieldAxisRow | kSxFieldAxisCol | kSxFieldAxisPage;
  const std::uint8_t axis = static_cast<std::uint8_t>(in.axis & kAxes);
  if ((in.axis & ~(kAxes | kSxFieldAxisData)) != 0U || (axis & (axis - 1U)) != 0U) {
    return Corrupt("field sits on several axes");
  }
  const std::uint32_t drag = in.flags & kFieldDragPair;
  if ((in.name_flags & ~(kSxFieldCompact | kSxFieldHasDisplayName)) != 0U || (in.flags & ~kFieldNamed) != kFieldFixed ||
      (drag != 0U && drag != kFieldDragPair) || in.item_page_count != kItemPageCount ||
      in.reserved16 != kFieldReserved16 || (in.subtotals & ~kSubtotalMask) != 0U ||
      (in.display_name && in.display_name->empty())) {
    return Unmeasured("field header bits");
  }
  pivot::PivotField field;
  field.axis = axis == kSxFieldAxisRow              ? pivot::PivotAxis::Row
               : axis == kSxFieldAxisCol            ? pivot::PivotAxis::Col
               : axis == kSxFieldAxisPage           ? pivot::PivotAxis::Page
               : (in.axis & kSxFieldAxisData) != 0U ? pivot::PivotAxis::Value
                                                    : pivot::PivotAxis::None;
  field.custom_name = in.display_name.value_or(std::string());
  if (in.num_fmt != 0U) {
    field.number_format = std::to_string(in.num_fmt);
  }
  field.subtotal_top = (in.flags & kSxFieldSubtotalTop) != 0U;
  field.default_subtotal = (in.subtotals & kSxFieldSubtotalDefault) != 0U;
  for (std::size_t k = 0; k < kPivotSubtotalNames.size(); ++k) {
    if ((in.subtotals & (1U << (k + 1U))) != 0U) {
      field.subtotal_fns.push_back(kPivotSubtotalNames[k].fn);
    }
  }
  if (field.default_subtotal && !field.subtotal_fns.empty()) {
    return Unmeasured("default subtotal together with custom functions");
  }
  // Windows Excel sets the ascending bit on manually ordered fields too; it is derived on write.
  if ((in.flags & kSxFieldAutoSort) != 0U) {
    field.sort.stated = true;
    field.sort.ascending = (in.flags & kSxFieldAscendSort) != 0U;
  } else {
    field.sort.manual = true;
  }
  RETURN_IF_ERROR(ReadItems(in, field));
  field.insert_blank_row = (in.flags & kSxFieldInsertBlankRow) != 0U;
  field.show_all = (in.flags & kSxFieldShowAll) != 0U;
  if ((in.flags & kFieldDragToRow) == 0U) {
    field.passthrough_attrs.emplace_back("dragToRow", "0");
  }
  if (drag == 0U) {
    field.passthrough_attrs.emplace_back("dragToCol", "0");
    field.passthrough_attrs.emplace_back("dragToPage", "0");
  }
  if ((in.flags & kSxFieldMeasureFilter) != 0U) {
    field.passthrough_attrs.emplace_back("measureFilter", "1");
  }
  // Measured on manually ordered fields only, where Excel also sets the ascending bit.
  if ((in.flags & kFieldNewItemsOutsideFilter) == 0U) {
    if ((in.flags & kSxFieldAutoSort) != 0U || (in.flags & kSxFieldAscendSort) == 0U) {
      return Unmeasured("new items in a filter on a field that is not manually ordered");
    }
    field.passthrough_attrs.emplace_back("includeNewItemsInFilter", "1");
  }
  RETURN_IF_ERROR(ReadAutoSortScope(in.auto_sort_scope, field.raw_children_xml));
  field.fill_down_labels = in.fill_down_labels;
  return field;
}

Expected<void, Error> ReadFieldOrder(const std::optional<std::vector<std::uint32_t>>& in, std::size_t field_count,
                                     std::vector<std::uint32_t>& order, std::optional<std::size_t>& values_position) {
  if (!in) {
    return Expected<void, Error>::Ok();
  }
  if (in->empty()) {
    return Unmeasured("empty axis field list");
  }
  for (std::size_t i = 0; i < in->size(); ++i) {
    const std::uint32_t x = (*in)[i];
    if (x == kValuesField) {
      if (values_position) {
        return Corrupt("axis field list names the Values field twice");
      }
      values_position = i;
      continue;
    }
    if (x >= field_count) {
      return Corrupt("axis field list names a field the table does not have");
    }
    order.push_back(x);
  }
  return Expected<void, Error>::Ok();
}

/// `<rowItems>` / `<colItems>` spelled as Excel writes them.
std::string LineItemsXml(const SxLineItems& block, std::string_view tag) {
  // Excel writes neither the block nor the element for an axis with no line items.
  if (block.items.empty()) {
    return std::string();
  }
  std::string out("<");
  out.append(tag);
  append_xml_attr_uint(out, "count", block.items.size());
  out.push_back('>');
  for (const SxLineItem& line : block.items) {
    out.append("<i");
    if (line.t != 0U) {
      append_xml_attr(out, "t", kLineItemTypes[line.t]);
    }
    if (line.r != 0U) {
      append_xml_attr_uint(out, "r", line.r);
    }
    if (line.i != 0U) {
      append_xml_attr_uint(out, "i", line.i);
    }
    if (line.x.empty()) {
      out.append("/>");
      continue;
    }
    out.push_back('>');
    for (const std::uint32_t v : line.x) {
      out.append("<x");
      if (v != 0U) {
        append_xml_attr_uint(out, "v", v);
      }
      out.append("/>");
    }
    out.append("</i>");
  }
  out.append("</");
  out.append(tag);
  out.push_back('>');
  return out;
}

Expected<void, Error> ReadDataFields(const PivotTableDoc& doc, pivot::PivotTable& table) {
  for (const SxDataField& in : doc.data_fields) {
    if (in.field >= doc.fields.size()) {
      return Corrupt("data field names a field the table does not have");
    }
    if (in.subtotal >= kPivotAggregationNames.size() || in.show_data_as >= std::size(kShowDataAs)) {
      return Unmeasured("data field function or display mode");
    }
    pivot::PivotDataField entry;
    entry.name = in.name;
    entry.field_index = in.field;
    entry.aggregation = static_cast<pivot::Aggregation>(kPivotAggregationNames[in.subtotal].value);
    if (in.num_fmt != 0U) {
      entry.number_format = std::to_string(in.num_fmt);
    }
    // The codec admits SXDI14 only with a known code, beside a plain SXDI display mode.
    entry.show_as = in.pivot_show_as ? kPivotShowAs[*in.pivot_show_as - kSxDataShowAsPercentOfParentRow]
                                     : kShowDataAs[in.show_data_as];
    entry.show_as_base_field = in.base_field;
    entry.show_as_base_item = in.base_item;
    table.mutable_data_fields().push_back(std::move(entry));
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> ReadStyle(const SxStyleClient& in, pivot::PivotTable& table) {
  constexpr std::uint16_t kNamed =
      kStyleFirstColumn | kStyleLastColumn | kStyleRowStripes | kStyleColStripes | kStyleRowHeaders | kStyleColHeaders;
  if ((in.flags & ~kNamed) != 0U || (in.name && in.name->empty())) {
    return Unmeasured("style client bits");
  }
  pivot::PivotStyleInfo info;
  info.name = in.name.value_or(std::string());
  info.show_last_column = (in.flags & kStyleLastColumn) != 0U;
  info.show_row_stripes = (in.flags & kStyleRowStripes) != 0U;
  info.show_col_stripes = (in.flags & kStyleColStripes) != 0U;
  info.show_row_headers = (in.flags & kStyleRowHeaders) != 0U;
  info.show_col_headers = (in.flags & kStyleColHeaders) != 0U;
  table.set_style_info(std::move(info));
  return Expected<void, Error>::Ok();
}

/// Each listed field sits on its list's axis, and a field carries the data bit
/// exactly when a data field refers to it, as the writer derives both.
Expected<void, Error> CheckAxes(const PivotTableDoc& doc, const pivot::PivotTable& table) {
  const auto listed_on = [&table](const std::vector<std::uint32_t>& order, pivot::PivotAxis axis) {
    return std::all_of(order.begin(), order.end(),
                       [&table, axis](std::uint32_t f) { return table.fields()[f].axis == axis; });
  };
  const bool pages_ok =
      std::all_of(table.page_fields().begin(), table.page_fields().end(), [&table](const pivot::PivotPageField& p) {
        return table.fields()[p.field_index].axis == pivot::PivotAxis::Page;
      });
  if (!listed_on(table.row_field_order(), pivot::PivotAxis::Row) ||
      !listed_on(table.col_field_order(), pivot::PivotAxis::Col) || !pages_ok) {
    return Corrupt("axis field list names a field on another axis");
  }
  for (std::size_t i = 0; i < doc.fields.size(); ++i) {
    const bool referenced =
        std::any_of(doc.data_fields.begin(), doc.data_fields.end(), [i](const SxDataField& d) { return d.field == i; });
    if (referenced != ((doc.fields[i].axis & kSxFieldAxisData) != 0U)) {
      return Corrupt("field data bit disagrees with the data fields");
    }
  }
  return Expected<void, Error>::Ok();
}

/// One SXFILTER (BrtBeginSXFilter payload `begin`, just read) as its `<filter>` element.
Expected<void, Error> AppendFilterXml(ByteSpan& cursor, ByteSpan begin, std::size_t field_count,
                                      std::size_t data_field_count, std::string& out) {
  ASSIGN_OR_RETURN(const std::uint32_t fld, read_u32(begin));
  ASSIGN_OR_RETURN(const std::uint32_t mp_fld, read_u32(begin));
  ASSIGN_OR_RETURN(const std::uint32_t type, read_u32(begin));
  ASSIGN_OR_RETURN(const std::uint32_t eval_order, read_u32(begin));
  ASSIGN_OR_RETURN(const std::uint32_t id, read_u32(begin));
  ASSIGN_OR_RETURN(const std::uint32_t measure_fld, read_u32(begin));
  ASSIGN_OR_RETURN(const std::uint32_t measure_hier, read_u32(begin));
  ASSIGN_OR_RETURN(const std::uint16_t flags, read_u16(begin));
  if (type == 0U || type >= std::size(kPivotFilterTypes) || (flags & ~0x000FU) != 0U) {
    return Unmeasured("filter type or string flags");
  }
  if (fld >= field_count || (measure_fld != kNoFilterIndex && measure_fld >= data_field_count)) {
    return Corrupt("filter names a field or data field the table does not have");
  }
  out.append("<filter");
  append_xml_attr_uint(out, "fld", fld);
  append_xml_attr_uint_if(out, "mpFld", mp_fld, kNoFilterIndex);
  append_xml_attr(out, "type", kPivotFilterTypes[type]);
  append_xml_attr_int(out, "evalOrder", static_cast<std::int32_t>(eval_order));
  append_xml_attr_uint(out, "id", id);
  append_xml_attr_uint_if(out, "iMeasureHier", measure_hier, kNoFilterIndex);
  append_xml_attr_uint_if(out, "iMeasureFld", measure_fld, kNoFilterIndex);
  for (std::size_t bit = 0; bit < std::size(kFilterStringAttrs); ++bit) {
    if ((flags & (1U << bit)) != 0U) {
      ASSIGN_OR_RETURN(const std::string text, read_xlwidestring(begin));
      append_xml_attr(out, kFilterStringAttrs[bit], text);
    }
  }
  if (begin.size != 0U) {
    return Unmeasured("filter header trailing data");
  }
  out.push_back('>');
  // Excel wraps a null uid ahead of each filter's AutoFilter.
  ASSIGN_OR_RETURN(const XlsbRecord ac, read_record(cursor));
  ASSIGN_OR_RETURN(const XlsbRecord uid, read_record(cursor));
  ASSIGN_OR_RETURN(const XlsbRecord ac_end, read_record(cursor));
  const bool null_uid =
      uid.payload.size == kGuidBytes &&
      std::all_of(uid.payload.data, uid.payload.data + uid.payload.size, [](std::uint8_t b) { return b == 0U; });
  if (ac.type != ids::kACBegin || ac.payload.size != kAcUidBeginPayload.size() ||
      !std::equal(kAcUidBeginPayload.begin(), kAcUidBeginPayload.end(), ac.payload.data) || uid.type != ids::kUid ||
      !null_uid || ac_end.type != ids::kACEnd || ac_end.payload.size != 0U) {
    return Unmeasured("filter uid block");
  }
  ASSIGN_OR_RETURN(const XlsbRecord afilter, read_record(cursor));
  if (afilter.type != kBeginAFilter) {
    return Corrupt("filter carries no AutoFilter");
  }
  ASSIGN_OR_RETURN(const AutoFilter criteria, read_auto_filter_block(cursor, afilter.payload));
  out.append(serialize_auto_filter(criteria));
  ASSIGN_OR_RETURN(const XlsbRecord end, read_record(cursor));
  if (end.type != ids::kEndSXFilter || end.payload.size != 0U) {
    return Unmeasured("filter future record");
  }
  out.append("</filter>");
  return Expected<void, Error>::Ok();
}

/// SXFILTERS spelled as the `<filters>` element the xlsx reader decodes.
Expected<std::string, Error> FiltersXml(const RawRecords& block, std::size_t field_count,
                                        std::size_t data_field_count) {
  std::vector<std::uint8_t> bytes;
  for (const ids::RawRecord& rec : block) {
    emit_record(bytes, rec.type, rec.payload);
  }
  ByteSpan cursor{bytes.data(), bytes.size()};
  ASSIGN_OR_RETURN(const XlsbRecord begin, read_record(cursor));
  if (begin.type != ids::kBeginSXFilters) {
    return Corrupt("filter list does not open with its count");
  }
  ByteSpan count_payload = begin.payload;
  ASSIGN_OR_RETURN(const std::uint32_t declared, read_u32(count_payload));
  std::string out("<filters>");
  for (std::uint32_t seen = 0;; ++seen) {
    ASSIGN_OR_RETURN(const XlsbRecord rec, read_record(cursor));
    if (rec.type == ids::kEndSXFilters) {
      if (seen != declared || cursor.size != 0U) {
        return Corrupt("filter count does not match the filters decoded");
      }
      break;
    }
    if (rec.type != ids::kBeginSXFilter) {
      return Corrupt("filter list carries an uncharacterised record");
    }
    RETURN_IF_ERROR(AppendFilterXml(cursor, rec.payload, field_count, data_field_count, out));
  }
  out.append("</filters>");
  return out;
}

}  // namespace

Expected<pivot::PivotTable, Error> pivot_table_from_xlsb(ByteSpan part) {
  ASSIGN_OR_RETURN(const PivotTableDoc doc, decode_pivot_table(part));
  RETURN_IF_ERROR(CheckView(doc));
  const SxView& view = doc.view;
  pivot::PivotTable table;
  table.set_definition_source(pivot::PivotDefinitionSource::kXlsb);
  table.set_name(view.name);
  table.set_data_caption(view.data_caption);
  table.set_pivot_cache_id(view.cache_id);
  // `rowGrandTotals` totals the rows into a bottom line, the model's column totals.
  table.set_grand_totals(/*rows=*/(view.flags5 & kSxViewColGrand) != 0U,
                         /*cols=*/(view.flags5 & kSxViewRowGrand) != 0U);
  const bool root_compact = (view.flags8 & kViewCompactPair) != 0U;
  const bool root_outline = (view.flags8 & kViewOutlinePair) != 0U;
  table.set_layout(FormOf(root_compact, root_outline));
  table.set_grand_total_caption(view.grand_total_caption);
  table.set_auto_formatting((view.flags5 & kSxViewAutoFormat) != 0U);
  table.set_show_error((view.flags5 & kSxViewShowError) != 0U);
  table.set_error_caption(view.error_caption);
  table.set_missing_caption(view.missing_caption);
  table.set_show_headers((view.flags3 & kSxViewNoHeaders) == 0U);
  table.set_grid_drop_zones((view.flags8 & kSxViewNoDropZones) == 0U);
  table.set_item_print_titles(true);
  table.set_multiple_field_filters((view.flags9 & kSxViewSingleFieldFilter) == 0U);
  table.set_merge_item((view.flags6 & kSxViewMergeItem) != 0U);
  table.set_show_empty_row((view.flags4 & kViewShowEmptyRow) != 0U);
  table.set_show_empty_col((view.flags4 & kViewShowEmptyCol) != 0U);
  table.set_page_over_then_down((view.flags5 & kViewPageOverThenDown) != 0U);
  table.set_subtotal_hidden_items((view.flags5 & kViewSubtotalHiddenItems) != 0U);
  table.set_page_wrap(view.flags13);
  table.set_indent(static_cast<std::uint32_t>(view.flags3 & kSxViewIndentMask));
  ReadRootAttrs(doc, table.mutable_passthrough_attrs());
  RETURN_IF_ERROR(ReadLocation(doc.location, table));

  for (const SxField& in : doc.fields) {
    ASSIGN_OR_RETURN(pivot::PivotField field, ReadField(in));
    table.mutable_fields().push_back(std::move(field));
  }
  const std::size_t field_count = doc.fields.size();
  std::optional<std::size_t> row_values;
  RETURN_IF_ERROR(ReadFieldOrder(doc.row_fields, field_count, table.mutable_row_field_order(), row_values));
  table.set_row_values_position(row_values);
  // The first row field's own form decides the table's, as in the xlsx reader; the view's row outline bit is set
  // when any row field is in outline form.
  bool row_outline = false;
  if (!table.row_field_order().empty()) {
    const SxField& first = doc.fields[table.row_field_order().front()];
    table.set_layout(FormOf((first.name_flags & kSxFieldCompact) != 0U, (first.flags & kSxFieldOutline) != 0U));
    for (const std::uint32_t index : table.row_field_order()) {
      row_outline = row_outline || (doc.fields[index].flags & kSxFieldOutline) != 0U;
    }
  }
  if (((view.flags6 & kViewRowOutline) != 0U) != row_outline ||
      (view.data_axis == kDataOnRows) != row_values.has_value()) {
    return Unmeasured("row form or data axis disagreeing with the row fields");
  }
  CaptureForm(root_compact, root_outline, table.layout(), /*root=*/true, table.mutable_passthrough_attrs());
  if (view.reserved24 != 0U) {
    table.mutable_passthrough_attrs().emplace_back("chartFormat", std::to_string(view.reserved24));
  }
  for (std::size_t i = 0; i < field_count; ++i) {
    CaptureForm((doc.fields[i].name_flags & kSxFieldCompact) != 0U, (doc.fields[i].flags & kSxFieldOutline) != 0U,
                table.layout(), /*root=*/false, table.mutable_fields()[i].passthrough_attrs);
  }
  std::optional<std::size_t> col_values;
  RETURN_IF_ERROR(ReadFieldOrder(doc.col_fields, field_count, table.mutable_col_field_order(), col_values));
  table.set_col_values_position(col_values);
  table.mutable_raw_passthrough_after_row_fields() = LineItemsXml(doc.row_items, "rowItems");
  table.mutable_raw_col_items() = LineItemsXml(doc.col_items, "colItems");

  if (doc.page_fields) {
    for (const SxPageField& in : *doc.page_fields) {
      if (in.field >= field_count) {
        return Corrupt("page field names a field the table does not have");
      }
      pivot::PivotPageField entry;
      entry.field_index = in.field;
      entry.item_index = in.item == kNoPageItem ? std::nullopt : std::optional<std::uint32_t>(in.item);
      entry.hierarchy = static_cast<std::int32_t>(in.hierarchy);
      table.mutable_page_fields().push_back(std::move(entry));
    }
  }
  RETURN_IF_ERROR(ReadDataFields(doc, table));
  RETURN_IF_ERROR(CheckAxes(doc, table));
  RETURN_IF_ERROR(ReadStyle(doc.style, table));
  ReadFormats(doc, table);
  if (!doc.filters.empty()) {
    ASSIGN_OR_RETURN(const std::string filters, FiltersXml(doc.filters, field_count, doc.data_fields.size()));
    if (!read_pivot_filters_xml(filters, &table)) {
      return Corrupt("filters do not spell as OOXML");
    }
  }
  table.set_hide_values_row(doc.view14.has_value());
  table.mutable_raw_passthrough_xml() = TableExtLst(doc.view14.has_value(), doc.version_info.has_value());
  lift_pivot_sort_scopes(table);
  return table;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
