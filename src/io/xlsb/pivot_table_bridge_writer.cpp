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
#include "io/xlsb/pivot_table_bridge.h"
#include "io/xlsb/pivot_table_bridge_internal.h"
#include "io/xlsb/pivot_table_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xml_utils.h"
#include "pivot/filter_engine.h"
#include "pivot/pivot_types.h"
#include "pugixml.hpp"
#include "utils/a1_ref.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using namespace pivot_table_bridge_detail;

/// The table's own report-form attributes, when the xlsx reader kept an override.
struct FormOverride {
  std::optional<bool> compact;
  std::optional<bool> compact_data;
  std::optional<bool> outline;
  std::optional<bool> outline_data;
};

bool ParseFlag(std::string_view value, std::optional<bool>& out) {
  if (out.has_value() || (value != "0" && value != "1")) {
    return false;
  }
  out = value == "1";
  return true;
}

/// Reads the root passthrough attributes into `doc`; false for any the records cannot hold.
bool WriteRootAttrs(const pivot::PivotTable& table, PivotTableDoc& doc, FormOverride& form) {
  enum : unsigned {
    kCreated = 1U << 6U,
    kUpdated = 1U << 7U,
    kMinRefreshable = 1U << 8U,
    kNsMc = 1U << 9U,
    kNsXrDecl = 1U << 10U,
    kIgnorable = 1U << 11U,
    kUidSeen = 1U << 12U,
    kAsteriskTotals = 1U << 13U,
    kUidBookkeeping = kNsMc | kNsXrDecl | kIgnorable | kUidSeen,
  };
  unsigned seen = 0;
  // Each attribute absent from the part is its schema default, but `updatedVersion`
  // is the xlsx writer's own.
  doc.view.flags4 = 0;
  doc.view.flags7 = 0;
  doc.view.created_version = 0;
  doc.view.min_refreshable_version = 0;
  doc.view.updated_version = kBuiltTableUpdatedVersion;
  const auto take = [&seen](unsigned bit) {
    if ((seen & bit) != 0U) {
      return false;
    }
    seen |= bit;
    return true;
  };
  const auto version = [](std::string_view text, std::uint8_t& out) {
    const std::optional<std::uint32_t> v = ParseUnsigned(text, 0xFFU);
    out = static_cast<std::uint8_t>(v.value_or(0U));
    return v.has_value();
  };
  for (const auto& [name, value] : table.passthrough_attrs()) {
    bool ok = false;
    std::size_t apply = 0;
    while (apply < std::size(kApplyFormats) && name != kApplyFormats[apply]) {
      ++apply;
    }
    if (apply < std::size(kApplyFormats)) {
      ok = take(1U << apply) && (value == "0" || (apply == kApplyWidthHeight && value == "1"));
      if (ok && value == "1") {
        doc.view.flags7 = kViewApplyWidthHeight;
      }
    } else if (name == "xmlns:mc") {
      ok = take(kNsMc) && value == kNsMarkupCompat;
    } else if (name == "xmlns:xr") {
      ok = take(kNsXrDecl) && value == kNsXr;
    } else if (name == "mc:Ignorable") {
      ok = take(kIgnorable) && value == kXrPrefix;
    } else if (name == "xr:uid") {
      std::array<std::uint8_t, kGuidBytes> guid{};
      ok = take(kUidSeen) && ParseGuid(value, guid);
      doc.uid = guid;
    } else if (name == "createdVersion") {
      ok = take(kCreated) && version(value, doc.view.created_version);
    } else if (name == "updatedVersion") {
      ok = take(kUpdated) && version(value, doc.view.updated_version);
    } else if (name == "minRefreshableVersion") {
      ok = take(kMinRefreshable) && version(value, doc.view.min_refreshable_version);
    } else if (name == "asteriskTotals") {
      // Excel omits it at 0.
      ok = take(kAsteriskTotals) && value == "1";
      doc.view.flags4 = kViewAsteriskTotals;
    } else if (name == "compact") {
      ok = ParseFlag(value, form.compact);
    } else if (name == "compactData") {
      ok = ParseFlag(value, form.compact_data);
    } else if (name == "outline") {
      ok = ParseFlag(value, form.outline);
    } else if (name == "outlineData") {
      ok = ParseFlag(value, form.outline_data);
    } else if (name == "chartFormat") {
      // Excel omits `chartFormat` at 0.
      const std::optional<std::uint32_t> id = ParseUnsigned(value, 0xFFFFFFFFU);
      ok = doc.view.reserved24 == 0U && id.has_value() && *id != 0U;
      doc.view.reserved24 = id.value_or(0U);
    } else {
      // `dataOnRows` restates the Values position; no Excel part carries a `dataPosition`.
      ok = name == "dataOnRows" && value == Flag(table.row_values_position().has_value());
    }
    if (!ok) {
      return false;
    }
  }
  const unsigned bookkeeping = seen & kUidBookkeeping;
  return bookkeeping == (doc.uid ? kUidBookkeeping : 0U);
}

/// The Excel defaults the header bytes this bridge does not name stand for.
bool HasExcelDefaults(const pivot::PivotTable& t) {
  return t.enable_drill() && t.show_missing() && t.show_drill() && !t.print_drill() && t.show_data_tips() &&
         !t.field_print_titles() && t.item_print_titles() && t.show_data_drop_down() && t.preserve_formatting() &&
         t.page_wrap() <= 0xFFU && t.custom_list_sort() && t.indent() <= kSxViewIndentMask && t.alt_text().empty() &&
         t.alt_text_summary().empty();
}

/// The field's `compact` / `outline`: its own override, else the table's form.
bool FieldForm(const pivot::PivotField& field, pivot::PivotLayout layout, bool& compact, bool& outline) {
  std::optional<bool> own_compact;
  std::optional<bool> own_outline;
  for (const auto& [name, value] : field.passthrough_attrs) {
    if ((name == "compact" && !ParseFlag(value, own_compact)) ||
        (name == "outline" && !ParseFlag(value, own_outline))) {
      return false;
    }
  }
  if (own_compact.has_value() != own_outline.has_value()) {
    return false;
  }
  compact = own_compact.value_or(layout == pivot::PivotLayout::Compact);
  outline = own_outline.value_or(layout != pivot::PivotLayout::Tabular);
  return true;
}

/// A decimal `numFmtId`, or 0 when absent; false for a spelling the record cannot keep.
bool NumFmt(const std::string& text, std::uint32_t& out) {
  if (text.empty()) {
    out = 0;
    return true;
  }
  const std::optional<std::uint32_t> id = ParseUnsigned(text, 0xFFFFFFFFU);
  out = id.value_or(0U);
  return out != 0U;
}

bool WriteItems(const pivot::PivotField& field, std::optional<std::vector<SxItem>>& out) {
  if (field.items.empty() && field.marker_items.empty()) {
    return true;
  }
  std::vector<SxItem> items;
  for (std::size_t i = 0; i < field.items.size(); ++i) {
    const pivot::PivotItem& item = field.items[i];
    SxItem rec;
    rec.kind = kItemData;
    rec.index = item.has_cache_index ? item.cache_index : static_cast<std::uint32_t>(i);
    rec.flags =
        static_cast<std::uint16_t>((item.visible ? 0U : kSxItemHidden) | (item.show_detail ? 0U : kSxItemCollapsed) |
                                   (item.calculated ? kSxItemCalculated : 0U));
    for (const auto& [name, value] : item.passthrough_attrs) {
      if (name != "m" || value != "1" || (rec.flags & kSxItemMissing) != 0U) {
        return false;
      }
      rec.flags = static_cast<std::uint16_t>(rec.flags | kSxItemMissing);
    }
    if (!item.caption.empty() || item.blank_caption) {
      rec.caption = item.caption;
    }
    items.push_back(std::move(rec));
  }
  // The subtotal entries follow the field's selection as the xlsx writer emits
  // them; `grand` and `blank` have no item kind.
  const std::vector<pivot::SubtotalFn> fns = pivot::ordered_subtotal_fns(field);
  std::vector<std::uint8_t> expected;
  if (field.default_subtotal && fns.empty()) {
    expected.push_back(kItemDefault);
  }
  for (const pivot::SubtotalFn fn : fns) {
    expected.push_back(static_cast<std::uint8_t>(kItemFirstFn + static_cast<std::uint8_t>(fn)));
  }
  std::vector<std::uint8_t> authored;
  bool default_detail = true;
  for (const pivot::PivotMarkerItem& marker : field.marker_items) {
    const std::uint8_t kind = SubtotalItemKind(marker.type);
    if (kind == kItemData) {
      return false;
    }
    authored.push_back(kind);
    if (kind == kItemDefault) {
      default_detail = marker.show_detail;
    }
  }
  const bool keep = authored == expected;
  for (std::size_t k = 0; k < expected.size(); ++k) {
    const std::uint8_t kind = expected[k];
    bool detail = kind != kItemDefault || default_detail;
    if (keep) {
      const pivot::PivotMarkerItem& marker = field.marker_items[k];
      const std::string plain = "<item t=\"" + std::string(marker.type) + "\"/>";
      const bool collapsed_default = kind == kItemDefault && marker.xml == "<item t=\"default\" sd=\"0\"/>";
      if (marker.xml != plain && !collapsed_default) {
        return false;
      }
      detail = kind != kItemDefault || marker.show_detail;
    }
    SxItem rec;
    rec.kind = kind;
    rec.index = kNoItemIndex;
    rec.flags = static_cast<std::uint16_t>(detail ? 0U : kSxItemCollapsed);
    items.push_back(std::move(rec));
  }
  out = std::move(items);
  return true;
}

/// True when `node` carries no attribute outside `allowed`.
bool OnlyAttrs(const pugi::xml_node& node, std::initializer_list<std::string_view> allowed) {
  for (pugi::xml_attribute attr = node.first_attribute(); attr; attr = attr.next_attribute()) {
    if (std::find(allowed.begin(), allowed.end(), std::string_view(attr.name())) == allowed.end()) {
      return false;
    }
  }
  return true;
}

/// The element children of `node`, or nullopt when it holds text or anything else.
std::optional<std::vector<pugi::xml_node>> Elements(const pugi::xml_node& node) {
  std::vector<pugi::xml_node> out;
  for (pugi::xml_node child = node.first_child(); child; child = child.next_sibling()) {
    if (child.type() != pugi::node_element) {
      return std::nullopt;
    }
    out.push_back(child);
  }
  return out;
}

/// Parses `xml` as one element named `name`.
pugi::xml_node LoadOne(pugi::xml_document& doc, std::string_view xml, std::string_view name) {
  if (!doc.load_buffer(xml.data(), xml.size(), pugi::parse_default, pugi::encoding_utf8)) {
    return pugi::xml_node();
  }
  const pugi::xml_node root = doc.first_child();
  if (!root || root.next_sibling() || root.type() != pugi::node_element || std::string_view(root.name()) != name) {
    return pugi::xml_node();
  }
  return root;
}

bool WriteAutoSortScope(std::string_view xml, RawRecords& out) {
  if (xml.empty()) {
    return true;
  }
  pugi::xml_document doc;
  const pugi::xml_node scope = LoadOne(doc, xml, "autoSortScope");
  const std::optional<std::vector<pugi::xml_node>> areas = scope ? Elements(scope) : std::nullopt;
  if (!areas || areas->size() != 1U || scope.first_attribute()) {
    return false;
  }
  const pugi::xml_node area_node = areas->front();
  const std::optional<std::vector<pugi::xml_node>> refs_nodes = Elements(area_node);
  if (std::string_view(area_node.name()) != "pivotArea" ||
      !OnlyAttrs(area_node, {"dataOnly", "outline", "fieldPosition"}) || attr_str(area_node, "outline") != "0" ||
      attr_str(area_node, "fieldPosition") != "0" ||
      (area_node.attribute("dataOnly") && attr_str(area_node, "dataOnly") != "0") || !refs_nodes ||
      refs_nodes->size() != 1U || std::string_view(refs_nodes->front().name()) != "references") {
    return false;
  }
  pivot_area::PivotArea area;
  area.flags = area_node.attribute("dataOnly") ? 0U : pivot_area::kAreaDataOnly;
  const pugi::xml_node refs = refs_nodes->front();
  const std::optional<std::vector<pugi::xml_node>> ref_nodes = Elements(refs);
  if (!OnlyAttrs(refs, {"count"}) || !ref_nodes || ref_nodes->empty() ||
      attr_u32(refs, "count", 0U) != ref_nodes->size()) {
    return false;
  }
  std::vector<pivot_area::PivotAreaFilter> filters;
  for (const pugi::xml_node& ref : *ref_nodes) {
    const std::optional<std::vector<pugi::xml_node>> xs = Elements(ref);
    const std::optional<std::uint32_t> field = parse_xml_u32_attr_strict(ref.attribute("field"));
    if (std::string_view(ref.name()) != "reference" || !OnlyAttrs(ref, {"field", "count", "selected"}) || !field ||
        !xs || attr_u32(ref, "count", 0U) != xs->size() ||
        (ref.attribute("selected") && attr_str(ref, "selected") != "0")) {
      return false;
    }
    pivot_area::PivotAreaFilter filter;
    filter.field = *field;
    filter.flags = kAreaFilterAlways | (ref.attribute("selected") ? 0U : pivot_area::kFilterSelected);
    for (const pugi::xml_node& x : *xs) {
      if (std::string_view(x.name()) != "x" || !OnlyAttrs(x, {"v"}) || x.first_child()) {
        return false;
      }
      filter.items.push_back(attr_u32(x, "v", 0U));
    }
    filters.push_back(std::move(filter));
  }
  area.filters = std::move(filters);
  out.push_back(ids::RawRecord{ids::kBeginAutoSortScope, {}});
  pivot_area::encode_pivot_area(area, &out);
  out.push_back(ids::RawRecord{ids::kEndAutoSortScope, {}});
  return true;
}

/// A typed pivot area in the one shape Excel's parts show; false for any other.
bool WriteTableArea(const pivot::PivotTableArea& in, pivot_area::PivotArea& out) {
  if (in.outline || in.field_position != std::optional<std::uint32_t>(0U) ||
      (in.axis && *in.axis != pivot::PivotAxis::Row)) {
    return false;
  }
  const auto* type = std::find_if(std::begin(kAreaTypes), std::end(kAreaTypes),
                                  [&in](const AreaTypeCode& entry) { return entry.type == in.type; });
  if (type == std::end(kAreaTypes)) {
    return false;
  }
  out.field = in.field.value_or(pivot_area::kNoField);
  out.type = type->code;
  out.flags = (in.axis ? pivot_area::kAreaAxisRow : 0U) | (in.data_only ? pivot_area::kAreaDataOnly : 0U) |
              (in.label_only ? pivot_area::kAreaLabelOnly : 0U) | (in.grand_row ? pivot_area::kAreaGrandRow : 0U) |
              (in.grand_col ? pivot_area::kAreaGrandCol : 0U) |
              (in.collapsed_levels_are_subtotals ? pivot_area::kAreaCollapsedLevelsAreSubtotals : 0U);
  if (in.references.empty()) {
    out.filters.reset();
    return true;
  }
  std::vector<pivot_area::PivotAreaFilter> filters;
  for (const pivot::PivotTableAreaReference& ref : in.references) {
    if (!ref.field || ref.items.empty()) {
      return false;
    }
    pivot_area::PivotAreaFilter filter;
    filter.field = *ref.field;
    filter.flags = kAreaFilterAlways | (ref.selected ? pivot_area::kFilterSelected : 0U);
    filter.items = ref.items;
    filters.push_back(std::move(filter));
  }
  out.filters = std::move(filters);
  return true;
}

/// The typed formats and chart formats; a table holding both is unmeasured.
bool WriteFormats(const pivot::PivotTable& table, PivotTableDoc& doc) {
  if (!table.formats().empty() && !table.chart_formats().empty()) {
    return false;
  }
  for (const pivot::PivotFormat& in : table.formats()) {
    SxFormat out;
    if (in.action != pivot::PivotFormatAction::Formatting || !in.dxf_id || !WriteTableArea(in.area, out.area)) {
      return false;
    }
    out.dxf_id = *in.dxf_id;
    doc.formats.push_back(std::move(out));
  }
  for (const pivot::PivotChartFormat& in : table.chart_formats()) {
    SxChartFormat out;
    if (in.chart != 0U || !in.series || !WriteTableArea(in.area, out.area)) {
      return false;
    }
    out.format = in.format;
    doc.chart_formats.push_back(std::move(out));
  }
  return true;
}

/// Reads the table extLst into whether SXView14 is written: Excel's two
/// `<ext>` elements as the reader keeps them, the SXView14 one optional, plus
/// any extension of a URI Office does not define, which Excel's xlsb save drops. The RichData version-info
/// extension after them sets `rich_version_info`.
bool WriteTableExtLst(std::string_view xml, bool& view14, bool& rich_version_info) {
  pugi::xml_document doc;
  const pugi::xml_node ext_lst = LoadOne(doc, xml, "extLst");
  const std::optional<std::vector<pugi::xml_node>> exts = ext_lst ? Elements(ext_lst) : std::nullopt;
  if (!exts || ext_lst.first_attribute()) {
    return false;
  }
  view14 = false;
  rich_version_info = false;
  bool view16 = false;
  bool view14_off_seen = false;
  for (const pugi::xml_node& ext : *exts) {
    const std::string_view uri = ext.attribute("uri").as_string();
    if (std::string_view(ext.name()) != "ext") {
      return false;
    }
    if (uri == kView14Uri || uri == kView16Uri) {
      // The SXView14 extension precedes the SXView16 one, each at most once.
      const bool is14 = uri == kView14Uri;
      const std::string ext_xml = raw_xml(ext);
      const bool view14_off = is14 && ext_xml == kView14OffExt;
      if (view16 || (is14 && (view14 || view14_off_seen)) ||
          (ext_xml != (is14 ? kView14Ext : kView16Ext) && !view14_off)) {
        return false;
      }
      if (view14_off) {
        view14_off_seen = true;
      } else {
        (is14 ? view14 : view16) = true;
      }
      continue;
    }
    if (uri == kVersionInfoUri) {
      if (!view16 || rich_version_info || raw_xml(ext) != kVersionInfoExt) {
        return false;
      }
      rich_version_info = true;
      continue;
    }
    if (uri.empty() || std::find(std::begin(kOfficeTableExtUris), std::end(kOfficeTableExtUris), uri) !=
                           std::end(kOfficeTableExtUris)) {
      return false;
    }
  }
  return view16;
}

bool WriteField(const pivot::PivotTable& table, std::size_t index, ExcelHost host, SxField& out) {
  const pivot::PivotField& field = table.fields()[index];
  // Of the x14 field extension only `fillDownLabels` has a record.
  if (!field.x14_attrs.empty() || !field.x14_children_xml.empty() || !field.raw_ext_xml.empty()) {
    return false;
  }
  bool used_as_data = false;
  for (const pivot::PivotDataField& df : table.data_fields()) {
    used_as_data = used_as_data || df.field_index == index;
  }
  const bool measure_filter =
      std::any_of(table.active_filters().begin(), table.active_filters().end(), [&](const pivot::PivotFilter& f) {
        return pivot::is_value_filter_type(f.type) && !detail::ReEmitsSource(table, f) &&
               detail::FilterFieldIndex(table, f) == index;
      });
  bool compact = false;
  bool outline = false;
  if (!FieldForm(field, table.layout(), compact, outline)) {
    return false;
  }
  bool drag_off = false;
  bool row_drag_off = false;
  bool measure_attr = false;
  bool include_new_items = false;
  unsigned drag_seen = 0;
  for (const auto& [name, value] : field.passthrough_attrs) {
    if (name == "dragToRow") {
      if (value != "0" || row_drag_off) {
        return false;
      }
      row_drag_off = true;
    } else if (name == "dragToCol" || name == "dragToPage") {
      const unsigned bit = name == "dragToCol" ? 1U : 2U;
      if (value != "0" || (drag_seen & bit) != 0U) {
        return false;
      }
      drag_seen |= bit;
    } else if (name == "measureFilter") {
      if (value != "1" || measure_attr) {
        return false;
      }
      measure_attr = true;
    } else if (name == "includeNewItemsInFilter") {
      if (value != "1" || include_new_items) {
        return false;
      }
      include_new_items = true;
    } else if (name != "compact" && name != "outline") {
      return false;
    }
  }
  if (drag_seen != 0U && drag_seen != 3U) {
    return false;
  }
  drag_off = drag_seen == 3U;
  switch (field.axis) {
    case pivot::PivotAxis::Row:
      out.axis = kSxFieldAxisRow;
      break;
    case pivot::PivotAxis::Col:
      out.axis = kSxFieldAxisCol;
      break;
    case pivot::PivotAxis::Page:
      out.axis = kSxFieldAxisPage;
      break;
    default:
      out.axis = 0;
      break;
  }
  if (used_as_data) {
    out.axis = static_cast<std::uint8_t>(out.axis | kSxFieldAxisData);
  }
  const std::vector<pivot::SubtotalFn> fns = pivot::ordered_subtotal_fns(field);
  out.subtotals = static_cast<std::uint16_t>(field.default_subtotal && fns.empty() ? kSxFieldSubtotalDefault : 0U);
  for (const pivot::SubtotalFn fn : fns) {
    out.subtotals = static_cast<std::uint16_t>(out.subtotals | (1U << (static_cast<unsigned>(fn) + 1U)));
  }
  out.name_flags = static_cast<std::uint8_t>(compact ? kSxFieldCompact : 0U);
  if (!NumFmt(field.number_format, out.num_fmt)) {
    return false;
  }
  const bool on_axis = field.axis == pivot::PivotAxis::Row || field.axis == pivot::PivotAxis::Col ||
                       field.axis == pivot::PivotAxis::Page;
  const bool auto_sort = (on_axis || field.sort.stated) && !field.sort.manual;
  if (include_new_items && (auto_sort || !field.sort.ascending)) {
    return false;
  }
  out.flags =
      kFieldFixed | (include_new_items ? 0U : kFieldNewItemsOutsideFilter) | (row_drag_off ? 0U : kFieldDragToRow) |
      (drag_off ? 0U : kFieldDragPair) | (field.show_all ? kSxFieldShowAll : 0U) | (outline ? kSxFieldOutline : 0U) |
      (field.insert_blank_row ? kSxFieldInsertBlankRow : 0U) | (field.subtotal_top ? kSxFieldSubtotalTop : 0U) |
      (auto_sort ? kSxFieldAutoSort : 0U) |
      ((auto_sort || include_new_items || host == ExcelHost::kWin365) && field.sort.ascending ? kSxFieldAscendSort
                                                                                              : 0U) |
      (measure_attr || measure_filter ? kSxFieldMeasureFilter : 0U);
  out.item_page_count = kItemPageCount;
  out.reserved16 = kFieldReserved16;
  if (!field.custom_name.empty()) {
    out.display_name = field.custom_name;
  }
  if (!WriteItems(field, out.items) ||
      !WriteAutoSortScope(pivot_field_trailing_children(table, field), out.auto_sort_scope)) {
    return false;
  }
  // SXVD14 is measured after the items of a field without a sort scope.
  if (field.fill_down_labels && (!out.items || !out.auto_sort_scope.empty())) {
    return false;
  }
  out.fill_down_labels = field.fill_down_labels;
  return true;
}

/// The axis list of `order` with the Values pseudo-field at `values_position`, as the xlsx writer spells it.
bool WriteFieldOrder(const std::vector<std::uint32_t>& order, std::optional<std::size_t> values_position,
                     std::size_t field_count, std::optional<std::vector<std::uint32_t>>& out) {
  if (order.empty() && !values_position) {
    return true;
  }
  if (values_position && *values_position > order.size()) {
    return false;
  }
  std::vector<std::uint32_t> list;
  for (std::size_t i = 0; i <= order.size(); ++i) {
    if (values_position && *values_position == i) {
      list.push_back(kValuesField);
    }
    if (i < order.size()) {
      if (order[i] >= field_count) {
        return false;
      }
      list.push_back(order[i]);
    }
  }
  out = std::move(list);
  return true;
}

bool WriteLineItems(std::string_view xml, std::string_view tag, SxLineItems& out) {
  if (xml.empty()) {
    return true;
  }
  pugi::xml_document doc;
  const pugi::xml_node root = LoadOne(doc, xml, tag);
  const std::optional<std::vector<pugi::xml_node>> lines = root ? Elements(root) : std::nullopt;
  if (!lines || !OnlyAttrs(root, {"count"}) || attr_u32(root, "count", 0U) != lines->size()) {
    return false;
  }
  for (const pugi::xml_node& node : *lines) {
    const std::optional<std::vector<pugi::xml_node>> xs = Elements(node);
    if (std::string_view(node.name()) != "i" || !OnlyAttrs(node, {"t", "r", "i"}) || !xs) {
      return false;
    }
    SxLineItem line;
    if (const pugi::xml_attribute t = node.attribute("t"); t) {
      const auto* it = std::find(std::begin(kLineItemTypes), std::end(kLineItemTypes), std::string_view(t.value()));
      if (it == std::end(kLineItemTypes)) {
        return false;
      }
      line.t = static_cast<std::uint8_t>(it - std::begin(kLineItemTypes));
    }
    const std::uint32_t r = attr_u32(node, "r", 0U);
    if (r > 0xFFFFU) {
      return false;
    }
    line.r = static_cast<std::uint16_t>(r);
    line.i = attr_u32(node, "i", 0U);
    for (const pugi::xml_node& x : *xs) {
      if (std::string_view(x.name()) != "x" || !OnlyAttrs(x, {"v"}) || x.first_child()) {
        return false;
      }
      line.x.push_back(attr_u32(x, "v", 0U));
    }
    out.items.push_back(std::move(line));
  }
  return true;
}

bool WritePages(const pivot::PivotTable& table, PivotTableDoc& doc) {
  const std::vector<std::uint32_t> order = table.page_field_order();
  if (order.empty()) {
    return true;
  }
  std::vector<SxPageField> pages;
  for (const std::uint32_t index : order) {
    const pivot::PivotPageField* entry = table.find_page_field(index);
    const std::optional<std::int32_t> hierarchy = entry != nullptr ? entry->hierarchy : std::optional<std::int32_t>(-1);
    if (!hierarchy || (entry != nullptr && (!entry->name.empty() || !entry->caption.empty()))) {
      return false;
    }
    SxPageField page;
    page.field = index;
    page.item = entry != nullptr && entry->item_index ? *entry->item_index : kNoPageItem;
    page.hierarchy = static_cast<std::uint32_t>(*hierarchy);
    pages.push_back(page);
  }
  doc.page_fields = std::move(pages);
  return true;
}

bool WriteDataFields(const pivot::PivotTable& table, PivotTableDoc& doc) {
  for (const pivot::PivotDataField& in : table.data_fields()) {
    SxDataField out;
    // An x14 display mode is SXDI14 beside a plain SXDI.
    const auto* x14 = std::find(std::begin(kPivotShowAs), std::end(kPivotShowAs), in.show_as);
    const bool is_x14 = x14 != std::end(kPivotShowAs);
    const pivot::ShowValuesAs show_as =
        is_x14 ? pivot::ShowValuesAs::Normal
               : (in.show_as == pivot::ShowValuesAs::RunningTotalInCol ? pivot::ShowValuesAs::RunningTotalInRow
                                                                       : in.show_as);
    const auto* code = std::find(std::begin(kShowDataAs), std::end(kShowDataAs), show_as);
    if (in.field_index >= table.fields().size() || code == std::end(kShowDataAs) ||
        !NumFmt(in.number_format, out.num_fmt) || !in.raw_ext_xml.empty() || !in.x14_attrs.empty() ||
        !in.x14_children_xml.empty()) {
      return false;
    }
    out.field = in.field_index;
    out.subtotal = static_cast<std::uint32_t>(kPivotAggregationNames.size());
    for (std::size_t k = 0; k < kPivotAggregationNames.size(); ++k) {
      if (kPivotAggregationNames[k].value == static_cast<std::uint8_t>(in.aggregation)) {
        out.subtotal = static_cast<std::uint32_t>(k);
        break;
      }
    }
    if (out.subtotal == kPivotAggregationNames.size()) {
      return false;
    }
    out.show_data_as = static_cast<std::uint32_t>(code - std::begin(kShowDataAs));
    out.base_field = in.show_as_base_field.value_or(0U);
    out.base_item = in.show_as_base_item.value_or(0U);
    out.name = in.name;
    if (is_x14) {
      out.pivot_show_as = kSxDataShowAsPercentOfParentRow + static_cast<std::uint32_t>(x14 - std::begin(kPivotShowAs));
    }
    doc.data_fields.push_back(std::move(out));
  }
  return true;
}

/// Splits framed records into raw records; false when they do not frame whole.
bool SplitRecords(const std::vector<std::uint8_t>& bytes, RawRecords& out) {
  ByteSpan cursor{bytes.data(), bytes.size()};
  while (cursor.size > 0U) {
    auto rec = read_record(cursor);
    if (!rec) {
      return false;
    }
    const ByteSpan payload = rec.value().payload;
    out.push_back(
        ids::RawRecord{rec.value().type, std::vector<std::uint8_t>(payload.data, payload.data + payload.size)});
  }
  return true;
}

/// One `<filter>` as its SXFILTER block.
bool WriteFilter(const pugi::xml_node& node, RawRecords& out) {
  const std::optional<std::vector<pugi::xml_node>> children = Elements(node);
  const std::optional<std::uint32_t> fld = parse_xml_u32_attr_strict(node.attribute("fld"));
  const std::optional<std::uint32_t> id = parse_xml_u32_attr_strict(node.attribute("id"));
  const auto* type = std::find(std::begin(kPivotFilterTypes) + 1, std::end(kPivotFilterTypes),
                               std::string_view(attr_str(node, "type")));
  if (std::string_view(node.name()) != "filter" || !fld || !id || type == std::end(kPivotFilterTypes) ||
      !OnlyAttrs(node, {"fld", "mpFld", "type", "evalOrder", "id", "iMeasureHier", "iMeasureFld", "name", "description",
                        "stringValue1", "stringValue2"}) ||
      !children || children->size() != 1U || std::string_view(children->front().name()) != "autoFilter") {
    return false;
  }
  const auto index_attr = [&node](const char* name) {
    const pugi::xml_attribute attr = node.attribute(name);
    return attr ? parse_xml_u32_attr_strict(attr) : std::optional<std::uint32_t>(kNoFilterIndex);
  };
  const std::optional<std::uint32_t> mp_fld = index_attr("mpFld");
  const std::optional<std::uint32_t> measure_fld = index_attr("iMeasureFld");
  const std::optional<std::uint32_t> measure_hier = index_attr("iMeasureHier");
  if (!mp_fld || !measure_fld || !measure_hier) {
    return false;
  }
  std::vector<std::uint8_t> p;
  emit_u32(p, *fld);
  emit_u32(p, *mp_fld);
  emit_u32(p, static_cast<std::uint32_t>(type - std::begin(kPivotFilterTypes)));
  emit_u32(p, static_cast<std::uint32_t>(attr_i32(node, "evalOrder", 0)));
  emit_u32(p, *id);
  emit_u32(p, *measure_fld);
  emit_u32(p, *measure_hier);
  std::uint16_t flags = 0;
  for (std::size_t bit = 0; bit < std::size(kFilterStringAttrs); ++bit) {
    if (node.attribute(std::string(kFilterStringAttrs[bit]).c_str())) {
      flags = static_cast<std::uint16_t>(flags | (1U << bit));
    }
  }
  emit_u16(p, flags);
  for (const std::string_view name : kFilterStringAttrs) {
    if (const pugi::xml_attribute attr = node.attribute(std::string(name).c_str()); attr) {
      emit_xlwidestring(p, attr.value());
    }
  }
  // The xlsx load rule's comparison state is a caption filter's: = and <> match as a pattern, ordering is literal.
  auto criteria = parse_auto_filter_xml(raw_xml(children->front()));
  if (!criteria) {
    return false;
  }
  std::vector<std::uint8_t> afilter;
  if (emit_auto_filter_block(afilter, criteria.value()) != 0U) {
    return false;
  }
  out.push_back(ids::RawRecord{ids::kBeginSXFilter, std::move(p)});
  out.push_back(
      ids::RawRecord{ids::kACBegin, std::vector<std::uint8_t>(kAcUidBeginPayload.begin(), kAcUidBeginPayload.end())});
  out.push_back(ids::RawRecord{ids::kUid, std::vector<std::uint8_t>(kGuidBytes, 0U)});
  out.push_back(ids::RawRecord{ids::kACEnd, {}});
  if (!SplitRecords(afilter, out)) {
    return false;
  }
  out.push_back(ids::RawRecord{ids::kEndSXFilter, {}});
  return true;
}

/// `<filters>` as the xlsx writer would spell it, encoded as SXFILTERS.
bool WriteFilters(const pivot::PivotTable& table, RawRecords& out) {
  std::string xml;
  detail::AppendFilters(xml, table);
  if (xml.empty()) {
    return true;
  }
  pugi::xml_document doc;
  const pugi::xml_node root = LoadOne(doc, xml, "filters");
  const std::optional<std::vector<pugi::xml_node>> filters = root ? Elements(root) : std::nullopt;
  if (!filters) {
    return false;
  }
  std::vector<std::uint8_t> count;
  emit_u32(count, static_cast<std::uint32_t>(filters->size()));
  out.push_back(ids::RawRecord{ids::kBeginSXFilters, std::move(count)});
  for (const pugi::xml_node& node : *filters) {
    if (!WriteFilter(node, out)) {
      return false;
    }
  }
  out.push_back(ids::RawRecord{ids::kEndSXFilters, {}});
  return true;
}

bool WriteLocation(const pivot::PivotTable& table, std::optional<PivotRenderedSpan> rendered_span, SxLocation& loc) {
  // A withheld report is saved as its anchor cell alone, with zero offsets and no page rows or columns.
  if (table.blocked()) {
    if (table.anchor_row() >= a1::kMaxRows || table.anchor_col() >= a1::kMaxCols) {
      return false;
    }
    loc = SxLocation{};
    loc.first_row = loc.last_row = loc.first_header_row = loc.first_data_row = table.anchor_row();
    loc.first_col = loc.last_col = loc.first_data_col = table.anchor_col();
    return true;
  }
  // An offset the model does not carry is derived from the report shape, as the .xlsx writer derives it.
  const detail::LocationDefaults derived = detail::ComputeLocationDefaults(table);
  const std::uint32_t header = table.location_first_header_row().value_or(derived.first_header_row);
  const std::uint32_t data_row = table.location_first_data_row().value_or(derived.first_data_row);
  const std::uint32_t data_col = table.location_first_data_col().value_or(derived.first_data_col);
  const std::optional<std::uint32_t> row_pages =
      table.location_row_page_count() ? table.location_row_page_count() : derived.row_page_count;
  const std::optional<std::uint32_t> col_pages =
      table.location_col_page_count() ? table.location_col_page_count() : derived.col_page_count;
  // The span rule of `write_pivot_table_definition`: a stale read span is
  // replaced by the projection, a caller's span is a floor under it.
  std::uint32_t span_rows = table.span_rows();
  std::uint32_t span_cols = table.span_cols();
  std::uint32_t top = table.anchor_row();
  if (!table.has_authored_span()) {
    if (!rendered_span) {
      return false;
    }
    const bool stale = table.span_read();
    span_rows = !stale && span_rows > rendered_span->rows ? span_rows : rendered_span->rows;
    span_cols = !stale && span_cols > rendered_span->cols ? span_cols : rendered_span->cols;
    top = top > rendered_span->top ? top : rendered_span->top;
  }
  if (span_rows == 0U || span_cols == 0U || top >= a1::kMaxRows || table.anchor_col() >= a1::kMaxCols ||
      span_rows > a1::kMaxRows - top || span_cols > a1::kMaxCols - table.anchor_col() || header >= a1::kMaxRows ||
      data_row >= a1::kMaxRows || data_col >= a1::kMaxCols || row_pages == 0U || col_pages == 0U) {
    return false;
  }
  loc.first_row = top;
  loc.last_row = top + span_rows - 1U;
  loc.first_col = table.anchor_col();
  loc.last_col = table.anchor_col() + span_cols - 1U;
  loc.first_header_row = top + header;
  loc.first_data_row = top + data_row;
  loc.first_data_col = table.anchor_col() + data_col;
  loc.row_page_count = row_pages.value_or(0U);
  loc.col_page_count = col_pages.value_or(0U);
  return true;
}

/// Fills `doc` from `table`; false for a state the records cannot hold or no Excel part shows.
bool WriteTable(const pivot::PivotTable& table, ExcelHost host, std::optional<PivotRenderedSpan> rendered_span,
                const PivotItemsXml* items, PivotTableDoc& doc) {
  if ((table.layout_dirty() && items == nullptr) || !HasExcelDefaults(table) || !table.style_info() ||
      static_cast<int>(table.grand_total_caption().has_value()) + static_cast<int>(table.error_caption().has_value()) +
              static_cast<int>(table.missing_caption().has_value()) >
          1) {
    return false;
  }
  // Untyped `<formats>` and its siblings have no records here; the table
  // extLst holds SXView14 / SXView16 and extensions Excel's xlsb save drops.
  // A table with no captured extLst takes the xlsx writer's; Excel adds SXView16 to it.
  bool view14 = table_hides_values_row(table);
  bool rich_version_info = false;
  if (!table.raw_passthrough_before_style_info().empty() ||
      (!table.raw_passthrough_xml().empty() &&
       (!WriteTableExtLst(table.raw_passthrough_xml(), view14, rich_version_info) ||
        table.hide_values_row().value_or(view14) != view14)) ||
      !WriteFormats(table, doc)) {
    return false;
  }
  FormOverride form;
  if (!WriteRootAttrs(table, doc, form)) {
    return false;
  }
  const bool form_override = form.compact || form.compact_data || form.outline || form.outline_data;
  if (form_override && (!form.compact || !form.compact_data || !form.outline || !form.outline_data ||
                        *form.compact != *form.compact_data || *form.outline != *form.outline_data)) {
    return false;
  }
  const bool compact = form_override ? *form.compact : table.layout() == pivot::PivotLayout::Compact;
  const bool outline = form_override ? *form.outline : table.layout() != pivot::PivotLayout::Tabular;
  // The view's row outline bit is set when any row field is in outline form.
  bool row_outline = false;
  for (const std::uint32_t index : table.row_field_order()) {
    bool field_compact = false;
    bool field_outline = false;
    if (index >= table.fields().size() ||
        !FieldForm(table.fields()[index], table.layout(), field_compact, field_outline)) {
      return false;
    }
    row_outline = row_outline || field_outline;
  }
  if (table.row_field_order().empty() && table.row_values_position()) {
    return false;
  }

  SxView& v = doc.view;
  v.flags1 = kViewFlags1;
  v.flags2 = kViewFlags2;
  v.flags3 = static_cast<std::uint8_t>(table.indent() | (table.show_headers() ? 0U : kSxViewNoHeaders));
  v.flags4 = static_cast<std::uint8_t>(kViewFlags4 | (v.flags4 & kViewAsteriskTotals) |
                                       (table.show_empty_row() ? kViewShowEmptyRow : 0U) |
                                       (table.show_empty_col() ? kViewShowEmptyCol : 0U));
  v.flags5 = static_cast<std::uint8_t>(
      kViewFlags5Fixed | (table.auto_formatting() ? kSxViewAutoFormat : 0U) |
      (table.show_error() ? kSxViewShowError : 0U) | (table.grand_totals_cols() ? kSxViewRowGrand : 0U) |
      (table.grand_totals_rows() ? kSxViewColGrand : 0U) | (table.page_over_then_down() ? kViewPageOverThenDown : 0U) |
      (table.subtotal_hidden_items() ? kViewSubtotalHiddenItems : 0U));
  v.flags13 = static_cast<std::uint8_t>(table.page_wrap());
  v.flags6 = static_cast<std::uint8_t>(kViewFlags6Fixed | (table.merge_item() ? kSxViewMergeItem : 0U) |
                                       (row_outline ? kViewRowOutline : 0U));
  v.flags8 = static_cast<std::uint8_t>((compact ? kViewCompactPair : 0U) | (outline ? kViewOutlinePair : 0U) |
                                       (table.grid_drop_zones() ? 0U : kSxViewNoDropZones));
  v.flags9 = static_cast<std::uint8_t>(table.multiple_field_filters() ? 0U : kSxViewSingleFieldFilter);
  v.data_axis = table.row_values_position() ? kDataOnRows : kDataOnCols;
  v.data_position = kNoDataPosition;
  v.reserved20 = kViewReserved20;
  v.cache_id = table.pivot_cache_id();
  v.name = table.name();
  v.data_caption = table.data_caption();
  v.grand_total_caption = table.grand_total_caption();
  v.error_caption = table.error_caption();
  v.missing_caption = table.missing_caption();

  if (!WriteLocation(table, rendered_span, doc.location)) {
    return false;
  }
  doc.fields.resize(table.fields().size());
  for (std::size_t i = 0; i < table.fields().size(); ++i) {
    if (!WriteField(table, i, host, doc.fields[i])) {
      return false;
    }
  }
  const std::size_t field_count = table.fields().size();
  if (!WriteFieldOrder(table.row_field_order(), table.row_values_position(), field_count, doc.row_fields) ||
      !WriteFieldOrder(table.col_field_order(), table.col_values_position(), field_count, doc.col_fields) ||
      !WriteLineItems(items != nullptr ? items->row_items : table.raw_passthrough_after_row_fields(), "rowItems",
                      doc.row_items) ||
      !WriteLineItems(items != nullptr ? items->col_items : table.raw_col_items(), "colItems", doc.col_items) ||
      !WritePages(table, doc) || !WriteDataFields(table, doc) || !WriteFilters(table, doc.filters)) {
    return false;
  }
  const pivot::PivotStyleInfo& style = *table.style_info();
  doc.style.flags = static_cast<std::uint16_t>(
      (style.show_last_column ? kStyleLastColumn : 0U) | (style.show_row_stripes ? kStyleRowStripes : 0U) |
      (style.show_col_stripes ? kStyleColStripes : 0U) | (style.show_row_headers ? kStyleRowHeaders : 0U) |
      (style.show_col_headers ? kStyleColHeaders : 0U) | (host == ExcelHost::kWin365 ? kStyleFirstColumn : 0U));
  if (!style.name.empty()) {
    doc.style.name = style.name;
  }
  if (view14) {
    doc.view14 = SxFrtBlock{{kView14Frt.begin(), kView14Frt.end()}, {kView14Payload.begin(), kView14Payload.end()}};
  }
  doc.view16.frt_payload.assign(kView16Frt.begin(), kView16Frt.end());
  doc.view16.payload.assign(kView16Payload.begin(), kView16Payload.end());
  if (rich_version_info) {
    doc.version_info =
        SxFrtBlock{{pivot_records::kVersionInfoFrtVersion.begin(), pivot_records::kVersionInfoFrtVersion.end()},
                   {pivot_records::kRichDataFeature.begin(), pivot_records::kRichDataFeature.end()}};
  }
  return true;
}
}  // namespace

Expected<std::vector<std::uint8_t>, Error> pivot_table_to_xlsb(const pivot::PivotTable& table, ExcelHost host,
                                                               std::size_t sheet_index,
                                                               std::optional<PivotRenderedSpan> rendered_span,
                                                               const PivotItemsXml* items) {
  PivotTableDoc doc;
  if (!WriteTable(table, host, rendered_span, items, doc)) {
    return not_writable(kNotWritableKind, sheet_index);
  }
  return encode_pivot_table(doc);
}
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
