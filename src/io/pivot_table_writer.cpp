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

using AttrList = std::vector<std::pair<std::string, std::string>>;

/// `<pivotField>` attribute names in CT_PivotField order.
constexpr std::string_view kPivotFieldAttrOrder[] = {
    "name",
    "axis",
    "dataField",
    "subtotalCaption",
    "showDropDowns",
    "hiddenLevel",
    "uniqueMemberProperty",
    "compact",
    "allDrilled",
    "numFmtId",
    "outline",
    "subtotalTop",
    "dragToRow",
    "dragToCol",
    "multipleItemSelectionAllowed",
    "dragToPage",
    "dragToData",
    "dragOff",
    "showAll",
    "insertBlankRow",
    "serverField",
    "insertPageBreak",
    "autoShow",
    "topAutoShow",
    "hideNewItems",
    "measureFilter",
    "includeNewItemsInFilter",
    "itemPageCount",
    "sortType",
    "dataSourceSort",
    "nonAutoSortDefault",
    "rankBy",
    "defaultSubtotal",
    "sumSubtotal",
    "countASubtotal",
    "avgSubtotal",
    "maxSubtotal",
    "minSubtotal",
    "productSubtotal",
    "countSubtotal",
    "stdDevSubtotal",
    "stdDevPSubtotal",
    "varSubtotal",
    "varPSubtotal",
    "showPropCell",
    "showPropTip",
    "showPropAsCaption",
    "defaultAttributeDrillState",
};

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

/// Appends `attrs` in the order of `order`; names outside it follow in their
/// given order.
template <std::size_t N>
void AppendAttrsInOrder(std::string& out, const AttrList& attrs, const std::string_view (&order)[N]) {
  AttrList sorted;
  sorted.reserve(attrs.size());
  std::vector<bool> taken(attrs.size(), false);
  for (const std::string_view name : order) {
    for (std::size_t i = 0; i < attrs.size(); ++i) {
      if (!taken[i] && attrs[i].first == name) {
        taken[i] = true;
        sorted.push_back(attrs[i]);
      }
    }
  }
  for (std::size_t i = 0; i < attrs.size(); ++i) {
    if (!taken[i]) {
      sorted.push_back(attrs[i]);
    }
  }
  append_raw_attrs(out, sorted);
}

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

/// Where the start tag at the head of an XML fragment spells one attribute.
struct StartTagAttr {
  /// The attribute with its leading whitespace, `[begin, end)`; `begin` is
  /// npos when absent.
  std::size_t begin = std::string_view::npos;
  std::size_t end = 0;
  /// Position of the `/` or `>` closing the start tag.
  std::size_t tag_end = std::string_view::npos;
  std::string_view value;
};

/// Scans the attributes of the start tag in `xml`, quoted either way.
StartTagAttr FindStartTagAttr(std::string_view xml, std::string_view attr) {
  constexpr std::string_view kSpace = " \t\r\n";
  StartTagAttr found;
  std::size_t pos = xml.find_first_of(kSpace);
  while (pos < xml.size()) {
    const std::size_t name_begin = xml.find_first_not_of(kSpace, pos);
    if (name_begin == std::string_view::npos || xml[name_begin] == '/' || xml[name_begin] == '>') {
      found.tag_end = name_begin;
      break;
    }
    const std::size_t eq = xml.find('=', name_begin);
    const std::size_t open = eq == std::string_view::npos ? eq : xml.find_first_not_of(kSpace, eq + 1);
    if (open == std::string_view::npos || (xml[open] != '"' && xml[open] != '\'')) {
      break;
    }
    const std::size_t close = xml.find(xml[open], open + 1);
    if (close == std::string_view::npos) {
      break;
    }
    std::string_view name = xml.substr(name_begin, eq - name_begin);
    name = name.substr(0, name.find_last_not_of(kSpace) + 1);
    if (name == attr) {
      found.begin = pos;
      found.end = close + 1;
      found.value = xml.substr(open + 1, close - open - 1);
    }
    pos = close + 1;
  }
  return found;
}

/// The boolean an xsd:boolean attribute spells, `fallback` when absent.
bool StartTagFlag(std::string_view xml, std::string_view attr, bool fallback) {
  const StartTagAttr found = FindStartTagAttr(xml, attr);
  if (found.begin == std::string_view::npos) {
    return fallback;
  }
  return found.value != "0" && found.value != "false";
}

/// `xml` with the boolean attribute `attr` of its head start tag set to
/// `value`: written when it differs from the schema default, removed when it
/// equals it. Every other attribute stays as captured.
std::string WithStartTagFlag(std::string_view xml, std::string_view attr, bool value, bool schema_default) {
  const StartTagAttr found = FindStartTagAttr(xml, attr);
  std::string spelled;
  if (value != schema_default) {
    spelled.push_back(' ');
    spelled.append(attr);
    spelled.append(value ? "=\"1\"" : "=\"0\"");
  }
  std::string out(xml);
  if (found.begin != std::string_view::npos) {
    out.replace(found.begin, found.end - found.begin, spelled);
  } else if (!spelled.empty() && found.tag_end != std::string_view::npos) {
    out.insert(found.tag_end, spelled);
  }
  return out;
}

/// `raw_ext_xml` with `fillDownLabels` on the `<pivotField>` element of the
/// x14 field extension set to `fill_down`; unchanged when that extension or
/// element is absent.
std::string WithVerbatimFillDownLabels(const std::string& raw_ext_xml, bool fill_down) {
  const std::string_view ext_uri(kPivotFieldExtUri);
  std::size_t uri = std::string::npos;
  for (std::size_t at = 0; at + ext_uri.size() <= raw_ext_xml.size(); ++at) {
    if (strings::case_insensitive_eq(std::string_view(raw_ext_xml).substr(at, ext_uri.size()), ext_uri)) {
      uri = at;
      break;
    }
  }
  if (uri == std::string::npos) {
    return raw_ext_xml;
  }
  for (std::size_t lt = raw_ext_xml.find('<', uri); lt != std::string::npos; lt = raw_ext_xml.find('<', lt + 1)) {
    const std::size_t name_end = raw_ext_xml.find_first_of(" \t\r\n/>", lt + 1);
    if (name_end == std::string::npos) {
      break;
    }
    const std::string_view name(raw_ext_xml.data() + lt + 1, name_end - lt - 1);
    if (name == "/ext" || (!name.empty() && name.front() == '/' && strings::ends_with(name, ":ext"))) {
      break;
    }
    const std::size_t colon = name.find(':');
    if ((colon == std::string_view::npos ? name : name.substr(colon + 1)) != "pivotField") {
      continue;
    }
    const std::string_view tail(raw_ext_xml.data() + lt, raw_ext_xml.size() - lt);
    if (StartTagFlag(tail, "fillDownLabels", false) == fill_down) {
      return raw_ext_xml;
    }
    return raw_ext_xml.substr(0, lt) + WithStartTagFlag(tail, "fillDownLabels", fill_down, false);
  }
  return raw_ext_xml;
}

/// True when the field's x14 `<pivotField>` extension is generated from the
/// model rather than kept verbatim in `raw_ext_xml`.
bool GeneratesFieldX14Ext(const pivot::PivotField& field) {
  return (field.fill_down_labels || !field.x14_attrs.empty() || !field.x14_children_xml.empty()) &&
         !strings::case_insensitive_contains(field.raw_ext_xml, kPivotFieldExtUri);
}

/// True when the field writes an `<extLst>`.
bool HasFieldExtLst(const pivot::PivotField& field) {
  return field.fill_down_labels || !field.x14_attrs.empty() || !field.x14_children_xml.empty() ||
         !field.raw_ext_xml.empty();
}

/// Emits the field's `<extLst>`: the x14 `<pivotField>` extension when it
/// carries anything, then the other extensions verbatim.
void AppendFieldExtLst(std::string& out, const pivot::PivotField& field) {
  out.append("<extLst>");
  if (GeneratesFieldX14Ext(field)) {
    out.append("<ext uri=\"");
    out.append(kPivotFieldExtUri);
    out.append("\"");
    append_xml_ns_decl(out, "x14", kNsX14);
    out.append("><x14:pivotField");
    if (field.fill_down_labels) {
      out.append(" fillDownLabels=\"1\"");
    }
    append_raw_attrs(out, field.x14_attrs);
    if (field.x14_children_xml.empty()) {
      out.append("/></ext>");
    } else {
      out.push_back('>');
      out.append(field.x14_children_xml);
      out.append("</x14:pivotField></ext>");
    }
  }
  out.append(WithVerbatimFillDownLabels(field.raw_ext_xml, field.fill_down_labels));
  out.append("</extLst>");
}

/// Emits one `<pivotField>` element. Self-closing when there are no
/// items; open/close pair otherwise.
///
/// `items` are the data items to write, which are `field.items` unless the
/// caller bound the field to derived shared items.
///
/// `used_as_data` says whether some `<dataField fld>` refers to this field.
/// Excel refuses a `dataField="1"` field that no data field refers to, so the
/// flag follows the data-field list rather than the field's axis.
///
/// `measure_filter` says whether a value filter is written against the field,
/// which Excel marks with `measureFilter="1"`.
void AppendPivotField(std::string& out, const pivot::PivotField& field, const std::vector<pivot::PivotItem>& items,
                      pivot::PivotLayout layout, bool used_as_data, bool measure_filter) {
  // Modelled and captured attributes alike are written in CT_PivotField order,
  // as Excel does; names the schema does not list follow.
  AttrList attrs;
  if (!field.custom_name.empty()) {
    attrs.emplace_back("name", field.custom_name);
  }
  if (field.axis != pivot::PivotAxis::Value && field.axis != pivot::PivotAxis::None) {
    attrs.emplace_back("axis", std::string(enum_name(kPivotAxisNames, field.axis)));
  }
  if (used_as_data) {
    attrs.emplace_back("dataField", "1");
  }
  // The report form is per field in Excel (`compact` / `outline` default to
  // true); every field carries the table's form, as Excel writes it, unless
  // the field was read with a form of its own (kept in `passthrough_attrs`).
  if (!HasFormOverride(field.passthrough_attrs)) {
    if (layout != pivot::PivotLayout::Compact) {
      attrs.emplace_back("compact", "0");
    }
    if (layout == pivot::PivotLayout::Tabular) {
      attrs.emplace_back("outline", "0");
    }
  }
  if (!field.number_format.empty()) {
    attrs.emplace_back("numFmtId", field.number_format);
  }
  // `subtotalTop` defaults to true in OOXML; only emit it when turned OFF
  // so an explicit "subtotals at bottom" choice survives the round trip.
  if (!field.subtotal_top) {
    attrs.emplace_back("subtotalTop", "0");
  }
  // Excel writes `showAll="0"` on every field; without it a refresh shows
  // items that have no data. `true` is the schema default and goes unwritten.
  if (!field.show_all) {
    attrs.emplace_back("showAll", "0");
  }
  if (field.insert_blank_row) {
    attrs.emplace_back("insertBlankRow", "1");
  }
  if (measure_filter && std::none_of(field.passthrough_attrs.begin(), field.passthrough_attrs.end(),
                                     [](const auto& attr) { return attr.first == "measureFilter"; })) {
    attrs.emplace_back("measureFilter", "1");
  }
  // An absent `sortType` reads back as manual (`<items>` order), so a
  // label-sorted axis field spells its direction out, as does a field off
  // the axes whose order was stated rather than defaulted.
  const bool on_axis = field.axis == pivot::PivotAxis::Row || field.axis == pivot::PivotAxis::Col ||
                       field.axis == pivot::PivotAxis::Page;
  if ((on_axis || field.sort.stated) && !field.sort.manual) {
    attrs.emplace_back("sortType", field.sort.ascending ? "ascending" : "descending");
  }
  // Custom functions replace the automatic subtotal, and Excel then writes
  // neither `defaultSubtotal` nor a `t="default"` item. Without custom
  // functions, `defaultSubtotal` is emitted only when turned OFF so an
  // explicit suppression survives the round trip.
  const std::vector<pivot::SubtotalFn> subtotal_fns = pivot::ordered_subtotal_fns(field);
  const bool automatic_subtotal = field.default_subtotal && subtotal_fns.empty();
  if (!field.default_subtotal && subtotal_fns.empty()) {
    attrs.emplace_back("defaultSubtotal", "0");
  }
  for (const pivot::SubtotalFn fn : subtotal_fns) {
    if (const PivotSubtotalName* names = find_pivot_subtotal_name(fn); names != nullptr) {
      attrs.emplace_back(std::string(names->attr), "1");
    }
  }
  attrs.insert(attrs.end(), field.passthrough_attrs.begin(), field.passthrough_attrs.end());
  out.append("<pivotField");
  AppendAttrsInOrder(out, attrs, kPivotFieldAttrOrder);
  if (items.empty() && field.marker_items.empty()) {
    if (field.raw_children_xml.empty() && !HasFieldExtLst(field)) {
      out.append("/>");
    } else {
      out.push_back('>');
      out.append(field.raw_children_xml);
      if (HasFieldExtLst(field)) {
        AppendFieldExtLst(out, field);
      }
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
  std::size_t count = items.size() + (keep_subtotals ? 0U : expected.size());
  for (const pivot::PivotMarkerItem& marker : field.marker_items) {
    count += keep_subtotals || !is_subtotal(marker.type) ? 1U : 0U;
  }
  out.append("><items count=\"");
  out.append(std::to_string(count));
  out.append("\">");
  const auto append_blanks_at = [&](std::size_t position) {
    for (const pivot::PivotMarkerItem& marker : field.marker_items) {
      const std::size_t anchor = std::min(marker.after_items, items.size());
      if (!closes_list(marker) && anchor == position) {
        out.append(marker.xml);
      }
    }
  };
  for (std::size_t i = 0; i < items.size(); ++i) {
    append_blanks_at(i);
    const pivot::PivotItem& item = items[i];
    // Re-emit the cache index the reader captured from `<item x="N">` so
    // the item still points at the same `shared_items` entry after a
    // round trip. When the source omitted `x` (or the item was built
    // from scratch), fall back to the document-order position, matching
    // Excel's implicit ordering. Item names are resolved against the
    // cache, not stored on the element, so no name attribute is emitted.
    const std::uint32_t x = item.has_cache_index ? item.cache_index : static_cast<std::uint32_t>(i);
    out.append("<item");
    if (item.calculated) {
      out.append(" f=\"1\"");
    }
    out.append(" x=\"");
    out.append(std::to_string(x));
    out.append("\"");
    if (!item.visible) {
      out.append(" h=\"1\"");
    }
    if (!item.caption.empty()) {
      append_xml_attr(out, "n", item.caption);
    }
    if (!item.show_detail) {
      out.append(" sd=\"0\"");
    }
    append_raw_attrs(out, item.passthrough_attrs);
    out.append("/>");
  }
  append_blanks_at(items.size());
  // A collapsed field carries `sd="0"` on its `default` marker too; the
  // captured marker's `sd` is patched once its state no longer matches.
  bool default_shows_detail = true;
  for (const pivot::PivotMarkerItem& marker : field.marker_items) {
    if (marker.type == "default") {
      default_shows_detail = marker.show_detail;
    }
  }
  const std::string_view default_marker =
      default_shows_detail ? "<item t=\"default\"/>" : "<item t=\"default\" sd=\"0\"/>";
  if (keep_subtotals) {
    for (const pivot::PivotMarkerItem& marker : field.marker_items) {
      if (!is_subtotal(marker.type)) {
        continue;
      }
      if (marker.type == "default" && marker.show_detail != StartTagFlag(marker.xml, "sd", true)) {
        out.append(WithStartTagFlag(marker.xml, "sd", marker.show_detail, true));
      } else {
        out.append(marker.xml);
      }
    }
  } else {
    for (const std::string_view type : expected) {
      if (type == "default") {
        out.append(default_marker);
        continue;
      }
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
  if (HasFieldExtLst(field)) {
    AppendFieldExtLst(out, field);
  }
  out.append("</pivotField>");
}

/// The items of a field the plan derives as discrete, bound to the derived
/// shared items; `field.items` verbatim for any other field.
///
/// An axis field with no authored items gets one per shared item, in its sort
/// order. Authored items resolve to a shared item by their cache index, else
/// by label (as `fm_workbook_pivot_field_add_item` binds), else by position;
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

/// The `<pivotFields>` index a public filter is written against, or
/// `nullopt` when it designates nothing a file can name.
///
/// A value filter prunes the leaves of its axis: the named field when it sits
/// on that axis, otherwise the axis's innermost field, which is how the value
/// pass reads an entry added through the API. A decoded entry and the caption
/// and date filters apply to the field they name.
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
  if (order.empty() && !filter.source.has_value()) {
    return std::nullopt;
  }
  if (filter.source.has_value() || std::find(order.begin(), order.end(), *named) != order.end()) {
    return static_cast<std::uint32_t>(*named);
  }
  return order.back();
}

/// True when a decoded filter is written back as its source bytes: it still
/// designates the field and data field it was read with, or its name never
/// resolved (the pivot had no cache to name it from).
bool ReEmitsSource(const pivot::PivotTable& table, const pivot::PivotFilter& filter) {
  return filter.source.has_value() &&
         (filter.field_name.empty() || (FilterFieldIndex(table, filter) == filter.source->field_index &&
                                        filter.data_field_index == filter.source->data_field_index));
}

/// A numeric filter payload; empty when unset, text, or not finite, which no
/// file can spell.
std::optional<double> FilterNumber(const std::variant<std::monostate, int, double, std::string>& value) {
  if (const auto* d = std::get_if<double>(&value)) {
    return std::isfinite(*d) ? std::optional<double>(*d) : std::nullopt;
  }
  if (const auto* i = std::get_if<int>(&value)) {
    return static_cast<double>(*i);
  }
  return std::nullopt;
}

/// Appends `<customFilter operator="..." val="..."/>`; an empty `op` is the
/// schema's default `equal` and is omitted, as Excel does.
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
  if (!op.empty()) {
    append_xml_attr(out, "operator", op);
  }
  append_xml_attr_number(out, "val", number);
  out.append("/>");
}

/// `text` with its `~` escapes removed: the literal a plain `<filters>`
/// value spells.
std::string UnescapeWildcards(std::string_view text) {
  std::string out;
  out.reserve(text.size());
  for (std::size_t i = 0; i < text.size(); ++i) {
    if (text[i] == '~' && i + 1U < text.size()) {
      ++i;
    }
    out.push_back(text[i]);
  }
  return out;
}

/// True when `text` holds a `*` or `?` that is not escaped by `~`.
bool HasLiveWildcard(std::string_view text) {
  for (std::size_t i = 0; i < text.size(); ++i) {
    if (text[i] == '~') {
      ++i;
    } else if (text[i] == '*' || text[i] == '?') {
      return true;
    }
  }
  return false;
}

/// The `<customFilters>` comparison operator of a single-operand value,
/// caption ordering or date comparison type; empty for `equal`.
std::string_view ComparisonOperator(pivot::FilterType type) {
  switch (type) {
    case pivot::FilterType::ValueNotEqual:
    case pivot::FilterType::LabelNotEqual:
    case pivot::FilterType::LabelNotBeginsWith:
    case pivot::FilterType::LabelNotEndsWith:
    case pivot::FilterType::LabelNotContains:
    case pivot::FilterType::DateNotEqual:
      return "notEqual";
    case pivot::FilterType::ValueGreaterThan:
    case pivot::FilterType::LabelGreaterThan:
    case pivot::FilterType::DateNewerThan:
      return "greaterThan";
    case pivot::FilterType::ValueGreaterThanOrEqual:
    case pivot::FilterType::LabelGreaterThanOrEqual:
    case pivot::FilterType::DateNewerThanOrEqual:
      return "greaterThanOrEqual";
    case pivot::FilterType::ValueLessThan:
    case pivot::FilterType::LabelLessThan:
    case pivot::FilterType::DateOlderThan:
      return "lessThan";
    case pivot::FilterType::ValueLessThanOrEqual:
    case pivot::FilterType::LabelLessThanOrEqual:
    case pivot::FilterType::DateOlderThanOrEqual:
      return "lessThanOrEqual";
    default:
      return {};
  }
}

/// True for the inclusive ranges, written `and="1"` with `>=` and `<=`; the
/// exclusive ones are written without `and` as `<` low, `>` high.
bool IsInclusiveRange(pivot::FilterType type) {
  return type == pivot::FilterType::ValueBetween || type == pivot::FilterType::LabelBetween ||
         type == pivot::FilterType::LabelDate;
}

/// Serialises one model filter as Excel writes it from the filter dialogs.
/// Returns false, emitting nothing, for an entry that filters nothing and
/// that the file has no spelling for: a missing or non-finite operand, a
/// range with no upper bound, a payload of the wrong kind.
bool AppendGeneratedFilter(std::string& out, const pivot::PivotFilter& filter, std::uint32_t field_index,
                           std::uint32_t id) {
  const pivot::FilterType type = filter.type;
  const bool caption = pivot::is_caption_filter_type(type);
  const bool ranged = pivot::is_range_filter_type(type);
  const bool operandless = pivot::is_operandless_filter_type(type);
  const std::string* low_text = std::get_if<std::string>(&filter.value);
  const std::string* high_text = std::get_if<std::string>(&filter.value_high);
  const std::optional<double> low = FilterNumber(filter.value);
  const std::optional<double> high = FilterNumber(filter.value_high);
  if (caption ? low_text == nullptr || (ranged && high_text == nullptr)
              : !operandless && (!low.has_value() || (ranged && !high.has_value()))) {
    return false;
  }
  out.append("<filter");
  append_xml_attr_uint(out, "fld", field_index);
  append_xml_attr(out, "type", kFilterTypeNames[static_cast<std::size_t>(type)]);
  out.append(" evalOrder=\"-1\"");
  append_xml_attr_uint(out, "id", id);
  if (pivot::is_value_filter_type(type)) {
    append_xml_attr_uint(out, "iMeasureFld", filter.data_field_index);
  }
  if (caption) {
    append_xml_attr(out, "stringValue1", *low_text);
    if (ranged) {
      append_xml_attr(out, "stringValue2", *high_text);
    }
  }
  out.append("><autoFilter ref=\"A1\"><filterColumn colId=\"0\">");
  const auto append_operand = [&](std::string_view op, bool high_side) {
    if (caption) {
      AppendCustomFilter(out, op, high_side ? *high_text : *low_text);
    } else {
      AppendCustomFilter(out, op, high_side ? *high : *low);
    }
  };
  switch (type) {
    case pivot::FilterType::ValueTop10:
    case pivot::FilterType::ValueBottomCount:
    case pivot::FilterType::ValueTopPercent:
    case pivot::FilterType::ValueBottomPercent:
    case pivot::FilterType::ValueTopSum:
    case pivot::FilterType::ValueBottomSum:
      out.append("<top10");
      if (type == pivot::FilterType::ValueBottomCount || type == pivot::FilterType::ValueBottomPercent ||
          type == pivot::FilterType::ValueBottomSum) {
        out.append(" top=\"0\"");
      }
      if (type == pivot::FilterType::ValueTopPercent || type == pivot::FilterType::ValueBottomPercent) {
        out.append(" percent=\"1\"");
      }
      append_xml_attr_number(out, "val", *low);
      append_xml_attr_number(out, "filterVal", *low);
      out.append("/>");
      break;
    case pivot::FilterType::LabelEqual:
      // A literal operand is written as a plain value list with its escapes
      // removed; a pattern or a number as a criterion.
      if (double number = 0.0; !HasLiveWildcard(*low_text) && !parse_xsd_double(*low_text, &number)) {
        out.append("<filters><filter");
        append_xml_attr(out, "val", UnescapeWildcards(*low_text));
        out.append("/></filters>");
        break;
      }
      out.append("<customFilters>");
      append_operand({}, false);
      out.append("</customFilters>");
      break;
    default:
      if (operandless) {
        out.append("<dynamicFilter");
        append_xml_attr(out, "type", kFilterTypeNames[static_cast<std::size_t>(type)]);
        out.append("/>");
      } else if (const std::optional<std::string> pattern =
                     caption ? pivot::caption_wildcard_pattern(type, *low_text) : std::nullopt;
                 pattern.has_value()) {
        out.append("<customFilters>");
        AppendCustomFilter(out, ComparisonOperator(type), *pattern);
        out.append("</customFilters>");
      } else if (ranged) {
        const bool inclusive = IsInclusiveRange(type);
        out.append(inclusive ? "<customFilters and=\"1\">" : "<customFilters>");
        append_operand(inclusive ? "greaterThanOrEqual" : "lessThan", false);
        append_operand(inclusive ? "lessThanOrEqual" : "greaterThan", true);
        out.append("</customFilters>");
      } else {
        out.append("<customFilters>");
        append_operand(ComparisonOperator(type), false);
        out.append("</customFilters>");
      }
      break;
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
    if (ReEmitsSource(table, filter)) {
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
  std::uint32_t ref_top = table.anchor_row();
  if (rendered_span.has_value() && !table.has_authored_span()) {
    span_rows = span_rows > rendered_span->rows ? span_rows : rendered_span->rows;
    span_cols = span_cols > rendered_span->cols ? span_cols : rendered_span->cols;
    // A page-field block taller than the anchor's headroom pushes the body down.
    ref_top = ref_top > rendered_span->top ? ref_top : rendered_span->top;
  }
  AppendXmlAttrEscaped(out, EncodeA1Range(ref_top, table.anchor_col(), span_rows, span_cols));
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
  // rowPageCount / colPageCount are optional: a table without page fields omits them.
  AppendOptionalLocationAttr(
      out, "rowPageCount",
      table.location_row_page_count() ? table.location_row_page_count() : loc_defaults.row_page_count);
  AppendOptionalLocationAttr(
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
