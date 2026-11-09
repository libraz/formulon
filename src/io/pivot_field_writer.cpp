//
// Emission of one `<pivotField>` element of a pivot-table definition: its
// attributes, items, and `<extLst>` with the x14 field extension.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "io/ooxml_defs.h"
#include "io/pivot_attr_names.h"
#include "io/pivot_table_writer_internal.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "utils/strings.h"

namespace formulon::io::detail {
namespace {

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
    if (name.substr(qname_prefix_colon(name).size()) != "pivotField") {
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

}  // namespace

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

}  // namespace formulon::io::detail
