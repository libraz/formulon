//
// Implementation of the pivot-table-definition reader. See
// pivot_table_reader.h for the public contract.

#include "io/pivot_table_reader.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <iterator>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <variant>
#include <vector>

#include "io/cell_parser.h"
#include "io/pivot_attr_names.h"
#include "io/xml_utils.h"
#include "io/xsd_double.h"
#include "pivot/filter_engine.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "utils/strings.h"

namespace formulon::io {
namespace {

/// Decodes one `<location ref="A3:D10"/>` (or single-cell `ref="A3"`)
/// into 0-based anchor + spans. Defers cell-coordinate decoding to the
/// shared `parse_a1` helper so the conventions match cell parsing
/// elsewhere.
Expected<void, Error> DecodeLocationRef(std::string_view ref, pivot::PivotTable* out) {
  if (ref.empty()) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt,
                      "pivotTableDefinition: <location> missing required ref attribute", "context=pivot_table_reader");
  }
  // A single-cell ref is the anchor Excel writes for an empty pivot.
  MergeRange rect;
  if (!parse_a1_range(ref, &rect)) {
    std::string ctx("context=pivot_table_reader ref=");
    ctx.append(ref);
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "pivotTableDefinition: <location> ref unparseable",
                      std::move(ctx));
  }
  out->set_anchor(rect.first_row, rect.first_col, rect.last_row - rect.first_row + 1U,
                  rect.last_col - rect.first_col + 1U);
  return Expected<void, Error>::Ok();
}

/// Walks the `<items>` child of a `<pivotField>` and populates
/// `field.items`. `<item t="default">` and `<item t="grand">` are
/// subtotal / grand-total markers in the OOXML schema, NOT real items;
/// they go to `field.marker_items` instead. For real items we capture visibility from
/// `h="1"` (hidden -> `visible = false`) and the `x` attribute, which
/// indexes the bound cache's `shared_items`. `name` is left empty here
/// and resolved against the cache after both parts are loaded (see
/// `resolve_pivot_names`). When `x` is absent we fall back to the item's
/// document-order position among real items so name resolution still has
/// an index to look up; `has_cache_index` stays false so the writer can
/// re-emit the original (attribute-absent) form.
Expected<void, Error> ParseItems(const pugi::xml_node& items_node, pivot::PivotField* field) {
  std::uint32_t ordinal = 0;
  for (pugi::xml_node it = items_node.child("item"); it; it = it.next_sibling("item")) {
    const std::string_view t = it.attribute("t").as_string();
    bool is_marker = t == "default" || t == "grand" || t == "blank";
    for (const PivotSubtotalName& e : kPivotSubtotalNames) {
      is_marker = is_marker || t == e.item;
    }
    if (is_marker) {
      // Subtotal / grand-total / blank marker; not a real item, but kept so
      // its attributes survive.
      field->marker_items.push_back(
          pivot::PivotMarkerItem{raw_xml(it), std::string(t), ordinal, attr_bool(it, "sd", true)});
      continue;
    }
    pivot::PivotItem entry;
    entry.visible = !parse_xml_bool_attr(it.attribute("h"));
    entry.caption = it.attribute("n").as_string();
    entry.show_detail = attr_bool(it, "sd", true);
    entry.calculated = attr_bool(it, "f", false);
    capture_unknown_attrs(it, {"x", "h", "t", "n", "sd", "f"}, entry.passthrough_attrs);
    if (pugi::xml_attribute x = it.attribute("x"); x) {
      const std::optional<std::uint32_t> cache_index = parse_xml_u32_attr_strict(x);
      if (!cache_index.has_value()) {
        return make_error(FormulonErrorCode::kIoSheetCorrupt, "pivotTableDefinition: <item> has invalid x attribute",
                          "context=pivot_table_reader");
      }
      entry.has_cache_index = true;
      entry.cache_index = *cache_index;
    } else {
      entry.cache_index = ordinal;
    }
    field->items.push_back(std::move(entry));
    ++ordinal;
  }
  return Expected<void, Error>::Ok();
}

/// Walks `<pivotFields>` in document order, materialising each
/// `<pivotField>` into the table.
Expected<void, Error> ParsePivotFields(const pugi::xml_node& fields_node, pivot::PivotTable* out) {
  for (pugi::xml_node f = fields_node.child("pivotField"); f; f = f.next_sibling("pivotField")) {
    pivot::PivotField field;
    // Axis resolution: an explicit `axis` attribute wins; otherwise a
    // `dataField="1"` marks a Value field; a field with neither is an
    // unused ("available") field, which must round-trip as None rather
    // than being stamped `dataField="1"` on write.
    if (pugi::xml_attribute axis_attr = f.attribute("axis"); axis_attr) {
      field.axis = enum_from_name(kPivotAxisNames, axis_attr.as_string(), pivot::PivotAxis::Value);
    } else if (attr_bool(f, "dataField", false)) {
      field.axis = pivot::PivotAxis::Value;
    } else {
      field.axis = pivot::PivotAxis::None;
    }
    if (pugi::xml_attribute name_attr = f.attribute("name"); name_attr) {
      field.custom_name = name_attr.value();
    }
    if (pugi::xml_attribute nf = f.attribute("numFmtId"); nf) {
      field.number_format = nf.value();
    }
    // `subtotalTop` defaults to true in OOXML (subtotals render above the
    // group). Only an explicit "0" moves them to the bottom.
    field.subtotal_top = attr_bool(f, "subtotalTop", true);
    // `defaultSubtotal` defaults to true in OOXML; only an explicit "0"
    // turns the implicit default subtotal off. The `*Subtotal` boolean
    // family selects custom subtotal functions; capture each present /
    // true attribute so a custom selection survives the round trip
    // instead of silently reverting to the default after save.
    if (pugi::xml_attribute ds = f.attribute("defaultSubtotal"); ds) {
      field.default_subtotal = parse_xml_bool_attr(ds);
    }
    for (const PivotSubtotalName& e : kPivotSubtotalNames) {
      if (parse_xml_bool_attr(f.attribute(e.attr.data()))) {
        field.subtotal_fns.push_back(e.fn);
      }
    }
    // Custom functions replace the default subtotal; Excel writes them
    // without `defaultSubtotal` or a `t="default"` item.
    if (!field.subtotal_fns.empty()) {
      field.default_subtotal = false;
    }
    // An absent `sortType` is the schema's "manual" default, so
    // the field follows its `<items>` document order.
    if (pugi::xml_attribute sort_attr = f.attribute("sortType"); sort_attr) {
      field.sort.stated = true;
      const std::string_view value = sort_attr.value();
      if (value == "descending") {
        field.sort.ascending = false;
      } else if (value == "manual") {
        field.sort.manual = true;
      }
    } else {
      field.sort.manual = true;
    }
    if (pugi::xml_node items_node = f.child("items"); items_node) {
      RETURN_IF_ERROR(ParseItems(items_node, &field));
    }
    // Preserve unmodelled `<pivotField>` attributes (`includeNewItemsInFilter`, ...) and children for verbatim
    // round-trip. `compact` / `outline` are the report form, which the writer derives from the table's layout;
    // `CaptureFormOverride` keeps them once the layout is known and they depart from it.
    field.insert_blank_row = attr_bool(f, "insertBlankRow", false);
    field.show_all = attr_bool(f, "showAll", true);
    capture_unknown_attrs(f, {"axis",          "dataField",      "name",
                              "numFmtId",      "subtotalTop",    "defaultSubtotal",
                              "sumSubtotal",   "countASubtotal", "avgSubtotal",
                              "maxSubtotal",   "minSubtotal",    "productSubtotal",
                              "countSubtotal", "stdDevSubtotal", "stdDevPSubtotal",
                              "varSubtotal",   "varPSubtotal",   "sortType",
                              "compact",       "outline",        "insertBlankRow",
                              "showAll"},
                          field.passthrough_attrs);
    capture_unknown_children(f, {"items", "extLst"}, field.raw_children_xml);
    // Excel spells the extension's URI with a lower-case `432a`, so it is
    // matched case-insensitively. An extension holding anything beside its one
    // `<pivotField>` (found under any prefix) is kept whole; `fillDownLabels`
    // is still read from it.
    for (pugi::xml_node ext : f.child("extLst").children("ext")) {
      if (!strings::case_insensitive_eq(ext.attribute("uri").as_string(), kPivotFieldExtUri)) {
        append_raw_xml(field.raw_ext_xml, ext);
        continue;
      }
      const pugi::xml_node x14 = child_local(ext, "pivotField");
      std::size_t elements = 0;
      for (pugi::xml_node child = ext.first_child(); child; child = child.next_sibling()) {
        elements += child.type() == pugi::node_element ? 1U : 0U;
      }
      field.fill_down_labels = attr_bool(x14, "fillDownLabels", false);
      // The writer spells the element `x14:`; children under another prefix
      // would lose their binding.
      const bool foreign_children = std::string_view(x14.name()) != "x14:pivotField" && has_element_child(x14);
      if (!x14 || elements != 1U || foreign_children) {
        append_raw_xml(field.raw_ext_xml, ext);
        continue;
      }
      capture_unknown_attrs(x14, {"fillDownLabels"}, field.x14_attrs);
      for (pugi::xml_node child = x14.first_child(); child; child = child.next_sibling()) {
        if (child.type() == pugi::node_element) {
          append_raw_xml(field.x14_children_xml, child);
        }
      }
    }
    out->mutable_fields().push_back(std::move(field));
  }
  return Expected<void, Error>::Ok();
}

/// Walks `<rowFields>` / `<colFields>` and pushes each real `<field
/// x="N">` onto `out`. Missing `x` defaults to 0 (matches OOXML schema).
///
/// `x="-2"` (`pivot::kValuesFieldPosition`) is the Values pseudo-field
/// marker Excel writes once the pivot has two or more data fields: it
/// names no field in `fields()`, so it is kept out of `out` and its
/// document position (counted over every `<field>` child including its
/// own slot) is recorded in `*out_values_position` instead. Every
/// consumer that walks `out` to build the row/column hierarchy
/// therefore never needs to recognise it; the writer re-inserts it at
/// that position on save.
void ParseFieldOrder(const pugi::xml_node& parent, std::vector<std::uint32_t>* out,
                     std::optional<std::size_t>* out_values_position) {
  std::size_t position = 0;
  for (pugi::xml_node f = parent.child("field"); f; f = f.next_sibling("field"), ++position) {
    const std::int32_t x = attr_i32(f, "x");
    if (x == pivot::kValuesFieldPosition) {
      if (!out_values_position->has_value()) {
        *out_values_position = position;
      }
      continue;
    }
    out->push_back(x >= 0 ? static_cast<std::uint32_t>(x) : 0U);
  }
}

/// Walks `<pageFields>` in document order, recording each page field's
/// position on the axis and its single-item selection (`item`).
///
/// A `<pageField>` with no `fld` is skipped rather than defaulted to field
/// 0: unlike `<field x>` on a row/column axis, where 0 is a real index the
/// element is simply spelling out, an absent `fld` here names no field at
/// all, and inventing one would place a page header over a field the user
/// never filtered by.
void ParsePageFields(const pugi::xml_node& parent, pivot::PivotTable* out) {
  for (pugi::xml_node node = parent.child("pageField"); node; node = node.next_sibling("pageField")) {
    const std::optional<std::uint32_t> field_index = parse_xml_u32_attr_strict(node.attribute("fld"));
    if (!field_index.has_value()) {
      continue;
    }
    pivot::PivotPageField entry;
    entry.field_index = *field_index;
    entry.item_index = parse_xml_u32_attr_strict(node.attribute("item"));
    out->mutable_page_fields().push_back(entry);
  }
}

/// Walks `<dataFields>` in document order. A missing or empty `name` is
/// kept as an empty name; no default is derived on load.
Expected<void, Error> ParseDataFields(const pugi::xml_node& parent, pivot::PivotTable* out) {
  for (pugi::xml_node df = parent.child("dataField"); df; df = df.next_sibling("dataField")) {
    pivot::PivotDataField entry;
    entry.name = df.attribute("name").as_string();
    entry.field_index = attr_u32(df, "fld");
    entry.aggregation =
        enum_from_name(kPivotAggregationNames, df.attribute("subtotal").as_string(), pivot::Aggregation::Sum);
    if (pugi::xml_attribute nf = df.attribute("numFmtId"); nf) {
      entry.number_format = nf.value();
    }
    if (pugi::xml_attribute sa = df.attribute("showDataAs"); sa) {
      entry.show_as = enum_from_name(kPivotShowDataAsNames, sa.as_string(), pivot::ShowValuesAs::Normal);
    }
    if (pugi::xml_attribute bf = df.attribute("baseField"); bf) {
      entry.show_as_base_field = parse_xml_u32_attr(bf, 0U);
    }
    if (pugi::xml_attribute bi = df.attribute("baseItem"); bi) {
      entry.show_as_base_item = parse_xml_u32_attr(bi, 0U);
    }
    for (pugi::xml_node ext : df.child("extLst").children("ext")) {
      if (ext.attribute("uri").as_string() != kPivotDataFieldExtUri) {
        append_raw_xml(entry.raw_ext_xml, ext);
        continue;
      }
      const pugi::xml_node x14 = ext.child("x14:dataField");
      if (pugi::xml_attribute sa = x14.attribute("pivotShowAs"); sa) {
        entry.show_as = enum_from_name(kPivotShowAsX14Names, sa.as_string(), entry.show_as);
      }
      capture_unknown_attrs(x14, {"pivotShowAs"}, entry.x14_attrs);
      for (pugi::xml_node child = x14.first_child(); child; child = child.next_sibling()) {
        if (child.type() == pugi::node_element) {
          append_raw_xml(entry.x14_children_xml, child);
        }
      }
    }
    out->mutable_data_fields().push_back(std::move(entry));
  }
  return Expected<void, Error>::Ok();
}

/// Collects the criterion strings a `<filter>`'s nested `<autoFilter>`
/// carries, in document order.
///
/// Two shapes reach the same place: `<customFilters>` holds one
/// `<customFilter val="...">` for a single-sided comparison and two for a
/// `Between`, while a plain equality is also spelled as a `<filters>`
/// list of `<filter val="...">`, whose values are literal text rather than
/// patterns (`plain` is set for it). Anything else (a `<top10>` or a
/// `<dynamicFilter>`) contributes no criterion.
std::vector<std::string_view> CollectFilterCriteria(const pugi::xml_node& filter_node, bool* plain = nullptr) {
  std::vector<std::string_view> out;
  if (plain != nullptr) {
    *plain = false;
  }
  const pugi::xml_node column = filter_node.child("autoFilter").child("filterColumn");
  if (!column) {
    return out;
  }
  if (pugi::xml_node custom = column.child("customFilters"); custom) {
    for (pugi::xml_node c = custom.child("customFilter"); c; c = c.next_sibling("customFilter")) {
      if (pugi::xml_attribute val = c.attribute("val"); val) {
        out.emplace_back(val.value());
      }
    }
    return out;
  }
  if (pugi::xml_node list = column.child("filters"); list) {
    if (plain != nullptr) {
      *plain = true;
    }
    for (pugi::xml_node f = list.child("filter"); f; f = f.next_sibling("filter")) {
      if (pugi::xml_attribute val = f.attribute("val"); val) {
        out.emplace_back(val.value());
      }
    }
  }
  return out;
}

/// Reads the item count out of a top-N filter's nested
/// `<top10 val="N"/>`. Returns `nullopt` when the element is absent, so
/// the entry is not decoded rather than defaulting to a count nobody wrote.
std::optional<double> CollectTopCount(const pugi::xml_node& filter_node) {
  const pugi::xml_node top = filter_node.child("autoFilter").child("filterColumn").child("top10");
  if (!top) {
    return std::nullopt;
  }
  const pugi::xml_attribute val = top.attribute("val");
  if (!val) {
    return std::nullopt;
  }
  double parsed = 0.0;
  if (!parse_xsd_double(val.value(), &parsed)) {
    return std::nullopt;
  }
  return parsed;
}

/// Parses one `<customFilter val="...">` payload, rejecting anything the
/// sheet path's lexer would not accept as a number. Sharing that lexer
/// keeps `0x10` from meaning sixteen here and nothing elsewhere.
std::optional<double> ParseCriterionNumber(std::string_view text) {
  double parsed = 0.0;
  if (!parse_xsd_double(text, &parsed)) {
    return std::nullopt;
  }
  return parsed;
}

/// A top-N amount as the model stores it: an `int` when the file spells a
/// whole number in range, so it reads back the way the API takes it.
std::variant<std::monostate, int, double, std::string> TopCountValue(double count) {
  if (count >= 0.0 && count <= static_cast<double>(std::numeric_limits<int>::max()) &&
      count == static_cast<double>(static_cast<int>(count))) {
    return static_cast<int>(count);
  }
  return count;
}

/// The `FilterType` a `<filter type>` spelling names, with the top / bottom
/// pairs resolved by `<top10 top>`. Empty for `unknown` and anything newer.
std::optional<pivot::FilterType> ParseFilterType(const pugi::xml_node& node) {
  const std::string_view name = attr_str(node, "type");
  for (std::size_t i = 0; i < std::size(kFilterTypeNames); ++i) {
    if (kFilterTypeNames[i] != name) {
      continue;
    }
    const auto type = static_cast<pivot::FilterType>(i);
    const pugi::xml_node top = node.child("autoFilter").child("filterColumn").child("top10");
    if (attr_bool(top, "top", true)) {
      return type;
    }
    switch (type) {
      case pivot::FilterType::ValueTop10:
        return pivot::FilterType::ValueBottomCount;
      case pivot::FilterType::ValueTopPercent:
        return pivot::FilterType::ValueBottomPercent;
      case pivot::FilterType::ValueTopSum:
        return pivot::FilterType::ValueBottomSum;
      default:
        return type;
    }
  }
  return std::nullopt;
}

/// `text` with the wildcards a plain `<filters>` value holds literally
/// escaped, so it reads as the operand the dialog took.
std::string EscapeWildcards(std::string_view text) {
  std::string out;
  out.reserve(text.size());
  for (const char c : text) {
    if (c == '*' || c == '?' || c == '~') {
      out.push_back('~');
    }
    out.push_back(c);
  }
  return out;
}

/// The operand of a begins / ends / contains `pattern`: the pattern without
/// the `*` the type affixes. Empty when the pattern does not carry them.
std::optional<std::string> StripCaptionAffix(pivot::FilterType type, std::string_view pattern) {
  const bool contains = type == pivot::FilterType::LabelContains || type == pivot::FilterType::LabelNotContains;
  const bool leading =
      contains || type == pivot::FilterType::LabelEndsWith || type == pivot::FilterType::LabelNotEndsWith;
  const bool trailing =
      contains || type == pivot::FilterType::LabelBeginsWith || type == pivot::FilterType::LabelNotBeginsWith;
  if (pattern.size() < (contains ? 2U : 1U)) {
    return std::nullopt;
  }
  if (leading) {
    if (pattern.front() != '*') {
      return std::nullopt;
    }
    pattern.remove_prefix(1);
  }
  if (trailing) {
    if (pattern.empty() || pattern.back() != '*') {
      return std::nullopt;
    }
    // An odd run of `~` before the last `*` escapes it.
    std::size_t tildes = 0;
    while (tildes + 1U < pattern.size() && pattern[pattern.size() - 2U - tildes] == '~') {
      ++tildes;
    }
    if (tildes % 2U != 0U) {
      return std::nullopt;
    }
    pattern.remove_suffix(1);
  }
  return std::string(pattern);
}

/// Decodes a caption `<filter>`'s operands into `out`. False when they do
/// not decode.
///
/// Excel writes the operand twice: `stringValue1` / `stringValue2` as
/// entered, and the nested `<autoFilter>` criterion. For begins / ends /
/// contains Excel refreshes from the nested pattern, so its affixes are
/// stripped to recover the operand and `stringValue1` is the fallback. The
/// other types prefer the direct attributes; a plain `<filters>` value is
/// literal text and gets its wildcards escaped.
bool DecodeCaptionOperands(const pugi::xml_node& node, pivot::PivotFilter& out) {
  const bool ranged = pivot::is_range_filter_type(out.type);
  const pugi::xml_attribute direct_low = node.attribute("stringValue1");
  const pugi::xml_attribute direct_high = node.attribute("stringValue2");
  bool plain = false;
  const std::vector<std::string_view> criteria = CollectFilterCriteria(node, &plain);
  if (pivot::caption_wildcard_pattern(out.type, "").has_value()) {
    if (!criteria.empty()) {
      std::optional<std::string> operand =
          plain ? std::optional<std::string>(EscapeWildcards(criteria[0])) : StripCaptionAffix(out.type, criteria[0]);
      if (!operand.has_value()) {
        return false;
      }
      out.value = std::move(*operand);
      return true;
    }
    if (!direct_low) {
      return false;
    }
    out.value = std::string(direct_low.value());
    return true;
  }
  if (direct_low && (!ranged || direct_high)) {
    out.value = std::string(direct_low.value());
    if (ranged) {
      out.value_high = std::string(direct_high.value());
    }
    return true;
  }
  if (criteria.size() < (ranged ? 2U : 1U)) {
    return false;
  }
  const auto operand = [plain](std::string_view text) { return plain ? EscapeWildcards(text) : std::string(text); };
  out.value = operand(criteria[0]);
  if (ranged) {
    out.value_high = operand(criteria[1]);
  }
  return true;
}

/// Decodes a `<filter>` into the `PivotFilter` model, or returns `nullopt`
/// when the model cannot express it: an unknown type, a member-property or
/// OLAP measure filter, a `fld` past the table's fields, or criteria that
/// do not decode. `field_name` is left empty and filled from the cache by
/// `pivot::resolve_pivot_names`.
std::optional<pivot::PivotFilter> DecodeFilter(const pugi::xml_node& node, const pivot::PivotTable& table) {
  const std::optional<std::uint32_t> fld = parse_xml_u32_attr_strict(node.attribute("fld"));
  if (!fld.has_value() || *fld >= table.fields().size()) {
    return std::nullopt;
  }
  if (node.attribute("mpFld") || node.attribute("iMeasureHier")) {
    return std::nullopt;
  }
  const std::optional<pivot::FilterType> type = ParseFilterType(node);
  if (!type.has_value()) {
    return std::nullopt;
  }
  pivot::PivotFilter out;
  out.type = *type;
  out.data_field_index = attr_u32(node, "iMeasureFld");
  switch (*type) {
    case pivot::FilterType::ValueTop10:
    case pivot::FilterType::ValueBottomCount:
    case pivot::FilterType::ValueTopPercent:
    case pivot::FilterType::ValueBottomPercent:
    case pivot::FilterType::ValueTopSum:
    case pivot::FilterType::ValueBottomSum: {
      const std::optional<double> count = CollectTopCount(node);
      if (!count.has_value()) {
        return std::nullopt;
      }
      out.value = TopCountValue(*count);
      break;
    }
    default:
      if (pivot::is_caption_filter_type(*type)) {
        if (!DecodeCaptionOperands(node, out)) {
          return std::nullopt;
        }
      } else if (!pivot::is_operandless_filter_type(*type)) {
        // Value and date comparisons: numeric criteria, two for a range.
        const bool ranged = pivot::is_range_filter_type(*type);
        const std::vector<std::string_view> criteria = CollectFilterCriteria(node);
        if (criteria.size() < (ranged ? 2U : 1U)) {
          return std::nullopt;
        }
        const std::optional<double> low = ParseCriterionNumber(criteria[0]);
        if (!low.has_value()) {
          return std::nullopt;
        }
        out.value = *low;
        if (ranged) {
          const std::optional<double> high = ParseCriterionNumber(criteria[1]);
          if (!high.has_value()) {
            return std::nullopt;
          }
          out.value_high = *high;
        }
      }
      break;
  }
  const auto on = [&](const std::vector<std::uint32_t>& order) {
    return std::find(order.begin(), order.end(), *fld) != order.end();
  };
  const pivot::PivotAxis field_axis = table.fields()[*fld].axis;
  out.axis = on(table.row_field_order())                                                  ? pivot::PivotAxis::Row
             : on(table.col_field_order())                                                ? pivot::PivotAxis::Col
             : field_axis == pivot::PivotAxis::None || pivot::is_value_filter_type(*type) ? pivot::PivotAxis::Row
                                                                                          : field_axis;
  out.source = pivot::PivotFilterSource{raw_xml(node), *fld, out.data_field_index, attr_u32(node, "id")};
  return out;
}

/// Splits a `<filters>` block between `active_filters()` and the verbatim
/// `retained_filters()`, which keeps the entries the model cannot express.
void ParseFilters(const pugi::xml_node& parent, pivot::PivotTable* out) {
  for (pugi::xml_node node = parent.child("filter"); node; node = node.next_sibling("filter")) {
    if (std::optional<pivot::PivotFilter> entry = DecodeFilter(node, *out)) {
      out->mutable_active_filters().push_back(std::move(*entry));
      continue;
    }
    out->mutable_retained_filters().push_back(
        pivot::RetainedPivotFilter{raw_xml(node), out->active_filters().size(), attr_u32(node, "id")});
  }
}

/// Keeps `node`'s report-form attributes in `out` when they depart from what
/// the writer derives for `layout`, so they are re-emitted instead of it.
///
/// `root` selects the `<pivotTableDefinition>` set (`compact`, `compactData`,
/// `outline`, `outlineData`, the form new fields take) over the
/// `<pivotField>` pair (`compact`, `outline`). Each attribute is compared at
/// its schema default when absent.
void CaptureFormOverride(const pugi::xml_node& node, pivot::PivotLayout layout, bool root,
                         std::vector<std::pair<std::string, std::string>>& out) {
  const bool compact = layout == pivot::PivotLayout::Compact;
  const bool outline = layout != pivot::PivotLayout::Tabular;
  bool matches = attr_bool(node, "compact", true) == compact && attr_bool(node, "outline", !root) == outline;
  if (root) {
    matches =
        matches && attr_bool(node, "compactData", true) == compact && attr_bool(node, "outlineData", false) == outline;
  }
  if (matches) {
    return;
  }
  // Every attribute of the set is spelled out, an absent one at its default,
  // since the writer emits these in place of the whole derived set.
  struct FormAttr {
    const char* name;
    bool fallback;
  };
  const FormAttr root_attrs[] = {{"compact", true}, {"compactData", true}, {"outline", false}, {"outlineData", false}};
  const FormAttr field_attrs[] = {{"compact", true}, {"outline", true}};
  const std::size_t count = root ? 4U : 2U;
  for (std::size_t i = 0; i < count; ++i) {
    const FormAttr& entry = root ? root_attrs[i] : field_attrs[i];
    const pugi::xml_attribute attr = node.attribute(entry.name);
    out.emplace_back(entry.name, attr ? std::string(attr.value()) : std::string(entry.fallback ? "1" : "0"));
  }
}

/// Moves `altText` / `altTextSummary` off the x14 `<pivotTableDefinition>`
/// extension into `table`, leaving the rest of `extlst` for the verbatim bin.
void ExtractAltText(pugi::xml_node extlst, pivot::PivotTable* table) {
  for (pugi::xml_node ext : extlst.children("ext")) {
    pugi::xml_node x14 = child_local(ext, "pivotTableDefinition");
    if (!x14) {
      continue;
    }
    if (pugi::xml_attribute attr = x14.attribute("altText"); attr) {
      table->set_alt_text(attr.value());
      x14.remove_attribute(attr);
    }
    if (pugi::xml_attribute attr = x14.attribute("altTextSummary"); attr) {
      table->set_alt_text_summary(attr.value());
      x14.remove_attribute(attr);
    }
  }
}

}  // namespace

bool read_pivot_filters_xml(std::string_view filters_xml, pivot::PivotTable* table) {
  pugi::xml_document doc;
  if (!doc.load_buffer(filters_xml.data(), filters_xml.size(), pugi::parse_default, pugi::encoding_utf8)) {
    return false;
  }
  ParseFilters(doc.child("filters"), table);
  return true;
}

Expected<pivot::PivotTable, Error> read_pivot_table_definition(const std::vector<std::uint8_t>& definition_bytes) {
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_xml_buffer(doc, definition_bytes, "pivot_table_reader", "pivotTable*.xml"));
  pugi::xml_node root = doc.child("pivotTableDefinition");
  if (!root) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid, "pivotTable*.xml: missing <pivotTableDefinition> root",
                      "context=pivot_table_reader");
  }

  pivot::PivotTable table;
  if (pugi::xml_attribute name_attr = root.attribute("name"); name_attr) {
    table.set_name(name_attr.value());
  }
  // `dataCaption` is a required attribute; capture the authored value so
  // the writer re-emits it verbatim (default stays "Values" when absent).
  if (pugi::xml_attribute cap = root.attribute("dataCaption"); cap) {
    table.set_data_caption(cap.value());
  }
  table.set_pivot_cache_id(attr_u32(root, "cacheId"));

  // Grand-total layout flags. OOXML defaults both to true when absent, so a
  // file that turned grand totals OFF carries `rowGrandTotals="0"` /
  // `colGrandTotals="0"` explicitly. The model also defaults to true; we
  // read whatever is present so an OFF state survives the round trip
  // instead of silently flipping back to ON. The attributes are named after
  // the axis they total: `rowGrandTotals` is the grand total row at the
  // bottom (the model's column totals), `colGrandTotals` the column at the
  // right (the model's row totals).
  {
    const bool bottom_row = attr_bool(root, "rowGrandTotals", true);
    const bool right_col = attr_bool(root, "colGrandTotals", true);
    table.set_grand_totals(/*rows=*/right_col, /*cols=*/bottom_row);
    table.set_subtotal_hidden_items(attr_bool(root, "subtotalHiddenItems", false));
  }

  // Report layout mode, provisionally from the table-level defaults Excel
  // gives new fields: `compact` (default true) wins over `outline` (default
  // false). The row fields' own attributes override it below.
  {
    const bool compact = attr_bool(root, "compact", true);
    const bool outline = attr_bool(root, "outline", false);
    table.set_layout(compact ? pivot::PivotLayout::Compact
                             : (outline ? pivot::PivotLayout::Outline : pivot::PivotLayout::Tabular));
  }

  // Capture every root attribute the model does not represent structurally
  // (`updatedVersion`, `createdVersion`, `itemPrintTitles`, `indent`, ...)
  // so the writer re-emits them verbatim.
  if (pugi::xml_attribute caption = root.attribute("grandTotalCaption"); caption) {
    table.set_grand_total_caption(std::string(caption.value()));
  }
  table.set_auto_formatting(attr_bool(root, "useAutoFormatting", false));
  table.set_enable_drill(attr_bool(root, "enableDrill", true));
  table.set_show_error(attr_bool(root, "showError", false));
  if (pugi::xml_attribute attr = root.attribute("errorCaption"); attr) {
    table.set_error_caption(std::string(attr.value()));
  }
  table.set_show_missing(attr_bool(root, "showMissing", true));
  if (pugi::xml_attribute attr = root.attribute("missingCaption"); attr) {
    table.set_missing_caption(std::string(attr.value()));
  }
  table.set_show_headers(attr_bool(root, "showHeaders", true));
  table.set_show_drill(attr_bool(root, "showDrill", true));
  table.set_print_drill(attr_bool(root, "printDrill", false));
  table.set_show_data_tips(attr_bool(root, "showDataTips", true));
  table.set_grid_drop_zones(attr_bool(root, "gridDropZones", false));
  table.set_show_empty_row(attr_bool(root, "showEmptyRow", false));
  table.set_show_empty_col(attr_bool(root, "showEmptyCol", false));
  table.set_field_print_titles(attr_bool(root, "fieldPrintTitles", false));
  table.set_item_print_titles(attr_bool(root, "itemPrintTitles", false));
  table.set_show_data_drop_down(attr_bool(root, "showDataDropDown", true));
  table.set_multiple_field_filters(attr_bool(root, "multipleFieldFilters", true));
  table.set_preserve_formatting(attr_bool(root, "preserveFormatting", true));
  table.set_merge_item(attr_bool(root, "mergeItem", false));
  table.set_indent(attr_u32(root, "indent", 1U));
  table.set_page_over_then_down(attr_bool(root, "pageOverThenDown", false));
  table.set_page_wrap(attr_u32(root, "pageWrap", 0U));
  table.set_custom_list_sort(attr_bool(root, "customListSort", true));
  capture_unknown_attrs(root,
                        {"name",
                         "cacheId",
                         "dataCaption",
                         "rowGrandTotals",
                         "colGrandTotals",
                         "compact",
                         "outline",
                         "compactData",
                         "outlineData",
                         "subtotalHiddenItems",
                         "grandTotalCaption",
                         "useAutoFormatting",
                         "enableDrill",
                         "dataOnRows",
                         "dataPosition",
                         "showError",
                         "errorCaption",
                         "showMissing",
                         "missingCaption",
                         "showHeaders",
                         "showDrill",
                         "printDrill",
                         "showDataTips",
                         "gridDropZones",
                         "showEmptyRow",
                         "showEmptyCol",
                         "fieldPrintTitles",
                         "itemPrintTitles",
                         "showDataDropDown",
                         "multipleFieldFilters",
                         "preserveFormatting",
                         "mergeItem",
                         "indent",
                         "pageOverThenDown",
                         "pageWrap",
                         "customListSort"},
                        table.mutable_passthrough_attrs());

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
    return std::move(status.error());
  }
  // The span now describes a range Excel itself wrote. Recording that keeps
  // the writer from replacing it with our own layout projection, which would
  // make a read -> write cycle rewrite an authored `ref`.
  table.mark_span_authored();

  // Capture the `<location>` offset attributes. ECMA-376 requires
  // `firstHeaderRow`, `firstDataRow`, and `firstDataCol`; `rowPageCount`
  // and `colPageCount` are optional but commonly present. We preserve each
  // exactly as present (absent stays absent) so the writer re-emits a
  // schema-valid `<location>` on round trip. Dropping the required ones
  // makes Excel flag the file for repair and rebuild the pivot, losing
  // layout.
  auto opt_u32_attr = [](const pugi::xml_node& node, const char* name) -> std::optional<std::uint32_t> {
    const pugi::xml_attribute a = node.attribute(name);
    if (!a) {
      return std::nullopt;
    }
    return parse_xml_u32_attr(a, 0U);
  };
  table.set_location_attributes(opt_u32_attr(loc, "firstHeaderRow"), opt_u32_attr(loc, "firstDataRow"),
                                opt_u32_attr(loc, "firstDataCol"), opt_u32_attr(loc, "rowPageCount"),
                                opt_u32_attr(loc, "colPageCount"));

  if (pugi::xml_node fields = root.child("pivotFields"); fields) {
    RETURN_IF_ERROR(ParsePivotFields(fields, &table));
  }
  if (pugi::xml_node rows = root.child("rowFields"); rows) {
    std::optional<std::size_t> values_position;
    ParseFieldOrder(rows, &table.mutable_row_field_order(), &values_position);
    table.set_row_values_position(values_position);
  }
  // Excel renders the report form per field (`compact` and `outline` both
  // default to true); the first row field's pair decides the table's form.
  if (!table.row_field_order().empty()) {
    std::uint32_t index = 0;
    for (pugi::xml_node f = root.child("pivotFields").child("pivotField"); f;
         f = f.next_sibling("pivotField"), ++index) {
      if (index == table.row_field_order().front()) {
        const bool compact = attr_bool(f, "compact", true);
        const bool outline = attr_bool(f, "outline", true);
        table.set_layout(compact ? pivot::PivotLayout::Compact
                                 : (outline ? pivot::PivotLayout::Outline : pivot::PivotLayout::Tabular));
        break;
      }
    }
  }
  CaptureFormOverride(root, table.layout(), /*root=*/true, table.mutable_passthrough_attrs());
  {
    std::size_t index = 0;
    for (pugi::xml_node f = root.child("pivotFields").child("pivotField"); f && index < table.fields().size();
         f = f.next_sibling("pivotField"), ++index) {
      CaptureFormOverride(f, table.layout(), /*root=*/false, table.mutable_fields()[index].passthrough_attrs);
    }
  }
  if (pugi::xml_node cols = root.child("colFields"); cols) {
    std::optional<std::size_t> values_position;
    ParseFieldOrder(cols, &table.mutable_col_field_order(), &values_position);
    table.set_col_values_position(values_position);
  }
  // `<pageFields>` follows the same decode-but-do-not-own split as
  // `<filters>` below: it stays out of `kRecognized`, so the block still
  // lands in the pre-dataFields passthrough bin and the writer re-emits
  // the authored bytes. What the decode adds is the page-axis report
  // order and the selected item, neither of which `<pivotFields>` records.
  if (pugi::xml_node pages = root.child("pageFields"); pages) {
    ParsePageFields(pages, &table);
  }
  if (pugi::xml_node data = root.child("dataFields"); data) {
    if (auto status = ParseDataFields(data, &table); !status) {
      return std::move(status.error());
    }
  }
  // `<filters>` is owned by the model: the public entries and the retained
  // rest are written back from `active_filters()` / `retained_filters()`.
  // Decoded after the axis orders, which a value entry's axis follows.
  if (pugi::xml_node filters = root.child("filters"); filters) {
    ParseFilters(filters, &table);
  }
  // Capture any remaining direct children of <pivotTableDefinition> that
  // we do not model structurally into position-keyed raw-XML buffers, so
  // the writer can re-emit them at schema-valid slots.
  //
  // `CT_pivotTableDefinition` mandates a strict child order. `<rowItems>`,
  // `<colItems>`, and `<pageFields>` are the only unmodelled elements the
  // schema places *before* `<dataFields>`; they are binned by name into
  // the two pre-dataFields slots (rowItems after `<rowFields>`; colItems
  // and pageFields, each in a bin of its own, after `<colFields>`). Every
  // other unmodelled element (`<formats>`, `<conditionalFormats>`,
  // `<chartFormats>`, `<calculatedFields>`, `<calculatedItems>`,
  // `<pivotTableStyleInfo>`, `<extLst>`, ...) belongs after `<dataFields>`. `<pivotTableStyleInfo>`
  // and `<filters>` are decoded into the structured state; the
  // post-dataFields elements the schema places before the style go to their
  // own bin, the rest to the tail buffer. Relative order within each bin
  // follows document order.
  static const std::string_view kRecognized[] = {"location",  "pivotFields", "rowFields",
                                                 "colFields", "dataFields",  "filters"};
  for (pugi::xml_node child = root.first_child(); child; child = child.next_sibling()) {
    if (child.type() != pugi::node_element) {
      continue;
    }
    const std::string_view name = child.name();
    bool recognised = false;
    for (std::string_view r : kRecognized) {
      if (r == name) {
        recognised = true;
        break;
      }
    }
    if (recognised) {
      continue;
    }
    if (name == "pivotTableStyleInfo") {
      pivot::PivotStyleInfo info;
      info.name = child.attribute("name").as_string();
      info.show_row_headers = child.attribute("showRowHeaders").as_bool(false);
      info.show_col_headers = child.attribute("showColHeaders").as_bool(false);
      info.show_row_stripes = child.attribute("showRowStripes").as_bool(false);
      info.show_col_stripes = child.attribute("showColStripes").as_bool(false);
      info.show_last_column = child.attribute("showLastColumn").as_bool(false);
      table.set_style_info(std::move(info));
      continue;
    }
    std::string* bucket = nullptr;
    if (name == "rowItems") {
      bucket = &table.mutable_raw_passthrough_after_row_fields();
    } else if (name == "colItems") {
      bucket = &table.mutable_raw_col_items();
    } else if (name == "pageFields") {
      bucket = &table.mutable_raw_passthrough_after_col_fields();
    } else if (name == "formats" || name == "conditionalFormats" || name == "chartFormats" ||
               name == "pivotHierarchies") {
      bucket = &table.mutable_raw_passthrough_before_style_info();
    } else {
      bucket = &table.mutable_raw_passthrough_xml();
      if (name == "extLst") {
        ExtractAltText(child, &table);
      }
    }
    append_raw_xml(*bucket, child);
  }

  return table;
}

}  // namespace formulon::io
