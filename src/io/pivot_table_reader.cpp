//
// Implementation of the pivot-table-definition reader. See
// pivot_table_reader.h for the public contract.

#include "io/pivot_table_reader.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstring>
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
      field->marker_items.push_back(pivot::PivotMarkerItem{raw_xml(it), std::string(t), ordinal});
      continue;
    }
    pivot::PivotItem entry;
    entry.visible = !parse_xml_bool_attr(it.attribute("h"));
    capture_unknown_attrs(it, {"x", "h", "t"}, entry.passthrough_attrs);
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
    // Preserve unmodelled `<pivotField>` attributes (`showAll`,
    // `includeNewItemsInFilter`, ...) and children for verbatim round-trip.
    // `compact` / `outline` are the report form, which the writer derives
    // from the table's layout; `CaptureFormOverride` keeps them once the
    // layout is known and they depart from it.
    capture_unknown_attrs(f, {"axis",          "dataField",      "name",
                              "numFmtId",      "subtotalTop",    "defaultSubtotal",
                              "sumSubtotal",   "countASubtotal", "avgSubtotal",
                              "maxSubtotal",   "minSubtotal",    "productSubtotal",
                              "countSubtotal", "stdDevSubtotal", "stdDevPSubtotal",
                              "varSubtotal",   "varPSubtotal",   "sortType",
                              "compact",       "outline"},
                          field.passthrough_attrs);
    capture_unknown_children(f, {"items"}, field.raw_children_xml);
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

/// Walks `<dataFields>` in document order. Returns `kIoSheetCorrupt` on
/// the first `<dataField>` that lacks a `name` attribute, since that
/// name is the GETPIVOTDATA lookup key.
Expected<void, Error> ParseDataFields(const pugi::xml_node& parent, pivot::PivotTable* out) {
  for (pugi::xml_node df = parent.child("dataField"); df; df = df.next_sibling("dataField")) {
    pivot::PivotDataField entry;
    pugi::xml_attribute name_attr = df.attribute("name");
    if (!name_attr || name_attr.value() == nullptr || *name_attr.value() == '\0') {
      return make_error(FormulonErrorCode::kIoSheetCorrupt,
                        "pivotTableDefinition: <dataField> missing required name attribute",
                        "context=pivot_table_reader");
    }
    entry.name = name_attr.value();
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

/// Maps a `<filter type="...">` body to a `CaptionPredicate`.
///
/// Returns `nullopt` for every non-caption member of `ST_PivotFilterType`
/// (the value, date, and relative-period families, plus `unknown`). Those
/// stay passthrough-only: the block is still re-emitted verbatim, it just
/// does not take part in evaluation.
std::optional<pivot::CaptionPredicate> ParseCaptionPredicate(std::string_view text) {
  struct Entry {
    std::string_view name;
    pivot::CaptionPredicate predicate;
  };
  static constexpr Entry kEntries[] = {
      {"captionEqual", pivot::CaptionPredicate::Equal},
      {"captionNotEqual", pivot::CaptionPredicate::NotEqual},
      {"captionBeginsWith", pivot::CaptionPredicate::BeginsWith},
      {"captionNotBeginsWith", pivot::CaptionPredicate::NotBeginsWith},
      {"captionEndsWith", pivot::CaptionPredicate::EndsWith},
      {"captionNotEndsWith", pivot::CaptionPredicate::NotEndsWith},
      {"captionContains", pivot::CaptionPredicate::Contains},
      {"captionNotContains", pivot::CaptionPredicate::NotContains},
      {"captionGreaterThan", pivot::CaptionPredicate::GreaterThan},
      {"captionGreaterThanOrEqual", pivot::CaptionPredicate::GreaterThanOrEqual},
      {"captionLessThan", pivot::CaptionPredicate::LessThan},
      {"captionLessThanOrEqual", pivot::CaptionPredicate::LessThanOrEqual},
      {"captionBetween", pivot::CaptionPredicate::Between},
      {"captionNotBetween", pivot::CaptionPredicate::NotBetween},
  };
  for (const Entry& e : kEntries) {
    if (e.name == text) {
      return e.predicate;
    }
  }
  return std::nullopt;
}

/// Collects the criterion strings a `<filter>`'s nested `<autoFilter>`
/// carries, in document order.
///
/// Two shapes reach the same place: `<customFilters>` holds one
/// `<customFilter val="...">` for a single-sided comparison and two for a
/// `Between`, while a plain equality is also spelled as a `<filters>`
/// list of `<filter val="...">`. Both are read; anything else (a
/// `<top10>` or a `<dynamicFilter>`) contributes no criterion and leaves
/// the entry to be dropped by the caller.
std::vector<std::string_view> CollectFilterCriteria(const pugi::xml_node& filter_node) {
  std::vector<std::string_view> out;
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
  if (pugi::xml_node plain = column.child("filters"); plain) {
    for (pugi::xml_node f = plain.child("filter"); f; f = f.next_sibling("filter")) {
      if (pugi::xml_attribute val = f.attribute("val"); val) {
        out.emplace_back(val.value());
      }
    }
  }
  return out;
}

/// Maps a `<filter type="...">` body to the value or date family it
/// evaluates as.
///
/// Only the members with an evaluation counterpart are mapped. Excel's
/// "top ten" dialog writes `count` for an item count, and `percent` /
/// `sum` for its other two flavours; those two rank on a different
/// quantity and are left undecoded rather than silently answered as an
/// item count. The relative-period families resolve against the current
/// date and are excluded on purpose — see `PivotTable`.
std::optional<pivot::FilterType> ParseValueFilterType(std::string_view text) {
  struct Entry {
    std::string_view name;
    pivot::FilterType type;
  };
  static constexpr Entry kEntries[] = {
      {"count", pivot::FilterType::ValueTop10},
      {"percent", pivot::FilterType::ValueTop10},
      {"sum", pivot::FilterType::ValueTop10},
      {"valueGreaterThan", pivot::FilterType::ValueGreaterThan},
      {"valueEqual", pivot::FilterType::ValueEqual},
      {"valueNotEqual", pivot::FilterType::ValueNotEqual},
      {"valueGreaterThanOrEqual", pivot::FilterType::ValueGreaterThanOrEqual},
      {"valueLessThan", pivot::FilterType::ValueLessThan},
      {"valueLessThanOrEqual", pivot::FilterType::ValueLessThanOrEqual},
      {"valueBetween", pivot::FilterType::ValueBetween},
      {"valueNotBetween", pivot::FilterType::ValueNotBetween},
      {"dateBetween", pivot::FilterType::LabelDate},
  };
  for (const Entry& entry : kEntries) {
    if (entry.name == text) {
      return entry.type;
    }
  }
  return std::nullopt;
}

/// Reads the item count out of a top-N filter's nested
/// `<top10 val="N"/>`. Returns `nullopt` when the element is absent, so
/// the entry is dropped rather than defaulting to a count nobody wrote.
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

/// Reads a top-N filter's `<top10 top="...">` direction. Defaults to
/// `true` (Top N) when the element or attribute is absent, matching the
/// schema default; `top="0"` is Excel's "Bottom N" dialog option.
bool CollectTopDirection(const pugi::xml_node& filter_node) {
  const pugi::xml_node top = filter_node.child("autoFilter").child("filterColumn").child("top10");
  return attr_bool(top, "top", true);
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

/// Which quantity a top-N entry counts out.
///
/// The three flavours share one element shape, so only the type name
/// separates them: `<top10 val="100"/>` under `sum` and under `count` are
/// byte-identical. Anything that is not a top-N type never reaches here.
pivot::TopNBasis ParseTopNBasis(std::string_view text) {
  if (text == "percent") {
    return pivot::TopNBasis::Percent;
  }
  if (text == "sum") {
    return pivot::TopNBasis::Sum;
  }
  return pivot::TopNBasis::Items;
}

/// Maps a relative-period `<filter type>` onto its evaluation counterpart.
///
/// These types carry no criteria at all — the window is implied by the
/// name — so the mapping is the entire decode. `M1`..`M12` / `Q1`..`Q4`
/// are absent because they select a recurring period rather than one
/// contiguous window; `ParseRecurringMonths` decodes those instead.
std::optional<pivot::RelativePeriod> ParseRelativePeriod(std::string_view text) {
  struct Entry {
    std::string_view name;
    pivot::RelativePeriod period;
  };
  static constexpr Entry kEntries[] = {
      {"today", pivot::RelativePeriod::Today},
      {"yesterday", pivot::RelativePeriod::Yesterday},
      {"tomorrow", pivot::RelativePeriod::Tomorrow},
      {"thisMonth", pivot::RelativePeriod::ThisMonth},
      {"lastMonth", pivot::RelativePeriod::LastMonth},
      {"nextMonth", pivot::RelativePeriod::NextMonth},
      {"thisQuarter", pivot::RelativePeriod::ThisQuarter},
      {"lastQuarter", pivot::RelativePeriod::LastQuarter},
      {"nextQuarter", pivot::RelativePeriod::NextQuarter},
      {"thisYear", pivot::RelativePeriod::ThisYear},
      {"lastYear", pivot::RelativePeriod::LastYear},
      {"nextYear", pivot::RelativePeriod::NextYear},
      {"yearToDate", pivot::RelativePeriod::YearToDate},
      {"thisWeek", pivot::RelativePeriod::ThisWeek},
      {"lastWeek", pivot::RelativePeriod::LastWeek},
      {"nextWeek", pivot::RelativePeriod::NextWeek},
  };
  for (const Entry& entry : kEntries) {
    if (entry.name == text) {
      return entry.period;
    }
  }
  return std::nullopt;
}

/// Maps a recurring `<filter type>` — `M1`..`M12`, `Q1`..`Q4` — onto the
/// inclusive calendar-month range it selects.
///
/// These name a position in the year rather than a window: `M1` keeps
/// every January of every year, so the decode is year-free and needs no
/// clock reading. A quarter is three adjacent months, which is why one
/// month range covers both families.
std::optional<std::pair<unsigned, unsigned>> ParseRecurringMonths(std::string_view text) {
  if (text.size() < 2U) {
    return std::nullopt;
  }
  const char family = text.front();
  if (family != 'M' && family != 'Q') {
    return std::nullopt;
  }
  unsigned index = 0;
  for (const char digit : text.substr(1)) {
    if (digit < '0' || digit > '9') {
      return std::nullopt;
    }
    index = index * 10U + static_cast<unsigned>(digit - '0');
    if (index > 12U) {
      return std::nullopt;  // Longer than any valid selector; stop before overflow.
    }
  }
  if (family == 'M') {
    if (index < 1U || index > 12U) {
      return std::nullopt;
    }
    return std::pair<unsigned, unsigned>{index, index};
  }
  if (index < 1U || index > 4U) {
    return std::nullopt;
  }
  const unsigned first = (index - 1U) * 3U + 1U;
  return std::pair<unsigned, unsigned>{first, first + 2U};
}

/// Decodes a recurring-period `<filter>` (`M1`..`M12`, `Q1`..`Q4`).
std::optional<pivot::AuthoredRecurringFilter> DecodeRecurringFilter(const pugi::xml_node& node) {
  const auto months_or = ParseRecurringMonths(attr_str(node, "type"));
  if (!months_or) {
    return std::nullopt;
  }
  pivot::AuthoredRecurringFilter entry;
  entry.field_index = attr_u32(node, "fld");
  entry.month_low = months_or->first;
  entry.month_high = months_or->second;
  return entry;
}

/// Decodes a relative-period `<filter>` (`thisMonth`, `yearToDate`, ...).
std::optional<pivot::AuthoredPeriodFilter> DecodePeriodFilter(const pugi::xml_node& node) {
  const auto period_or = ParseRelativePeriod(attr_str(node, "type"));
  if (!period_or) {
    return std::nullopt;
  }
  pivot::AuthoredPeriodFilter entry;
  entry.field_index = attr_u32(node, "fld");
  entry.period = *period_or;
  return entry;
}

/// Decodes a value or date `<filter>`.
///
/// Unlike a caption filter, these carry no `stringValue*` shorthand: the
/// criteria live only in the nested `<autoFilter>`, as one
/// `<customFilter val="...">` for a threshold and two for a range. Date
/// bounds are written as serials, which is the form the evaluation pass
/// wants anyway — the rendered label would lose both precision and
/// locale.
std::optional<pivot::AuthoredValueFilter> DecodeValueFilter(const pugi::xml_node& node) {
  const auto type_or = ParseValueFilterType(attr_str(node, "type"));
  if (!type_or) {
    return std::nullopt;
  }
  pivot::AuthoredValueFilter entry;
  entry.field_index = attr_u32(node, "fld");
  entry.data_field_index = attr_u32(node, "iMeasureFld");
  entry.type = *type_or;
  entry.top_n_basis = ParseTopNBasis(attr_str(node, "type"));

  if (*type_or == pivot::FilterType::ValueTop10) {
    const auto count_or = CollectTopCount(node);
    if (!count_or) {
      return std::nullopt;
    }
    entry.value = *count_or;
    entry.top = CollectTopDirection(node);
    return entry;
  }
  const bool ranged = *type_or == pivot::FilterType::ValueBetween || *type_or == pivot::FilterType::ValueNotBetween ||
                      *type_or == pivot::FilterType::LabelDate;
  const std::vector<std::string_view> criteria = CollectFilterCriteria(node);
  if (criteria.size() < (ranged ? 2U : 1U)) {
    return std::nullopt;
  }
  const auto low_or = ParseCriterionNumber(criteria[0]);
  if (!low_or) {
    return std::nullopt;
  }
  entry.value = *low_or;
  if (ranged) {
    const auto high_or = ParseCriterionNumber(criteria[1]);
    if (!high_or) {
      return std::nullopt;
    }
    entry.value_high = *high_or;
  }
  return entry;
}

/// Decodes a caption `<filter>`.
///
/// An entry is dropped when its type is outside the caption family or when
/// it carries no criterion (`Between` needs two, everything else one).
///
/// Excel writes the criterion twice. `stringValue1` / `stringValue2` on
/// the `<filter>` element carry it as entered, while the nested
/// `<autoFilter>` repeats it as a wildcard pattern (`v*`, `*v`, `*v*`).
/// For begins / ends / contains Excel evaluates that pattern on refresh,
/// not `stringValue1`, so the pattern is the decoded value and
/// `stringValue1` is the fallback, wrapped into the same pattern. The
/// other predicates prefer the direct attributes, falling back to the
/// nested criterion.
std::optional<pivot::AuthoredCaptionFilter> DecodeCaptionFilter(const pugi::xml_node& f) {
  const auto predicate_or = ParseCaptionPredicate(attr_str(f, "type"));
  if (!predicate_or) {
    return std::nullopt;
  }
  const bool ranged =
      *predicate_or == pivot::CaptionPredicate::Between || *predicate_or == pivot::CaptionPredicate::NotBetween;
  const pugi::xml_attribute direct_low = f.attribute("stringValue1");
  const pugi::xml_attribute direct_high = f.attribute("stringValue2");
  const bool direct = direct_low && (!ranged || direct_high);
  std::optional<std::string> direct_pattern = pivot::caption_wildcard_pattern(*predicate_or, direct_low.value());

  std::string low;
  std::string high;
  if (direct_pattern) {
    const std::vector<std::string_view> criteria = CollectFilterCriteria(f);
    if (!criteria.empty()) {
      low.assign(criteria[0]);
    } else if (direct_low) {
      low = std::move(*direct_pattern);
    } else {
      return std::nullopt;
    }
  } else if (direct) {
    low.assign(direct_low.value());
    if (ranged) {
      high.assign(direct_high.value());
    }
  } else {
    const std::vector<std::string_view> criteria = CollectFilterCriteria(f);
    if (criteria.size() < (ranged ? 2U : 1U)) {
      return std::nullopt;
    }
    low.assign(criteria[0]);
    if (ranged) {
      high.assign(criteria[1]);
    }
  }

  pivot::AuthoredCaptionFilter entry;
  entry.field_index = attr_u32(f, "fld");
  entry.predicate = *predicate_or;
  entry.value = std::move(low);
  entry.value_high = std::move(high);
  return entry;
}

/// A top-N count as the public `PivotFilter` stores it: an `int` when the
/// file spells a whole number in range, so it reads back the way the API
/// takes it.
std::variant<int, double, std::string> TopCountValue(double count) {
  if (count >= 0.0 && count <= static_cast<double>(std::numeric_limits<int>::max()) &&
      count == static_cast<double>(static_cast<int>(count))) {
    return static_cast<int>(count);
  }
  return count;
}

/// Decodes a `<filter>` into the public `PivotFilter` vocabulary, or returns
/// `nullopt` when the entry needs something that vocabulary lacks.
///
/// Only an exact match is taken, so the entry evaluates the same either way:
/// a top-N must count items from the top, and a contains / begins-with entry
/// must carry a `stringValue1` whose pattern is the nested criterion. A value
/// entry must sit on the row or column axis, which is where both passes look
/// for it. `field_name` is left empty and filled from the cache by
/// `pivot::resolve_pivot_names`.
std::optional<pivot::PivotFilter> DecodePublicFilter(const pugi::xml_node& node, const pivot::PivotTable& table) {
  const std::optional<std::uint32_t> fld = parse_xml_u32_attr_strict(node.attribute("fld"));
  if (!fld.has_value() || *fld >= table.fields().size()) {
    return std::nullopt;
  }
  const auto on = [&](const std::vector<std::uint32_t>& order) {
    return std::find(order.begin(), order.end(), *fld) != order.end();
  };
  const std::optional<pivot::PivotAxis> report_axis =
      on(table.row_field_order())   ? std::optional<pivot::PivotAxis>(pivot::PivotAxis::Row)
      : on(table.col_field_order()) ? std::optional<pivot::PivotAxis>(pivot::PivotAxis::Col)
                                    : std::nullopt;
  pivot::PivotFilter out;
  out.data_field_index = attr_u32(node, "iMeasureFld");
  const std::string_view type = attr_str(node, "type");
  const bool caption = type == "captionContains" || type == "captionBeginsWith";
  if (caption) {
    const pugi::xml_attribute operand = node.attribute("stringValue1");
    if (!operand) {
      return std::nullopt;
    }
    const pivot::CaptionPredicate predicate =
        type == "captionContains" ? pivot::CaptionPredicate::Contains : pivot::CaptionPredicate::BeginsWith;
    const std::vector<std::string_view> criteria = CollectFilterCriteria(node);
    if (!criteria.empty() && criteria[0] != *pivot::caption_wildcard_pattern(predicate, operand.value())) {
      return std::nullopt;
    }
    out.type = predicate == pivot::CaptionPredicate::Contains ? pivot::FilterType::LabelContains
                                                              : pivot::FilterType::LabelBeginsWith;
    out.value = std::string(operand.value());
  } else {
    const std::optional<pivot::AuthoredValueFilter> decoded = DecodeValueFilter(node);
    if (!decoded) {
      return std::nullopt;
    }
    switch (decoded->type) {
      case pivot::FilterType::ValueTop10:
        if (decoded->top_n_basis != pivot::TopNBasis::Items || !decoded->top) {
          return std::nullopt;
        }
        out.value = TopCountValue(decoded->value);
        break;
      case pivot::FilterType::ValueGreaterThan:
        out.value = decoded->value;
        break;
      case pivot::FilterType::ValueBetween:
      case pivot::FilterType::LabelDate:
        out.value = decoded->value;
        out.value_high = *decoded->value_high;
        break;
      default:
        return std::nullopt;
    }
    out.type = decoded->type;
  }
  if (pivot::is_value_filter_type(out.type)) {
    if (!report_axis) {
      return std::nullopt;
    }
    out.axis = *report_axis;
  } else {
    const pivot::PivotAxis field_axis = table.fields()[*fld].axis;
    out.axis = report_axis.value_or(field_axis == pivot::PivotAxis::None ? pivot::PivotAxis::Row : field_axis);
  }
  out.source = pivot::PivotFilterSource{raw_xml(node), *fld, out.data_field_index, attr_u32(node, "id")};
  return out;
}

/// Splits a `<filters>` block between the public `active_filters()` and the
/// verbatim `retained_filters()`, decoding the evaluable retained entries
/// into the matching `authored_*` list. The authored type sets are disjoint,
/// so at most one decoder claims an entry; one no decoder claims is kept
/// for the round trip only.
void ParseFilters(const pugi::xml_node& parent, pivot::PivotTable* out) {
  for (pugi::xml_node node = parent.child("filter"); node; node = node.next_sibling("filter")) {
    if (std::optional<pivot::PivotFilter> entry = DecodePublicFilter(node, *out)) {
      out->mutable_active_filters().push_back(std::move(*entry));
      continue;
    }
    const std::size_t active_before = out->active_filters().size();
    out->mutable_retained_filters().push_back(
        pivot::RetainedPivotFilter{raw_xml(node), active_before, attr_u32(node, "id")});
    if (std::optional<pivot::AuthoredCaptionFilter> caption = DecodeCaptionFilter(node)) {
      out->mutable_authored_caption_filters().push_back(std::move(*caption));
    } else if (std::optional<pivot::AuthoredValueFilter> value = DecodeValueFilter(node)) {
      value->active_before = active_before;
      out->mutable_authored_value_filters().push_back(*value);
    } else if (std::optional<pivot::AuthoredPeriodFilter> period = DecodePeriodFilter(node)) {
      out->mutable_authored_period_filters().push_back(*period);
    } else if (std::optional<pivot::AuthoredRecurringFilter> recurring = DecodeRecurringFilter(node)) {
      out->mutable_authored_recurring_filters().push_back(*recurring);
    }
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
  // instead of silently flipping back to ON.
  {
    const bool row_grand = attr_bool(root, "rowGrandTotals", true);
    const bool col_grand = attr_bool(root, "colGrandTotals", true);
    table.set_grand_totals(row_grand, col_grand);
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
  capture_unknown_attrs(root,
                        {"name", "cacheId", "dataCaption", "rowGrandTotals", "colGrandTotals", "compact", "outline",
                         "compactData", "outlineData", "subtotalHiddenItems"},
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
  // and pageFields after `<colFields>`). Every other unmodelled element
  // (`<formats>`, `<conditionalFormats>`, `<chartFormats>`,
  // `<calculatedFields>`, `<calculatedItems>`, `<pivotTableStyleInfo>`,
  // `<extLst>`, ...) belongs after `<dataFields>`. `<pivotTableStyleInfo>`
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
    } else if (name == "colItems" || name == "pageFields") {
      bucket = &table.mutable_raw_passthrough_after_col_fields();
    } else if (name == "formats" || name == "conditionalFormats" || name == "chartFormats" ||
               name == "pivotHierarchies") {
      bucket = &table.mutable_raw_passthrough_before_style_info();
    } else {
      bucket = &table.mutable_raw_passthrough_xml();
    }
    append_raw_xml(*bucket, child);
  }

  return table;
}

}  // namespace formulon::io
