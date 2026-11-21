//
// Decodes a pivot table's `<filters>` block into the model's active and
// retained filter lists.

#include <algorithm>
#include <cstddef>
#include <iterator>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <variant>
#include <vector>

#include "io/pivot_attr_names.h"
#include "io/pivot_table_reader.h"
#include "io/xml_utils.h"
#include "io/xsd_double.h"
#include "pivot/filter_engine.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pugixml.hpp"

namespace formulon::io {
namespace {

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
  out.literal = out.type == pivot::FilterType::LabelEqual && plain && !criteria.empty();
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

}  // namespace

void parse_pivot_filters(const pugi::xml_node& parent, pivot::PivotTable* out) {
  for (pugi::xml_node node = parent.child("filter"); node; node = node.next_sibling("filter")) {
    if (std::optional<pivot::PivotFilter> entry = DecodeFilter(node, *out)) {
      out->mutable_active_filters().push_back(std::move(*entry));
      continue;
    }
    out->mutable_retained_filters().push_back(
        pivot::RetainedPivotFilter{raw_xml(node), out->active_filters().size(), attr_u32(node, "id")});
  }
}

bool read_pivot_filters_xml(std::string_view filters_xml, pivot::PivotTable* table) {
  pugi::xml_document doc;
  if (!doc.load_buffer(filters_xml.data(), filters_xml.size(), pugi::parse_default, pugi::encoding_utf8)) {
    return false;
  }
  parse_pivot_filters(doc.child("filters"), table);
  return true;
}

}  // namespace formulon::io
