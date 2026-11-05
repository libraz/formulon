//
// The `<filters>` block of a pivot table definition: public filter entries
// serialised from the model, interleaved with retained source filters.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <variant>
#include <vector>

#include "io/pivot_attr_names.h"
#include "io/pivot_table_writer_internal.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "io/xsd_double.h"
#include "pivot/field_lookup.h"
#include "pivot/filter_engine.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "utils/strings.h"

namespace formulon::io {
namespace detail {

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

}  // namespace detail

namespace {

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
  const std::optional<double> low = pivot::payload_number(filter.value);
  const std::optional<double> high = pivot::payload_number(filter.value_high);
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

}  // namespace

namespace detail {

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

}  // namespace detail
}  // namespace formulon::io
