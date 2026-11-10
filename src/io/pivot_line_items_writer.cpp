//
// `<rowItems>` / `<colItems>` serialization: each line's `t`, `r`, `i` in that
// order, each omitted at its default, and `<x>` with `v` omitted at 0, as
// Excel writes them.
//
// @size-budget: 12 KB

#include "io/pivot_line_items_writer.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <map>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "io/pivot_attr_names.h"
#include "io/pivot_items_binding.h"
#include "io/xml_utils.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_types.h"
#include "pivot/value_order.h"
#include "value.h"

namespace formulon::io {
namespace {

using PositionMap = std::map<Value, std::uint32_t, pivot::ValueLess>;

static_assert(static_cast<std::size_t>(pivot::PivotLineType::kVarP) -
                      static_cast<std::size_t>(pivot::PivotLineType::kSum) + 1U ==
                  kPivotSubtotalNames.size(),
              "the custom subtotal line types follow kPivotSubtotalNames");

/// The `ST_ItemType` token of a line; empty for `data`, which writes no `t`.
std::string_view LineTypeToken(pivot::PivotLineType type) {
  switch (type) {
    case pivot::PivotLineType::kData:
      return {};
    case pivot::PivotLineType::kDefault:
      return "default";
    case pivot::PivotLineType::kGrand:
      return "grand";
    case pivot::PivotLineType::kBlank:
      return "blank";
    default:
      return kPivotSubtotalNames[static_cast<std::size_t>(type) - static_cast<std::size_t>(pivot::PivotLineType::kSum)]
          .item;
  }
}

/// True for a marker `AppendPivotField` writes after every item: a subtotal or `grand`.
bool ClosesItemList(std::string_view type) {
  if (type == "default" || type == "grand") {
    return true;
  }
  return std::any_of(kPivotSubtotalNames.begin(), kPivotSubtotalNames.end(),
                     [type](const PivotSubtotalName& name) { return name.item == type; });
}

/// Position of each identified item in the field's written `<items>`. Two
/// items one identity cannot tell apart (a date and the number with its
/// serial) get no position, so a line naming either cannot be encoded.
PositionMap BuildPositions(const pivot::PivotTable& table, const pivot::PivotCache& cache,
                           const PivotCacheWritePlan* plan, std::uint32_t field_index) {
  const pivot::PivotField& field = table.fields()[field_index];
  const PivotCacheFieldWritePlan* fp =
      plan != nullptr && field_index < plan->fields.size() ? &plan->fields[field_index] : nullptr;
  const bool derived = fp != nullptr && fp->derived;
  std::vector<bool> by_position;
  const std::vector<pivot::PivotItem> items = detail::BoundItems(field, fp, &by_position);
  PositionMap positions;
  std::vector<Value> ambiguous;
  for (std::size_t i = 0; i < items.size(); ++i) {
    // `blank` markers anchored at or before the item are written ahead of it.
    std::size_t markers = 0;
    for (const pivot::PivotMarkerItem& marker : field.marker_items) {
      markers += !ClosesItemList(marker.type) && std::min(marker.after_items, items.size()) <= i ? 1U : 0U;
    }
    const pivot::PivotItem& item = items[i];
    std::optional<Value> identity;
    if (!derived) {
      identity = pivot::pivot_item_record_value(cache, field_index, item);
    } else if (!by_position[i] && item.cache_index < fp->items.size()) {
      identity = fp->items[item.cache_index];
    }
    if (identity.has_value() && !positions.emplace(*identity, static_cast<std::uint32_t>(i + markers)).second) {
      ambiguous.push_back(*identity);
    }
  }
  for (const Value& identity : ambiguous) {
    positions.erase(identity);
  }
  return positions;
}

/// Appends one `<rowItems>` / `<colItems>` element; false when an item has no position.
bool AppendLines(std::string& out, std::string_view tag, const std::vector<pivot::PivotLine>& lines,
                 const pivot::PivotTable& table, const pivot::PivotCache& cache, const PivotCacheWritePlan* plan,
                 std::vector<std::optional<PositionMap>>& positions) {
  out.push_back('<');
  out.append(tag);
  append_xml_attr_uint(out, "count", lines.size());
  out.push_back('>');
  for (const pivot::PivotLine& line : lines) {
    out.append("<i");
    if (const std::string_view token = LineTypeToken(line.type); !token.empty()) {
      append_xml_attr(out, "t", token);
    }
    append_xml_attr_uint_if(out, "r", line.r, 0U);
    append_xml_attr_uint_if(out, "i", line.i, 0U);
    if (line.xs.empty()) {
      out.append("/>");
      continue;
    }
    out.push_back('>');
    for (const pivot::PivotLineX& x : line.xs) {
      std::uint32_t v = 0;
      if (x.values_level) {
        v = x.data_index;
      } else if (x.node != nullptr) {
        if (x.field_index >= table.fields().size() || !x.node->identity.has_value()) {
          return false;
        }
        std::optional<PositionMap>& field_positions = positions[x.field_index];
        if (!field_positions.has_value()) {
          field_positions = BuildPositions(table, cache, plan, x.field_index);
        }
        const auto it = field_positions->find(*x.node->identity);
        if (it == field_positions->end()) {
          return false;
        }
        v = it->second;
      }
      out.append("<x");
      append_xml_attr_uint_if(out, "v", v, 0U);
      out.append("/>");
    }
    out.append("</i>");
  }
  out.append("</");
  out.append(tag);
  out.push_back('>');
  return true;
}

}  // namespace

std::optional<PivotItemsXml> write_pivot_line_items(const pivot::PivotTable& table, const pivot::PivotCache& cache,
                                                    const PivotCacheWritePlan* plan,
                                                    const pivot::PivotLineItems& lines) {
  if (lines.row_refusal.has_value() || lines.col_refusal.has_value()) {
    return std::nullopt;
  }
  std::vector<std::optional<PositionMap>> positions(table.fields().size());
  PivotItemsXml xml;
  if (!AppendLines(xml.row_items, "rowItems", lines.rows, table, cache, plan, positions) ||
      !AppendLines(xml.col_items, "colItems", lines.cols, table, cache, plan, positions)) {
    return std::nullopt;
  }
  return xml;
}

}  // namespace formulon::io
