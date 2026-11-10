// @size-budget: 6 KB
//
// Remapping a pivot table onto a rebuilt cache. See the header.

#include "pivot/field_remap.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "pivot/field_grouping.h"
#include "pivot/field_lookup.h"
#include "pivot/filter_engine.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/value_order.h"
#include "utils/index_sort.h"
#include "value.h"

namespace formulon::pivot {
namespace {

using IndexMap = std::vector<std::optional<std::uint32_t>>;
// Per old position in a field's items, its new position; empty when removed.
using PositionMap = std::vector<std::optional<std::size_t>>;

std::optional<std::uint32_t> mapped(const IndexMap& map, std::size_t index) {
  return index < map.size() ? map[index] : std::nullopt;
}

bool flag_at(const std::vector<bool>& flags, std::size_t i) {
  return i < flags.size() && flags[i];
}

// Sets or clears the `name="1"` passthrough attribute.
void set_flag_attr(std::vector<std::pair<std::string, std::string>>& attrs, std::string_view name, bool on) {
  attrs.erase(std::remove_if(attrs.begin(), attrs.end(), [name](const auto& attr) { return attr.first == name; }),
              attrs.end());
  if (on) {
    attrs.emplace_back(std::string(name), "1");
  }
}

// The value item `cache_index` of `cache_field` stands for, for sorting.
Value item_value(const PivotCacheField& cache_field, std::uint32_t cache_index) {
  if (is_group_field(cache_field)) {
    const std::vector<Value>& items = cache_field.field_group->group_items;
    return cache_index < items.size() ? items[cache_index] : Value::blank();
  }
  return cache_index < cache_field.shared_items.size() ? cache_field.shared_items[cache_index] : Value::blank();
}

// Rewrites `field`'s items through `map` onto `cache_field`, filling `positions`.
void remap_items(PivotField& field, const IndexMap& map, const PivotCacheField& cache_field, PositionMap& positions) {
  if (field.items.empty()) {
    return;
  }
  const bool grouped = is_group_field(cache_field);
  const std::size_t item_count =
      grouped ? cache_field.field_group->group_items.size() : cache_field.shared_items.size();
  std::vector<bool> named(item_count, false);
  std::vector<PivotItem> kept;
  positions.assign(field.items.size(), std::nullopt);
  for (std::size_t k = 0; k < field.items.size(); ++k) {
    PivotItem& item = field.items[k];
    const std::optional<std::uint32_t> target = mapped(map, item.has_cache_index ? item.cache_index : k);
    if (!target.has_value() || *target >= item_count) {
      continue;
    }
    item.has_cache_index = true;
    item.cache_index = *target;
    if (!grouped) {
      set_flag_attr(item.passthrough_attrs, "m", flag_at(cache_field.shared_item_is_missing, *target));
      if (!item.name.empty()) {
        item.name = display_string(cache_field.shared_items[*target]);
      }
    }
    named[*target] = true;
    positions[k] = kept.size();
    kept.push_back(std::move(item));
  }
  const bool hides = std::any_of(kept.begin(), kept.end(), [](const PivotItem& item) { return !item.visible; });
  for (std::uint32_t n = 0; n < item_count; ++n) {
    if (named[n] || flag_at(cache_field.shared_item_is_missing, n) ||
        flag_at(cache_field.shared_item_is_calculated, n)) {
      continue;
    }
    PivotItem added;
    added.has_cache_index = true;
    added.cache_index = n;
    added.visible = !hides;
    if (!grouped) {
      added.name = display_string(cache_field.shared_items[n]);
    }
    kept.push_back(std::move(added));
  }

  if (!grouped && !field.sort.manual && field.sort.stated && field.sort.by_field.empty()) {
    const bool ascending = field.sort.ascending;
    const auto less = [&](std::uint32_t a, std::uint32_t b) {
      const Value va = item_value(cache_field, kept[a].cache_index);
      const Value vb = item_value(cache_field, kept[b].cache_index);
      return ascending ? value_less(va, vb) : value_less(vb, va);
    };
    std::vector<std::uint32_t> order;
    sorted_index_order(order, static_cast<std::uint32_t>(kept.size()), make_index_less(less));
    std::vector<std::size_t> rank(kept.size());
    std::vector<PivotItem> sorted;
    sorted.reserve(kept.size());
    for (const std::uint32_t i : order) {
      rank[i] = sorted.size();
      sorted.push_back(std::move(kept[i]));
    }
    for (std::optional<std::size_t>& position : positions) {
      if (position.has_value()) {
        position = rank[*position];
      }
    }
    kept = std::move(sorted);
  }

  for (PivotMarkerItem& marker : field.marker_items) {
    marker.after_items = std::min(marker.after_items, kept.size());
  }
  field.items = std::move(kept);
}

// The new position of item position `item` of old field `field`: through the
// field's item positions when its items were rebuilt, else through the item
// map (an implicit item list is in shared-item order).
std::optional<std::uint32_t> remap_item_position(const FieldRemap& remap, const std::vector<PositionMap>& positions,
                                                 std::size_t field, std::uint32_t item) {
  if (field < positions.size() && !positions[field].empty()) {
    const std::optional<std::size_t> moved = item < positions[field].size() ? positions[field][item] : std::nullopt;
    return moved.has_value() ? std::optional<std::uint32_t>(static_cast<std::uint32_t>(*moved)) : std::nullopt;
  }
  if (field < remap.items.size() && !remap.items[field].empty()) {
    return mapped(remap.items[field], item);
  }
  return item;
}

// Remaps an axis order, returning the new count of fields before `marker`.
std::optional<std::size_t> remap_order(std::vector<std::uint32_t>& order, const IndexMap& fields,
                                       std::optional<std::size_t> marker) {
  std::vector<std::uint32_t> kept;
  std::optional<std::size_t> moved_marker = marker;
  for (std::size_t i = 0; i < order.size(); ++i) {
    if (marker.has_value() && i == *marker) {
      moved_marker = kept.size();
    }
    if (const std::optional<std::uint32_t> target = mapped(fields, order[i])) {
      kept.push_back(*target);
    }
  }
  if (marker.has_value() && *marker >= order.size()) {
    moved_marker = kept.size();
  }
  order = std::move(kept);
  return moved_marker;
}

}  // namespace

void remap_table(PivotTable& table, const PivotCache& cache, const FieldRemap& remap) {
  const std::size_t old_count = table.fields().size();
  const std::size_t new_count = cache.fields().size();
  IndexMap fields(old_count);
  for (std::size_t i = 0; i < old_count; ++i) {
    const std::optional<std::uint32_t> target = mapped(remap.fields, i);
    if (target.has_value() && *target < new_count) {
      fields[i] = target;
    }
  }

  // Data fields and filters on a dropped field go while its name is still known.
  for (std::size_t i = 0; i < old_count; ++i) {
    if (!fields[i].has_value()) {
      drop_field_aggregates(table, i);
    }
  }
  // Name-keyed state takes the new spelling of the field it names.
  const auto renamed = [&](const std::string& name) -> std::optional<std::string> {
    if (remap.names.empty()) {
      return std::nullopt;
    }
    for (std::size_t i = 0; i < old_count; ++i) {
      if (fields[i].has_value() && *fields[i] < remap.names.size() && table.fields()[i].source_name == name &&
          remap.names[*fields[i]] != name) {
        return remap.names[*fields[i]];
      }
    }
    return std::nullopt;
  };
  for (PivotFilter& filter : table.mutable_active_filters()) {
    if (std::optional<std::string> name = renamed(filter.field_name)) {
      filter.field_name = std::move(*name);
    }
    if (filter.source.has_value() && mapped(fields, filter.source->field_index) != filter.source->field_index) {
      filter.source.reset();
    }
  }
  for (PivotField& field : table.mutable_fields()) {
    if (std::optional<std::string> name = renamed(field.sort.by_field)) {
      field.sort.by_field = std::move(*name);
    }
  }

  std::vector<PositionMap> positions(old_count);
  std::vector<PivotField> old_fields = std::move(table.mutable_fields());
  std::vector<std::optional<PivotField>> placed(new_count);
  for (std::size_t i = 0; i < old_count; ++i) {
    if (!fields[i].has_value()) {
      continue;
    }
    const std::uint32_t target = *fields[i];
    PivotField field = std::move(old_fields[i]);
    if (target < remap.names.size()) {
      field.source_name = remap.names[target];
    }
    if (i < remap.items.size() && !remap.items[i].empty()) {
      remap_items(field, remap.items[i], cache.fields()[target], positions[i]);
    }
    placed[target] = std::move(field);
  }
  std::vector<PivotField> new_fields;
  new_fields.reserve(new_count);
  for (std::size_t j = 0; j < new_count; ++j) {
    new_fields.push_back(placed[j].has_value() ? std::move(*placed[j]) : unplaced_pivot_field(cache.fields()[j]));
  }
  table.mutable_fields() = std::move(new_fields);

  for (PivotDataField& data_field : table.mutable_data_fields()) {
    data_field.field_index = mapped(fields, data_field.field_index).value_or(data_field.field_index);
    // Excel leaves the base of a plain data field as it was: it addresses nothing.
    if (!data_field.show_as_base_field.has_value() || data_field.show_as == ShowValuesAs::Normal) {
      continue;
    }
    const std::uint32_t base = *data_field.show_as_base_field;
    data_field.show_as_base_field = mapped(fields, base);
    if (!data_field.show_as_base_field.has_value()) {
      data_field.show_as_base_item.reset();
    } else if (data_field.show_as_base_item.has_value() && *data_field.show_as_base_item != kShowAsBasePrev &&
               *data_field.show_as_base_item != kShowAsBaseNext) {
      data_field.show_as_base_item = remap_item_position(remap, positions, base, *data_field.show_as_base_item);
    }
  }

  table.set_row_values_position(remap_order(table.mutable_row_field_order(), fields, table.row_values_position()));
  table.set_col_values_position(remap_order(table.mutable_col_field_order(), fields, table.col_values_position()));

  std::vector<PivotPageField> pages;
  for (PivotPageField page : table.page_fields()) {
    const std::optional<std::uint32_t> target = mapped(fields, page.field_index);
    if (!target.has_value()) {
      continue;
    }
    if (page.item_index.has_value()) {
      page.item_index = remap_item_position(remap, positions, page.field_index, *page.item_index);
    }
    page.field_index = *target;
    pages.push_back(std::move(page));
  }
  table.mutable_page_fields() = std::move(pages);
}

PivotField unplaced_pivot_field(const PivotCacheField& field) {
  PivotField out;
  out.source_name = field.name;
  out.axis = PivotAxis::None;
  if (field.formula.has_value()) {
    out.default_subtotal = false;
    out.passthrough_attrs = {{"dragToRow", "0"}, {"dragToCol", "0"}, {"dragToPage", "0"}};
  }
  return out;
}

std::optional<std::uint32_t> axis_position(const PivotTable& table, std::size_t field) {
  for (const std::vector<std::uint32_t>& order :
       {table.row_field_order(), table.col_field_order(), table.page_field_order()}) {
    if (const auto position = field_position_in_order(order, field)) {
      return static_cast<std::uint32_t>(*position);
    }
  }
  return std::nullopt;
}

void place_calculated_item_areas(PivotCache& cache, const PivotTable& table) {
  for (PivotCalculatedItem& entry : cache.mutable_calculated_items()) {
    if (!entry.area.field_position.has_value()) {
      continue;
    }
    if (const std::optional<std::uint32_t> position = axis_position(table, entry.cache_field_index)) {
      entry.area.field_position = position;
    }
  }
}

void clear_derived_state(PivotTable& table) {
  table.clear_last_result();
  table.clear_span_authored();
  table.mark_layout_dirty();
}

void drop_derived_location(PivotTable& table) {
  if (!table.has_authored_span() && table.projected_extent().has_value()) {
    table.set_location_attributes(std::nullopt, std::nullopt, std::nullopt, std::nullopt, std::nullopt);
  }
}

}  // namespace formulon::pivot
