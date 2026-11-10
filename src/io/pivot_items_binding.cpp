#include "io/pivot_items_binding.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <utility>
#include <vector>

#include "pivot/pivot_types.h"
#include "pivot/value_order.h"

namespace formulon::io::detail {

std::vector<pivot::PivotItem> BoundItems(const pivot::PivotField& field, const PivotCacheFieldWritePlan* fp,
                                         std::vector<bool>* bound_by_position) {
  if (bound_by_position != nullptr) {
    bound_by_position->clear();
  }
  if (fp == nullptr || !fp->derived) {
    if (bound_by_position != nullptr) {
      bound_by_position->assign(field.items.size(), false);
    }
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
    if (bound_by_position != nullptr) {
      bound_by_position->assign(bound.size(), false);
    }
    return bound;
  }
  for (std::size_t pos = 0; pos < field.items.size(); ++pos) {
    pivot::PivotItem item = field.items[pos];
    bool by_position = false;
    if (!item.has_cache_index) {
      std::size_t target = pos;
      by_position = true;
      if (!item.name.empty()) {
        for (std::size_t k = 0; k < shared; ++k) {
          if (pivot::display_string(fp->items[k]) == item.name) {
            target = k;
            by_position = false;
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
    if (bound_by_position != nullptr) {
      bound_by_position->push_back(by_position);
    }
  }
  for (std::size_t k = 0; k < shared; ++k) {
    if (!used[k]) {
      pivot::PivotItem item;
      item.has_cache_index = true;
      item.cache_index = static_cast<std::uint32_t>(k);
      bound.push_back(std::move(item));
      if (bound_by_position != nullptr) {
        bound_by_position->push_back(false);
      }
    }
  }
  return bound;
}

}  // namespace formulon::io::detail
