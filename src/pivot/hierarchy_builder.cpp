
#include "pivot/hierarchy_builder.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/text_format/display_text.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_types.h"
#include "pivot/record_access.h"
#include "pivot/value_order.h"
#include "utils/checked_index.h"
#include "value.h"

namespace formulon::pivot {
namespace {

/// Removes the leaves of `node`'s subtree whose position in `keep` is
/// false and any interior node left empty. `leaf_cursor` advances once per
/// visited leaf. Returns true iff `node` or any descendant survives.
bool prune_node(AxisHierarchyNode& node, const std::vector<bool>& keep, std::size_t& leaf_cursor) {
  if (node.children.empty()) {
    const bool survives = (leaf_cursor < keep.size()) ? keep[leaf_cursor] : true;
    ++leaf_cursor;
    return survives;
  }
  std::vector<AxisHierarchyNode> kept;
  kept.reserve(node.children.size());
  for (auto& child : node.children) {
    if (prune_node(child, keep, leaf_cursor)) {
      kept.push_back(std::move(child));
    }
  }
  node.children = std::move(kept);
  return !node.children.empty();
}

std::size_t count_axis_leaves(const AxisHierarchyNode& node) {
  if (node.children.empty()) {
    return 1;
  }
  std::size_t n = 0;
  for (const auto& child : node.children) {
    n += count_axis_leaves(child);
  }
  return n;
}

void collect_axis_leaf_groups(const AxisHierarchyNode& node, std::size_t depth, std::size_t target_depth,
                              std::size_t parent, std::size_t& next_parent, std::size_t& leaf_cursor,
                              std::vector<AxisLeafGroup>& out) {
  if (depth == target_depth || node.children.empty()) {
    const std::size_t count = count_axis_leaves(node);
    out.push_back({leaf_cursor, count, parent});
    leaf_cursor += count;
    return;
  }
  const std::size_t child_parent = depth + 1U == target_depth ? next_parent++ : parent;
  for (const auto& child : node.children) {
    collect_axis_leaf_groups(child, depth + 1U, target_depth, child_parent, next_parent, leaf_cursor, out);
  }
}

}  // namespace

HierNode* insert_path(const PivotCache& cache, const std::vector<HierLevel>& levels, const PivotCacheRecord& record,
                      std::size_t record_index, HierNode& root, std::string_view blank_item_label, bool date1904,
                      ExcelProfile profile) {
  HierNode* cursor = &root;
  for (const HierLevel& level : levels) {
    const Value raw = cell_value(cache, record, level.field_index);
    auto [it, inserted] = cursor->children.try_emplace(HierarchyKey{raw});
    if (inserted) {
      std::string label_override;
      const std::string* caption = nullptr;
      if (level.captions != nullptr) {
        if (const auto it_caption = level.captions->find(raw); it_caption != level.captions->end()) {
          caption = &it_caption->second;
        }
      }
      if (caption != nullptr) {
        label_override = *caption;
      } else if (level.group_items != nullptr && raw.is_number()) {
        const std::optional<std::size_t> item = index_from_double(raw.as_number(), level.group_items->size());
        if (item.has_value()) {
          label_override = display_string((*level.group_items)[*item]);
        }
      } else if (raw.is_blank()) {
        // A source cell with no value still occupies a group on the axis, so
        // the group needs a name: the label is what the grid draws, and it is
        // the only literal GETPIVOTDATA can address the group by, since that
        // lookup matches labels exactly. The placeholder text itself belongs
        // to the locale and arrives from the caller.
        label_override.assign(blank_item_label);
      } else if (!level.number_format.empty() && raw.is_number()) {
        label_override = text_format::format_value_for_display(raw, level.number_format, date1904, profile).text;
      }
      if (!label_override.empty()) {
        it->second.label_override = std::move(label_override);
      }
    }
    cursor = &it->second;
    cursor->first_record = std::min(cursor->first_record, record_index);
    // Only value-sorted levels consume record indices.
    if (level.value_sort_field.has_value() && level.value_sort_aggregation.has_value()) {
      cursor->record_indices.push_back(record_index);
    }
  }
  return cursor;
}

std::string node_label(const HierarchyKey& key, const HierNode& child) {
  if (!child.label_override.empty()) {
    return child.label_override;
  }
  return display_string(key.value);
}

void finalize_hierarchy(HierNode& tree, const std::vector<HierLevel>& levels, std::size_t depth,
                        std::vector<AxisHierarchyNode>& out, std::vector<HierNode*>& leaves,
                        std::deque<std::string>& text_storage) {
  if (tree.children.empty()) {
    return;
  }
  out.reserve(tree.children.size());
  const std::vector<OrderedHierarchyChild> entries = ordered_children(tree, levels, depth);
  const auto append = [&](const HierarchyKey& key, HierNode& child) {
    AxisHierarchyNode node;
    node.label = node_label(key, child);
    if (key.value.is_text()) {
      text_storage.emplace_back(key.value.as_text());
      node.identity = Value::text(text_storage.back());
    } else {
      node.identity = key.value;
    }
    if (depth < levels.size()) {
      const HierLevel& level = levels[depth];
      node.identity_is_group_item = level.group_items != nullptr;
      const bool captioned = level.captions != nullptr && level.captions->count(key.value) != 0U;
      if (key.value.is_number() && !captioned) {
        node.number_format = level.number_format;
      }
    }
    if (child.children.empty()) {
      child.leaf_index = leaves.size();
      leaves.push_back(&child);
    } else {
      finalize_hierarchy(child, levels, depth + 1U, node.children, leaves, text_storage);
    }
    out.push_back(std::move(node));
  };
  for (const OrderedHierarchyChild& entry : entries) {
    append(*entry.key, *entry.node);
  }
}

void prune_top_level(std::vector<AxisHierarchyNode>& roots, const std::vector<bool>& keep) {
  std::size_t cursor = 0;
  std::vector<AxisHierarchyNode> kept;
  kept.reserve(roots.size());
  for (auto& root : roots) {
    if (prune_node(root, keep, cursor)) {
      kept.push_back(std::move(root));
    }
  }
  roots = std::move(kept);
}

std::vector<AxisLeafGroup> axis_leaf_groups_at_depth(const std::vector<AxisHierarchyNode>& roots,
                                                     std::size_t target_depth) {
  std::vector<AxisLeafGroup> groups;
  std::size_t next_parent = target_depth == 0 ? 1U : 0U;
  std::size_t leaf_cursor = 0;
  for (const auto& root : roots) {
    collect_axis_leaf_groups(root, 0, target_depth, 0, next_parent, leaf_cursor, groups);
  }
  return groups;
}

}  // namespace formulon::pivot
