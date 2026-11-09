
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

/// The label a node keyed `raw` on `level` draws instead of its cache value's
/// text; empty when it draws that text.
std::string label_override_for(const HierLevel& level, const Value& raw, std::string_view blank_item_label,
                               bool date1904, ExcelProfile profile) {
  if (level.captions != nullptr) {
    if (const auto it_caption = level.captions->find(raw); it_caption != level.captions->end()) {
      return it_caption->second;
    }
  }
  if (level.group_items != nullptr && raw.is_number()) {
    const std::optional<std::size_t> item = index_from_double(raw.as_number(), level.group_items->size());
    return item.has_value() ? display_string((*level.group_items)[*item]) : std::string();
  }
  if (raw.is_blank()) {
    // A source cell with no value still occupies a group on the axis, so
    // the group needs a name: the label is what the grid draws, and it is
    // the only literal GETPIVOTDATA can address the group by, since that
    // lookup matches labels exactly. The placeholder text itself belongs
    // to the locale and arrives from the caller.
    return std::string(blank_item_label);
  }
  if (!level.number_format.empty() && raw.is_number()) {
    return text_format::format_value_for_display(raw, level.number_format, date1904, profile).text;
  }
  return {};
}

/// Copies the subtree shape of `src` (keys, labels, calculated flags) into `dst`.
void merge_shape(HierNode& dst, const HierNode& src) {
  for (const auto& [key, child] : src.children) {
    auto [it, inserted] = dst.children.try_emplace(key);
    if (inserted) {
      it->second.label_override = child.label_override;
      it->second.calculated = child.calculated;
    }
    merge_shape(it->second, child);
  }
}

void add_recordless_at(HierNode& node, const std::vector<HierLevel>& levels,
                       const std::vector<std::vector<RecordlessItem>>& items, std::size_t depth,
                       std::string_view blank_item_label, bool date1904, ExcelProfile profile) {
  if (depth >= levels.size()) {
    return;
  }
  if (depth < items.size()) {
    std::vector<const HierNode*> real_children;
    for (const auto& entry : node.children) {
      if (!entry.second.calculated) {
        real_children.push_back(&entry.second);
      }
    }
    for (const RecordlessItem& item : items[depth]) {
      auto [it, inserted] = node.children.try_emplace(HierarchyKey{item.key});
      if (!inserted) {
        continue;
      }
      it->second.label_override = label_override_for(levels[depth], item.key, blank_item_label, date1904, profile);
      it->second.calculated = item.calculated;
      if (item.calculated && depth + 1U < levels.size()) {
        for (const HierNode* sibling : real_children) {
          merge_shape(it->second, *sibling);
        }
      }
    }
  }
  for (auto& entry : node.children) {
    add_recordless_at(entry.second, levels, items, depth + 1U, blank_item_label, date1904, profile);
  }
}

void collect_paths_at(HierNode& node, const std::vector<HierLevel>& levels, std::size_t depth,
                      std::vector<LeafPathStep>& path, AxisLeafPaths& out) {
  const std::vector<OrderedHierarchyChild> entries = ordered_children(node, levels, depth);
  const std::size_t siblings = out.sibling_keys.size();
  out.sibling_keys.emplace_back();
  for (const OrderedHierarchyChild& entry : entries) {
    out.sibling_keys[siblings].push_back(entry.key->value);
  }
  for (const OrderedHierarchyChild& entry : entries) {
    path.push_back(LeafPathStep{depth < levels.size() ? levels[depth].field_index : 0U, entry.key->value, siblings});
    if (entry.node->children.empty()) {
      if (entry.node->leaf_index >= out.paths.size()) {
        out.paths.resize(entry.node->leaf_index + 1U);
      }
      out.paths[entry.node->leaf_index] = path;
    } else {
      collect_paths_at(*entry.node, levels, depth + 1U, path, out);
    }
    path.pop_back();
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
      it->second.label_override = label_override_for(level, raw, blank_item_label, date1904, profile);
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

void add_recordless_nodes(HierNode& tree, const std::vector<HierLevel>& levels,
                          const std::vector<std::vector<RecordlessItem>>& items, std::string_view blank_item_label,
                          bool date1904, ExcelProfile profile) {
  add_recordless_at(tree, levels, items, 0U, blank_item_label, date1904, profile);
}

AxisLeafPaths collect_leaf_paths(HierNode& tree, const std::vector<HierLevel>& levels) {
  AxisLeafPaths out;
  if (levels.empty()) {
    out.paths.emplace_back();
    return out;
  }
  std::vector<LeafPathStep> path;
  collect_paths_at(tree, levels, 0U, path, out);
  return out;
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
      node.identity = intern_text(text_storage, key.value.as_text());
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
