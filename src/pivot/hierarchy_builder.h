//
// Hierarchy construction for the pivot evaluator.
//
// The hierarchy is built as a nested `std::map<Value, HierNode, ValueLess>`
// so the ordering emerges naturally from `ValueLess`. After every record
// is inserted, `finalize_hierarchy` flattens the tree into the public
// `RowHierarchyNode` / `ColHierarchyNode` shape and remembers the leaf
// path that each surviving record lands on so the per-leaf aggregation
// pass can reuse the work without rewalking the tree.
//
// Date-grouping is handled inside `insert_path`: when a level carries a
// `PivotDateGroup`, the raw cache value is bucketed first (year /
// quarter / month / ...); the bucket's display label is stashed on the
// inserted child for the renderer to surface in place of the raw value.

#ifndef FORMULON_PIVOT_HIERARCHY_BUILDER_H_
#define FORMULON_PIVOT_HIERARCHY_BUILDER_H_

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <map>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/value_order.h"
#include "value.h"

namespace formulon::pivot {

struct HierNode {
  std::map<Value, HierNode, ValueLess> children;
  /// Cache-record indices below this node. They permit value-based sorting
  /// without reconstructing a hierarchy path after aggregation.
  std::vector<std::size_t> record_indices;
  /// Aggregate used when the owning field has SortSpec::by_field. Absent
  /// means normal label ordering remains in effect.
  std::optional<Value> value_sort_key;
  /// Index into the flat leaf array assigned during finalisation. Leaves only.
  std::size_t leaf_index = static_cast<std::size_t>(-1);
  /// When non-empty, used in place of `display_string(key)` for this
  /// node's label. Set by `insert_path` for date-grouped fields where
  /// the bucket label diverges from the raw value's textual form.
  std::string label_override;
};

struct HierLevel {
  std::uint32_t field_index;         ///< Index into `PivotTable::fields()`.
  const PivotDateGroup* date_group;  ///< Non-null when this level buckets dates.
  bool ascending;                    ///< False reverses this field's item order.
  std::optional<std::uint32_t> value_sort_field;
  std::optional<Aggregation> value_sort_aggregation;
};

struct OrderedHierarchyChild {
  const Value* key;
  HierNode* node;
};

/// Returns children in the same display order used by both hierarchy
/// finalisation and subtotal emission. Keeping this comparator in one place
/// is important: the raw map order is not necessarily the rendered order
/// when a field is descending or sorted by a value field.
inline std::vector<OrderedHierarchyChild> ordered_children(HierNode& tree, const std::vector<HierLevel>& levels,
                                                           std::size_t depth) {
  std::vector<OrderedHierarchyChild> entries;
  entries.reserve(tree.children.size());
  for (auto& [key, child] : tree.children) {
    entries.push_back({&key, &child});
  }
  const bool ascending = depth >= levels.size() || levels[depth].ascending;
  const bool sort_by_value = depth < levels.size() && levels[depth].value_sort_field.has_value();
  std::sort(entries.begin(), entries.end(), [&](const OrderedHierarchyChild& lhs, const OrderedHierarchyChild& rhs) {
    if (sort_by_value && lhs.node->value_sort_key.has_value() && rhs.node->value_sort_key.has_value()) {
      const ValueLess less;
      if (less(*lhs.node->value_sort_key, *rhs.node->value_sort_key)) {
        return ascending;
      }
      if (less(*rhs.node->value_sort_key, *lhs.node->value_sort_key)) {
        return !ascending;
      }
    }
    const ValueLess less;
    return ascending ? less(*lhs.key, *rhs.key) : less(*rhs.key, *lhs.key);
  });
  return entries;
}

/// Inserts `record` into `tree`, walking `levels`. Returns the leaf
/// `HierNode*`. The caller assigns leaf indices in a second pass. When
/// a level carries a `date_group`, the cache value is bucketed first;
/// the label is stashed on the inserted child for the renderer to
/// surface. A blank cache value takes `blank_item_label` (the locale's
/// placeholder) the same way, so no axis node is left unnamed.
HierNode* insert_path(const PivotCache& cache, const std::vector<HierLevel>& levels, const PivotCacheRecord& record,
                      std::size_t record_index, HierNode& root, std::string_view blank_item_label);

/// Returns the display label for `(key, child)`: the override if set,
/// otherwise the standard `display_string(key)`. Used by all hierarchy
/// flatten / subtotal-walk sites so date-grouped buckets surface their
/// formatted label rather than the synthetic numeric sort key.
std::string node_label(const Value& key, const HierNode& child);

/// Recursively flattens a hierarchy into `Node` form (templated so we
/// can produce both `RowHierarchyNode` and `ColHierarchyNode` from one
/// implementation). On the way, assigns each leaf a dense index and
/// pushes the corresponding `HierNode*` into `leaves` so a second pass
/// can attach record indices.
template <class Node>
void finalize_hierarchy(HierNode& tree, const std::vector<HierLevel>& levels, std::size_t depth, std::vector<Node>& out,
                        std::vector<HierNode*>& leaves) {
  if (tree.children.empty()) {
    return;
  }
  out.reserve(tree.children.size());
  const std::vector<OrderedHierarchyChild> entries = ordered_children(tree, levels, depth);
  const auto append = [&](const Value& key, HierNode& child) {
    Node node;
    node.label = node_label(key, child);
    if (child.children.empty()) {
      child.leaf_index = leaves.size();
      leaves.push_back(&child);
    } else {
      finalize_hierarchy<Node>(child, levels, depth + 1U, node.children, leaves);
    }
    out.push_back(std::move(node));
  };
  for (const OrderedHierarchyChild& entry : entries) {
    append(*entry.key, *entry.node);
  }
}

/// Removes leaves at positions where `keep[i]` is false, then prunes
/// any interior node whose subtree becomes empty. `leaf_cursor` is
/// advanced once per visited leaf so the caller's flat `keep` vector
/// lines up with the DFS pre-order leaf enumeration produced by
/// `finalize_hierarchy`. Returns true iff `node` (or any of its
/// descendants) survives the prune.
template <class Node>
bool prune_node(Node& node, const std::vector<bool>& keep, std::size_t& leaf_cursor) {
  if (node.children.empty()) {
    const bool survives = (leaf_cursor < keep.size()) ? keep[leaf_cursor] : true;
    ++leaf_cursor;
    return survives;
  }
  std::vector<Node> kept;
  kept.reserve(node.children.size());
  for (auto& child : node.children) {
    if (prune_node(child, keep, leaf_cursor)) {
      kept.push_back(std::move(child));
    }
  }
  node.children = std::move(kept);
  return !node.children.empty();
}

/// Top-level driver for `prune_node`: walks each root in document order
/// while threading a single leaf cursor through the whole tree so the
/// caller's `keep` vector, indexed by DFS pre-order leaf position, lines
/// up correctly across roots. Roots whose subtrees become empty are
/// discarded.
template <class Node>
void prune_top_level(std::vector<Node>& roots, const std::vector<bool>& keep) {
  std::size_t cursor = 0;
  std::vector<Node> kept;
  kept.reserve(roots.size());
  for (auto& root : roots) {
    if (prune_node(root, keep, cursor)) {
      kept.push_back(std::move(root));
    }
  }
  roots = std::move(kept);
}

/// Walks `tree` in display order. At every non-leaf level whose
/// `PivotField` declares `subtotal_top` or any `subtotal_fns`, calls
/// `emit_subtotal(labels, depth, collected_start, stack_leaves)` after
/// all descendant leaves have been pushed onto `stack_leaves`.
///
/// `labels` is the label path from root to the subtotal owner; `depth`
/// is the field-order depth of the owner; `collected_start` is the
/// index into `stack_leaves` where this owner's descendants begin
/// (so `[collected_start, stack_leaves.size())` is its leaf set).
template <class EmitSubtotal>
void walk_subtotal_tree(HierNode& tree, const std::vector<HierLevel>& levels, const PivotTable& table,
                        std::vector<std::size_t>& stack_leaves, EmitSubtotal&& emit_subtotal) {
  auto field_at_depth = [&](std::size_t depth) -> const PivotField* {
    if (depth >= levels.size()) {
      return nullptr;
    }
    const std::uint32_t fi = levels[depth].field_index;
    if (fi >= table.fields().size()) {
      return nullptr;
    }
    return &table.fields()[fi];
  };

  struct Frame {
    HierNode* node;
    std::vector<OrderedHierarchyChild> children;
    std::size_t child_cursor;
    std::size_t depth;
    std::size_t collected_start;
    std::vector<std::string> labels;
  };

  std::vector<Frame> stack;
  stack.push_back({&tree, ordered_children(tree, levels, 0), 0, 0, 0, {}});

  while (!stack.empty()) {
    Frame& top = stack.back();
    if (top.child_cursor == top.children.size()) {
      if (top.depth > 0 && !top.node->children.empty()) {
        const PivotField* field = field_at_depth(top.depth - 1);
        // `subtotal_top` is only the position (top vs bottom) of the
        // subtotal row; whether a subtotal is emitted at all is governed
        // by `default_subtotal` (OOXML default true) plus any explicit
        // custom subtotal functions.
        const bool wants_subtotal = field != nullptr && (field->default_subtotal || !field->subtotal_fns.empty());
        if (wants_subtotal) {
          emit_subtotal(top.labels, top.depth - 1, top.collected_start, stack_leaves);
        }
      }
      stack.pop_back();
      continue;
    }
    const OrderedHierarchyChild& entry = top.children[top.child_cursor++];
    HierNode* child = entry.node;
    const std::string label = node_label(*entry.key, *child);
    if (child->children.empty()) {
      stack_leaves.push_back(child->leaf_index);
    } else {
      std::vector<std::string> labels = top.labels;
      labels.push_back(label);
      stack.push_back({child, ordered_children(*child, levels, top.depth + 1U), 0, top.depth + 1U, stack_leaves.size(),
                       std::move(labels)});
    }
  }
}

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_HIERARCHY_BUILDER_H_
