//
// Hierarchy construction for the pivot evaluator.
//
// The hierarchy is built as a nested `std::map<Value, HierNode, ValueLess>`
// so the ordering emerges naturally from `ValueLess`. After every record
// is inserted, `finalize_hierarchy` flattens the tree into the public
// `AxisHierarchyNode` shape and remembers the leaf
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
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/value_order.h"
#include "utils/index_sort.h"
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
  /// Non-null when this field's `SortSpec::manual` is set: maps each
  /// `<items>`-enumerated cache value to its document position, so
  /// `ordered_children` can sort siblings by authored order instead of
  /// by label. A value the map does not contain (the cache carries data
  /// `<items>` never enumerated) sorts after every mapped value, in the
  /// field's natural ascending order.
  const std::map<Value, std::size_t, ValueLess>* manual_order;
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
  const std::map<Value, std::size_t, ValueLess>* manual_order =
      depth < levels.size() ? levels[depth].manual_order : nullptr;
  sort_by_index(entries, [&](const OrderedHierarchyChild& lhs, const OrderedHierarchyChild& rhs) {
    if (manual_order != nullptr) {
      const auto lit = manual_order->find(*lhs.key);
      const auto rit = manual_order->find(*rhs.key);
      const bool l_mapped = lit != manual_order->end();
      const bool r_mapped = rit != manual_order->end();
      if (l_mapped && r_mapped) {
        if (lit->second != rit->second) {
          return lit->second < rit->second;
        }
        // Equal authored position (a malformed file): fall through to
        // label order below.
      } else if (l_mapped != r_mapped) {
        return l_mapped;  // Authored items sort before ones the file never enumerated.
      }
    }
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
/// a level carries a `date_group`, the cache value is bucketed first,
/// against the workbook's `date1904` epoch flag; the label is stashed on
/// the inserted child for the renderer to surface. A blank cache value
/// takes `blank_item_label` (the locale's placeholder) the same way, so
/// no axis node is left unnamed.
HierNode* insert_path(const PivotCache& cache, const std::vector<HierLevel>& levels, const PivotCacheRecord& record,
                      std::size_t record_index, HierNode& root, std::string_view blank_item_label,
                      bool date1904 = false);

/// Returns the display label for `(key, child)`: the override if set,
/// otherwise the standard `display_string(key)`. Used by all hierarchy
/// flatten / subtotal-walk sites so date-grouped buckets surface their
/// formatted label rather than the synthetic numeric sort key.
std::string node_label(const Value& key, const HierNode& child);

/// Recursively flattens a hierarchy into `AxisHierarchyNode` form. On the
/// way, assigns each leaf a dense index and pushes the corresponding
/// `HierNode*` into `leaves` so a second pass can attach record indices.
void finalize_hierarchy(HierNode& tree, const std::vector<HierLevel>& levels, std::size_t depth,
                        std::vector<AxisHierarchyNode>& out, std::vector<HierNode*>& leaves);

/// Removes the leaves at DFS pre-order positions where `keep[i]` is false,
/// then drops any node whose subtree becomes empty, including roots. One
/// leaf cursor is threaded through every root so `keep` lines up with the
/// leaf enumeration `finalize_hierarchy` produced.
void prune_top_level(std::vector<AxisHierarchyNode>& roots, const std::vector<bool>& keep);

/// One axis node at a chosen depth, as the contiguous run of DFS pre-order
/// leaves beneath it. `parent` numbers the node's depth-minus-one ancestor;
/// every group of a depth-0 query shares parent 0.
struct AxisLeafGroup {
  std::size_t first_leaf = 0;
  std::size_t leaf_count = 0;
  std::size_t parent = 0;
};

/// Partitions the (possibly already pruned) axis tree `roots` into its
/// nodes at `target_depth`, in document order. Grouping is structural: it
/// does not depend on whether any field displays subtotals.
std::vector<AxisLeafGroup> axis_leaf_groups_at_depth(const std::vector<AxisHierarchyNode>& roots,
                                                     std::size_t target_depth);

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
