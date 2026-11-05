//
// Axis planning for the pivot layout: projects an evaluated axis hierarchy into
// the row or column entries the emitters lay out.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "pivot/field_lookup.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_layout_internal.h"
#include "pivot/pivot_locale.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/value_order.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon::pivot::detail {
namespace {

void collect_axis_leaves_impl(const AxisHierarchyNode& node, std::vector<std::string>& path,
                              std::vector<AxisLeaf>& leaves) {
  path.push_back(node.label);
  if (node.children.empty()) {
    leaves.push_back({path});
  } else {
    for (const AxisHierarchyNode& child : node.children) {
      collect_axis_leaves_impl(child, path, leaves);
    }
  }
  path.pop_back();
}

}  // namespace

std::vector<AxisLeaf> collect_axis_leaves(const std::vector<AxisHierarchyNode>& roots, std::size_t depth) {
  std::vector<AxisLeaf> leaves;
  if (depth == 0) {
    leaves.push_back(AxisLeaf{});
    return leaves;
  }
  std::vector<std::string> path;
  for (const AxisHierarchyNode& root : roots) {
    collect_axis_leaves_impl(root, path, leaves);
  }
  return leaves;
}

namespace {

bool labels_equal(const std::vector<std::string>& a, const std::vector<std::string>& b) {
  return a.size() == b.size() && std::equal(a.begin(), a.end(), b.begin());
}

}  // namespace

std::vector<std::size_t> subtotal_counts_for_axis(const PivotTable& table,
                                                  const std::vector<std::uint32_t>& field_order) {
  std::vector<std::size_t> counts(field_order.size(), 0);
  for (std::size_t depth = 0; depth < field_order.size(); ++depth) {
    const std::uint32_t field_index = field_order[depth];
    if (field_index >= table.fields().size()) {
      continue;
    }
    const PivotField& field = table.fields()[field_index];
    const std::size_t custom = ordered_subtotal_fns(field).size();
    counts[depth] = custom > 0 ? custom : (field.default_subtotal ? 1U : 0U);
  }
  return counts;
}

/// Per depth, the field the axis places there.
std::vector<const PivotField*> fields_for_axis(const PivotTable& table, const std::vector<std::uint32_t>& field_order) {
  std::vector<const PivotField*> fields(field_order.size(), nullptr);
  for (std::size_t depth = 0; depth < field_order.size(); ++depth) {
    if (field_order[depth] < table.fields().size()) {
      fields[depth] = &table.fields()[field_order[depth]];
    }
  }
  return fields;
}

/// Per depth, the field's custom subtotal functions in Excel's fixed order.
std::vector<std::vector<SubtotalFn>> subtotal_fns_for_axis(const PivotTable& table,
                                                           const std::vector<std::uint32_t>& field_order) {
  std::vector<std::vector<SubtotalFn>> fns(field_order.size());
  for (std::size_t depth = 0; depth < field_order.size(); ++depth) {
    if (field_order[depth] < table.fields().size()) {
      fns[depth] = ordered_subtotal_fns(table.fields()[field_order[depth]]);
    }
  }
  return fns;
}

namespace {

/// The two fields of a `RowSubtotal` / `ColSubtotal` the projection reads,
/// so one non-template body serves both axes.
struct SubtotalView {
  std::uint32_t depth = 0;
  const std::vector<std::string>* labels = nullptr;
};

/// Checks that the evaluator's subtotals follow the hierarchy in its sorted
/// post-order, the order the projection consumes them in.
Expected<void, Error> check_axis_impl(const AxisHierarchyNode& node, std::vector<std::string>& path,
                                      const std::vector<SubtotalView>& subtotals, std::size_t& subtotal_cursor,
                                      const std::vector<std::size_t>& subtotal_counts, const std::string& axis) {
  path.push_back(node.label);
  if (node.children.empty()) {
    path.pop_back();
    return Expected<void, Error>::Ok();
  }
  const std::size_t depth = path.size() - 1;
  for (const AxisHierarchyNode& child : node.children) {
    auto child_or = check_axis_impl(child, path, subtotals, subtotal_cursor, subtotal_counts, axis);
    if (!child_or) {
      path.pop_back();
      return std::move(child_or.error());
    }
  }
  // The field count bounds this node's block, which also keeps distinct
  // nodes apart when their display labels coincide after formatting.
  const std::size_t expected_count = depth < subtotal_counts.size() ? subtotal_counts[depth] : 0;
  if (subtotal_cursor + expected_count > subtotals.size()) {
    path.pop_back();
    return make_error(FormulonErrorCode::kEvalPivotInvalid,
                      "pivot layout: " + axis + " subtotal metadata ended before hierarchy projection",
                      "depth=" + std::to_string(depth) + " cursor=" + std::to_string(subtotal_cursor) +
                          " expected=" + std::to_string(expected_count) + " total=" + std::to_string(subtotals.size()));
  }
  for (std::size_t offset = 0; offset < expected_count; ++offset) {
    const SubtotalView& subtotal = subtotals[subtotal_cursor + offset];
    if (subtotal.depth != depth || !labels_equal(*subtotal.labels, path)) {
      path.pop_back();
      return make_error(FormulonErrorCode::kEvalPivotInvalid,
                        "pivot layout: " + axis + " subtotal metadata does not match sorted hierarchy projection",
                        "index=" + std::to_string(subtotal_cursor + offset) +
                            " depth=" + std::to_string(subtotal.depth) + " expected_depth=" + std::to_string(depth));
    }
  }
  subtotal_cursor += expected_count;
  path.pop_back();
  return Expected<void, Error>::Ok();
}

// The `<item>` a node renders: by group item number, else by the cache label
// `PivotItem::name` resolves to.
const PivotItem* item_for_node(const PivotField& field, const AxisHierarchyNode& node) {
  if (!node.identity.has_value()) {
    return nullptr;
  }
  const Value& identity = *node.identity;
  if (node.identity_is_group_item) {
    for (const PivotItem& item : field.items) {
      if (identity.is_number() && item.has_cache_index &&
          static_cast<double>(item.cache_index) == identity.as_number()) {
        return &item;
      }
    }
    return nullptr;
  }
  const std::string label = display_string(identity);
  for (const PivotItem& item : field.items) {
    if (item.name == label) {
      return &item;
    }
  }
  return nullptr;
}

class AxisWalk {
 public:
  AxisWalk(const PivotTable& table, const AxisShape& shape) : table_(table), shape_(shape) {}

  std::vector<AxisEntry> run(const std::vector<AxisHierarchyNode>& roots) {
    level(roots, 0, std::nullopt, 0);
    return std::move(entries_);
  }

 private:
  void push(EntryKind kind, std::optional<std::size_t> data_field, std::size_t index = 0) {
    AxisEntry entry;
    entry.labels = path_;
    entry.nodes = node_path_;
    entry.kind = kind;
    entry.data_field = data_field;
    if (kind == EntryKind::Leaf) {
      entry.leaf_index = index;
    } else if (kind == EntryKind::Subtotal) {
      entry.subtotal_index = index;
    }
    entries_.push_back(std::move(entry));
  }

  // Projects the nodes at real depth `depth`, first expanding the Values
  // level when it sits here. `leaf` is the evaluated leaf a path that has
  // run out of real levels stands for.
  void level(const std::vector<AxisHierarchyNode>& nodes, std::size_t depth, std::optional<std::size_t> data_field,
             std::size_t leaf) {
    if (!data_field && shape_.values_at == depth) {
      // Each data field repeats the same subtree, so each starts from the
      // same leaf and subtotal positions.
      const std::size_t leaf_start = leaf_cursor_;
      const std::size_t subtotal_start = subtotal_cursor_;
      for (std::size_t df = 0; df < table_.data_fields().size(); ++df) {
        leaf_cursor_ = leaf_start;
        subtotal_cursor_ = subtotal_start;
        path_.push_back(table_.data_fields()[df].name);
        node_path_.push_back(nullptr);
        if (depth == shape_.depth) {
          push(EntryKind::Leaf, df, leaf);
        } else {
          if (shape_.group_rows) {
            push(EntryKind::Group, df);
          }
          level(nodes, depth, df, leaf);
        }
        path_.pop_back();
        node_path_.pop_back();
      }
      return;
    }
    if (depth == shape_.depth) {
      push(EntryKind::Leaf, data_field, leaf);
      return;
    }
    for (const AxisHierarchyNode& node : nodes) {
      this->node(node, depth, data_field);
    }
  }

  // Whether `node` hides its children: its item's `show_detail`, else the
  // field's `default` marker. The innermost field has nothing to hide.
  bool is_collapsed(const AxisHierarchyNode& node, std::size_t depth) const {
    if (node.children.empty() || depth + 1 >= shape_.depth || depth >= shape_.fields.size() ||
        shape_.fields[depth] == nullptr) {
      return false;
    }
    const PivotField& field = *shape_.fields[depth];
    if (const PivotItem* item = item_for_node(field, node); item != nullptr) {
      return !item->show_detail;
    }
    for (const PivotMarkerItem& marker : field.marker_items) {
      if (marker.type == "default") {
        return !marker.show_detail;
      }
    }
    return false;
  }

  // Advances the cursors past `node`'s descendants without projecting them.
  void skip_descendants(const AxisHierarchyNode& node, std::size_t depth) {
    for (const AxisHierarchyNode& child : node.children) {
      if (child.children.empty()) {
        ++leaf_cursor_;
        continue;
      }
      skip_descendants(child, depth + 1);
      subtotal_cursor_ += depth + 1 < shape_.subtotal_counts.size() ? shape_.subtotal_counts[depth + 1] : 0;
    }
  }

  void node(const AxisHierarchyNode& node, std::size_t depth, std::optional<std::size_t> data_field) {
    path_.push_back(node.label);
    node_path_.push_back(&node);
    // A group holding the Values level has no single value to show on its
    // own row, so Excel opens it with a label-only row and moves its
    // subtotals below it, one per data field.
    const bool values_below = !data_field && shape_.values_at.has_value() && *shape_.values_at > depth;
    const std::size_t count =
        !node.children.empty() && depth < shape_.subtotal_counts.size() ? shape_.subtotal_counts[depth] : 0;
    // A collapsed item shows its subtotal in place of its children, so one
    // without a subtotal stays expanded.
    const bool collapsed = count > 0 && is_collapsed(node, depth);
    // Excel ignores `subtotal_top` once a field carries several functions:
    // the header row stays blank and every function gets a row below.
    const bool first = count == 1 && !values_below &&
                       (collapsed || (depth < shape_.subtotal_first.size() && shape_.subtotal_first[depth]));
    const std::vector<SubtotalFn>* fns = depth < shape_.subtotal_fns.size() ? &shape_.subtotal_fns[depth] : nullptr;
    // Any other group shows its label on a row of its own: a subtotal at the
    // bottom, or none, cannot stand in for it.
    if (shape_.group_rows && (values_below || (!node.children.empty() && !first))) {
      push(EntryKind::Group, std::nullopt);
    }
    if (node.children.empty()) {
      level({}, depth + 1, data_field, leaf_cursor_++);
      path_.pop_back();
      node_path_.pop_back();
      return;
    }
    if (collapsed && values_below) {
      skip_descendants(node, depth);
      push_collapsed_values(depth, subtotal_cursor_);
      subtotal_cursor_ += count;
      close_node(depth);
      return;
    }
    const std::size_t group_begin = entries_.size();
    if (collapsed) {
      skip_descendants(node, depth);
    } else {
      level(node.children, depth + 1, data_field, 0);
    }
    const std::size_t subtotal_begin = entries_.size();
    const std::size_t subtotal_index = subtotal_cursor_;
    subtotal_cursor_ += count;
    for (std::size_t offset = 0; offset < count; ++offset) {
      if (values_below) {
        for (std::size_t df = 0; df < table_.data_fields().size(); ++df) {
          push(EntryKind::Subtotal, df, subtotal_index + offset);
          entries_.back().subtotal_label = node.label + " " + table_.data_fields()[df].name;
        }
      } else {
        push(EntryKind::Subtotal, data_field, subtotal_index + offset);
        entries_.back().heads_group = first;
        if (!first && fns != nullptr && offset < fns->size()) {
          entries_.back().custom_fn = (*fns)[offset];
        }
      }
    }
    if (first) {
      std::rotate(entries_.begin() + static_cast<std::ptrdiff_t>(group_begin),
                  entries_.begin() + static_cast<std::ptrdiff_t>(subtotal_begin), entries_.end());
    }
    close_node(depth);
  }

  // A collapsed item with the Values level below it stands in for its
  // subtree: the skipped levels stay empty and each data field still gets
  // its own entry, carrying the item's subtotal.
  void push_collapsed_values(std::size_t depth, std::size_t subtotal_index) {
    const std::size_t skipped = *shape_.values_at - depth - 1;
    for (std::size_t i = 0; i < skipped; ++i) {
      path_.emplace_back();
      node_path_.push_back(nullptr);
    }
    for (std::size_t df = 0; df < table_.data_fields().size(); ++df) {
      path_.push_back(table_.data_fields()[df].name);
      node_path_.push_back(nullptr);
      push(EntryKind::Subtotal, df, subtotal_index);
      entries_.back().heads_group = true;
      path_.pop_back();
      node_path_.pop_back();
    }
    path_.resize(path_.size() - skipped);
    node_path_.resize(node_path_.size() - skipped);
  }

  // Leaves the node at `depth`, closing its block with a blank row when its field asks for one.
  void close_node(std::size_t depth) {
    path_.pop_back();
    node_path_.pop_back();
    if (depth < shape_.blank_after.size() && shape_.blank_after[depth]) {
      push(EntryKind::Blank, std::nullopt);
    }
  }

  const PivotTable& table_;
  const AxisShape& shape_;
  std::size_t leaf_cursor_ = 0;
  std::size_t subtotal_cursor_ = 0;
  std::vector<std::string> path_;
  std::vector<const AxisHierarchyNode*> node_path_;
  std::vector<AxisEntry> entries_;
};

}  // namespace

/// Projects one axis hierarchy into its entries. `axis` names the axis in
/// diagnostics.
template <typename Subtotal>
Expected<std::vector<AxisEntry>, Error> collect_axis_entries(const PivotTable& table,
                                                             const std::vector<AxisHierarchyNode>& roots,
                                                             const std::vector<Subtotal>& subtotals,
                                                             const AxisShape& shape, const std::string& axis) {
  if (shape.depth == 0 && !subtotals.empty()) {
    return make_error(FormulonErrorCode::kEvalPivotInvalid,
                      "pivot layout: " + axis + " subtotals require a non-empty " + axis + " hierarchy",
                      "subtotals=" + std::to_string(subtotals.size()));
  }
  std::vector<SubtotalView> views;
  views.reserve(subtotals.size());
  for (const Subtotal& subtotal : subtotals) {
    views.push_back(SubtotalView{subtotal.depth, &subtotal.labels});
  }
  // Without subtotals the evaluator emitted none, whatever the fields ask for.
  const std::vector<std::size_t> counts =
      subtotals.empty() ? std::vector<std::size_t>(shape.depth, 0) : shape.subtotal_counts;
  std::vector<std::string> path;
  std::size_t subtotal_cursor = 0;
  if (shape.depth > 0) {
    for (const AxisHierarchyNode& root : roots) {
      auto root_or = check_axis_impl(root, path, views, subtotal_cursor, counts, axis);
      if (!root_or) {
        return std::move(root_or.error());
      }
    }
  }
  if (subtotal_cursor != views.size()) {
    return make_error(FormulonErrorCode::kEvalPivotInvalid,
                      "pivot layout: " + axis + " subtotal metadata was not consumed by sorted hierarchy projection",
                      "consumed=" + std::to_string(subtotal_cursor) + " total=" + std::to_string(views.size()));
  }
  AxisShape effective = shape;
  effective.subtotal_counts = counts;
  static const std::vector<AxisHierarchyNode> kNoNodes;
  return AxisWalk(table, effective).run(shape.depth > 0 ? roots : kNoNodes);
}

template Expected<std::vector<AxisEntry>, Error> collect_axis_entries<RowSubtotal>(
    const PivotTable&, const std::vector<AxisHierarchyNode>&, const std::vector<RowSubtotal>&, const AxisShape&,
    const std::string&);
template Expected<std::vector<AxisEntry>, Error> collect_axis_entries<ColSubtotal>(
    const PivotTable&, const std::vector<AxisHierarchyNode>&, const std::vector<ColSubtotal>&, const AxisShape&,
    const std::string&);

}  // namespace formulon::pivot::detail
