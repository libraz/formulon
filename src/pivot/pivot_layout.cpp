//
// Implementation of the PivotResult -> grid-cell projection.

#include "pivot/pivot_layout.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <deque>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "pivot/field_lookup.h"
#include "pivot/pivot_locale.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/value_order.h"
#include "utils/checked_mul.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon::pivot {
namespace {

struct AxisLeaf {
  std::vector<std::string> labels;
};

/// Columns a page-field header row occupies: the field name and the item it
/// is showing. Excel leaves the rest of the row empty however wide the
/// report below it is.
constexpr std::size_t kPageHeaderCols = 2;

const std::string& data_field_name(const PivotTable& table, std::size_t index) {
  static const std::string kNone;
  return index < table.data_fields().size() ? table.data_fields()[index].name : kNone;
}

const std::string& data_field_format(const PivotTable& table, std::size_t index) {
  static const std::string kNone;
  return index < table.data_fields().size() ? table.data_fields()[index].number_format : kNone;
}

// A show-as data field states its total cells itself; a blank one stays blank instead of
// being re-summed from the raw cells.
bool total_is_derived(const PivotTable& table, std::size_t index) {
  return index < table.data_fields().size() && table.data_fields()[index].show_as != ShowValuesAs::Normal;
}

Value text_value(PivotCells& cells, std::string text) {
  cells.text_storage.push_back(std::move(text));
  return Value::text(cells.text_storage.back());
}

Value reify_value(PivotCells& cells, const Value& value) {
  if (!value.is_text()) {
    return value;
  }
  return text_value(cells, std::string(value.as_text()));
}

void append_cell(PivotCells& cells, std::uint32_t row, std::uint32_t col, Value value, PivotCellKind kind,
                 PivotCellRole role, std::uint32_t depth, std::string field_name = {}, std::string number_format = {}) {
  PivotCell cell;
  cell.row = row;
  cell.col = col;
  cell.value = value;
  cell.kind = kind;
  cell.role = role;
  cell.depth = depth;
  cell.field_name = std::move(field_name);
  cell.number_format = std::move(number_format);
  cells.cells.push_back(std::move(cell));
}

void append_blank(PivotCells& cells, std::uint32_t row, std::uint32_t col, std::uint32_t depth = 0) {
  append_cell(cells, row, col, Value::blank(), PivotCellKind::Blank, PivotCellRole::None, depth);
}

// A depth-0 cell of data field `df`, carrying its name and number format.
void append_df_cell(PivotCells& cells, std::uint32_t row, std::uint32_t col, Value value, PivotCellKind kind,
                    PivotCellRole role, const PivotTable& table, std::size_t df) {
  append_cell(cells, row, col, value, kind, role, 0, data_field_name(table, df), data_field_format(table, df));
}

PivotCellRole value_role(PivotCellKind kind) {
  switch (kind) {
    case PivotCellKind::RowSubtotal:
    case PivotCellKind::ColSubtotal:
      return PivotCellRole::SubtotalValue;
    case PivotCellKind::GrandTotal:
      return PivotCellRole::GrandTotalValue;
    default:
      return PivotCellRole::DataValue;
  }
}

// The cell Excel writes for `value`: an error under `showError` shows the error
// caption, and an empty cell shows a non-empty missing caption, where a lone
// "-" is written as the number 0.
Value shown_value(const PivotTable& table, PivotCells& cells, const Value& value) {
  if (value.is_error() && table.show_error()) {
    const std::string caption = table.error_caption().value_or(std::string());
    return caption.empty() ? Value::blank() : text_value(cells, caption);
  }
  if (value.is_blank() && table.missing_caption().has_value() && !table.missing_caption()->empty()) {
    return *table.missing_caption() == "-" ? Value::number(0.0) : text_value(cells, *table.missing_caption());
  }
  return reify_value(cells, value);
}

// A value cell of data field `df`, as Excel shows `value`.
void append_value_cell(PivotCells& cells, std::uint32_t row, std::uint32_t col, const Value& value, PivotCellKind kind,
                       const PivotTable& table, std::size_t df) {
  append_df_cell(cells, row, col, shown_value(table, cells, value), kind, value_role(kind), table, df);
}

// A value cell left empty whatever the captions say: a group heading row's or
// an inserted blank row's.
void append_empty_value_cell(PivotCells& cells, std::uint32_t row, std::uint32_t col) {
  append_cell(cells, row, col, Value::blank(), PivotCellKind::Blank, PivotCellRole::DataValue, 0);
}

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

bool labels_equal(const std::vector<std::string>& a, const std::vector<std::string>& b) {
  return a.size() == b.size() && std::equal(a.begin(), a.end(), b.begin());
}

enum class EntryKind : std::uint8_t {
  Leaf,      ///< An evaluated leaf (one data field of it when Values sits on this axis).
  Subtotal,  ///< A group's subtotal.
  Group,     ///< A label-only row opening a group whose values sit on the rows below.
  Blank,     ///< An empty row closing an item block of an `insert_blank_row` field.
};

/// One row or column of the projected grid.
struct AxisEntry {
  /// Display path, outermost first. The Values level, when it sits on this
  /// axis, contributes the data-field name at its position.
  std::vector<std::string> labels;
  /// The node behind each of `labels`; null at the Values level.
  std::vector<const AxisHierarchyNode*> nodes;
  EntryKind kind = EntryKind::Leaf;
  std::size_t leaf_index = 0;      ///< Leaf: index into the evaluated leaves.
  std::size_t subtotal_index = 0;  ///< Subtotal: index into the result's subtotals.
  /// The data field this entry carries when Values sits on this axis.
  std::optional<std::size_t> data_field;
  /// Subtotal over a group that contains the Values level: Excel names it
  /// "<group> <data field>" in place of the localized subtotal wording.
  std::string subtotal_label;
  /// Subtotal that doubles as its group's heading row, labelled bare. A
  /// collapsed item's subtotals are such entries in every layout and on
  /// either axis.
  bool heads_group = false;
  /// Custom subtotal function of a labelled subtotal: Excel names the row
  /// "<group> <function>" in place of the localized subtotal wording.
  std::optional<SubtotalFn> custom_fn;
};

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

/// How one axis is projected into entries.
struct AxisShape {
  std::size_t depth = 0;                              ///< Real fields on the axis.
  std::optional<std::size_t> values_at;               ///< Level the Values pseudo-field occupies, if on this axis.
  bool group_rows = false;                            ///< Compact / outline row axis: groups open with their own row.
  std::vector<bool> subtotal_first;                   ///< Per depth: the subtotal leads its group.
  std::vector<std::size_t> subtotal_counts;           ///< Per depth: subtotals the evaluator emitted.
  std::vector<std::vector<SubtotalFn>> subtotal_fns;  ///< Per depth: custom functions, in row order.
  std::vector<const PivotField*> fields;              ///< Per depth: the field, or null when out of range.
  std::vector<bool> blank_after;                      ///< Per depth: a blank row closes each item block.
};

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

Expected<void, Error> validate_result_shape(const PivotTable& table, const PivotResult& result,
                                            std::size_t row_leaf_count, std::size_t col_leaf_count) {
  if (result.values.size() != row_leaf_count) {
    return make_error(
        FormulonErrorCode::kEvalPivotInvalid, "pivot layout: row leaf count does not match values",
        "rows=" + std::to_string(row_leaf_count) + " values.rows=" + std::to_string(result.values.size()));
  }
  for (std::size_t r = 0; r < row_leaf_count; ++r) {
    if (result.values[r].size() != col_leaf_count) {
      return make_error(FormulonErrorCode::kEvalPivotInvalid, "pivot layout: col leaf count does not match values",
                        "row=" + std::to_string(r) + " cols=" + std::to_string(col_leaf_count) +
                            " values.cols=" + std::to_string(result.values[r].size()));
    }
    for (std::size_t c = 0; c < col_leaf_count; ++c) {
      if (result.values[r][c].size() != table.data_fields().size()) {
        return make_error(FormulonErrorCode::kEvalPivotInvalid, "pivot layout: data field count does not match values",
                          "row=" + std::to_string(r) + " col=" + std::to_string(c) +
                              " data_fields=" + std::to_string(table.data_fields().size()) +
                              " values.data_fields=" + std::to_string(result.values[r][c].size()));
      }
    }
  }
  return Expected<void, Error>::Ok();
}

/// Per-call state derived once from the table, result and options, then only
/// read by the emission phases.
struct LayoutPlan {
  LayoutPlan(const PivotTable& t, const PivotResult& r, const PivotLayoutOptions& o)
      : table(t), result(r), options(o) {}

  const PivotTable& table;
  const PivotResult& result;
  const PivotLayoutOptions& options;

  std::size_t row_depth = 0;
  std::size_t col_depth = 0;
  std::size_t data_field_count = 0;

  bool tabular = false;
  bool outline = false;
  bool multi_col_layout = false;
  bool compact = false;
  bool legacy = false;
  std::string subtotal_suffix;
  std::string grand_total_label;  ///< The table's caption, else the locale's word.

  bool values_on_rows = false;
  bool values_on_cols = false;
  AxisShape row_shape;
  AxisShape col_shape;
  std::size_t row_levels = 0;
  std::size_t col_levels = 0;

  std::vector<AxisLeaf> row_leaves;
  std::vector<AxisLeaf> col_leaves;
  std::vector<AxisEntry> row_entries;
  std::vector<AxisEntry> col_entries;
  std::size_t data_cols = 0;

  std::size_t row_header_cols = 0;
  std::size_t header_rows = 0;
  /// Compact / tabular / outline with column fields: the row above the column
  /// levels holding the "Column Labels" caption or the level names.
  bool caption_row = false;
  /// `gridDropZones` row above the header block.
  std::size_t drop_zone_rows = 0;
  bool emit_grand_totals_rows_strip = false;
  bool emit_grand_totals_cols_strip = false;
  std::size_t total_strip_cols = 0;
  std::size_t total_strip_rows = 0;
  std::size_t page_rows = 0;
  std::size_t page_lines = 0;  ///< Page fields per column (down-then-over) or row (over-then-down).
  std::size_t total_rows = 0;
  std::size_t total_cols = 0;
  std::size_t cell_count = 0;
  std::size_t corner_extent = 0;

  std::uint32_t cells_top = 0;
  std::uint32_t left = 0;
  std::uint32_t top = 0;
  std::uint32_t header_top = 0;    ///< First row of the header block, below any drop-zone row.
  std::uint32_t col_rows_top = 0;  ///< First column-level row.
  std::uint32_t data_top = 0;
  std::uint32_t data_left = 0;
  std::uint32_t total_left = 0;
  std::uint32_t row_header_row = 0;
  std::uint32_t total_header_row = 0;
};

Expected<LayoutPlan, Error> build_layout_plan(const PivotTable& table, const PivotResult& result,
                                              const PivotLayoutOptions& options) {
  LayoutPlan plan(table, result, options);
  const std::size_t row_depth = table.row_field_order().size();
  const std::size_t col_depth = table.col_field_order().size();
  const std::size_t data_field_count = table.data_fields().size();

  // Compact-form rendering uses Excel's "Row Labels" / "Column Labels"
  // placeholders in the corner of the header block and drops the row-
  // field-name row that classical English defaults emit. We treat any
  // non-empty `row_labels_label` as the locale opting in to this mode.
  const bool locale_opted_in = !options.row_labels_label.empty();
  // Tabular and Outline layouts are honoured only when the locale opts
  // into Excel-style rendering and the pivot has row fields (there is
  // no per-row-field column to give its own header without one).
  const bool tabular = locale_opted_in && table.layout() == PivotLayout::Tabular && row_depth > 0;
  const bool outline = locale_opted_in && table.layout() == PivotLayout::Outline && row_depth > 0;
  const bool multi_col_layout = tabular || outline;
  const bool compact = locale_opted_in && !multi_col_layout;
  const bool legacy = !compact && !multi_col_layout;
  plan.subtotal_suffix = options.subtotal_suffix.empty() && options.subtotal_prefix.empty()
                             ? std::string(" ") + options.grand_total_label
                             : options.subtotal_suffix;
  // The caption renames only the grand total; subtotals keep the locale's word.
  const std::optional<std::string>& caption = table.grand_total_caption();
  plan.grand_total_label = caption.has_value() && !caption->empty() ? *caption : options.grand_total_label;

  // With several data fields the Values pseudo-field is one level of the
  // row or column axis, at the position its `x="-2"` marker holds; a pivot
  // that records no marker keeps it innermost on the columns.
  const bool values_on_rows = data_field_count > 1 && table.row_values_position().has_value();
  const bool values_on_cols = data_field_count > 1 && !values_on_rows;
  AxisShape row_shape;
  row_shape.depth = row_depth;
  if (values_on_rows) {
    row_shape.values_at = std::min(*table.row_values_position(), row_depth);
  }
  row_shape.group_rows = compact || outline;
  // Compact and outline layouts honour the owner field's `subtotal_top`
  // flag. Tabular (and the historical English projection) retains its
  // established below-group order regardless of that flag.
  row_shape.subtotal_first.assign(row_depth, false);
  if (compact || outline) {
    for (std::size_t depth = 0; depth < row_depth; ++depth) {
      const std::uint32_t field_index = table.row_field_order()[depth];
      if (field_index < table.fields().size()) {
        row_shape.subtotal_first[depth] = table.fields()[field_index].subtotal_top;
      }
    }
  }
  row_shape.subtotal_counts = subtotal_counts_for_axis(table, table.row_field_order());
  row_shape.subtotal_fns = subtotal_fns_for_axis(table, table.row_field_order());
  row_shape.fields = fields_for_axis(table, table.row_field_order());
  // Blank rows follow only the blocks of an outer row field.
  row_shape.blank_after.assign(row_depth, false);
  for (std::size_t depth = 0; depth + 1 < row_depth; ++depth) {
    row_shape.blank_after[depth] = row_shape.fields[depth] != nullptr && row_shape.fields[depth]->insert_blank_row;
  }
  AxisShape col_shape;
  col_shape.depth = col_depth;
  if (values_on_cols) {
    col_shape.values_at = std::min(table.col_values_position().value_or(col_depth), col_depth);
  }
  col_shape.subtotal_counts = subtotal_counts_for_axis(table, table.col_field_order());
  col_shape.subtotal_fns = subtotal_fns_for_axis(table, table.col_field_order());
  col_shape.fields = fields_for_axis(table, table.col_field_order());
  const std::size_t row_levels = row_depth + (values_on_rows ? 1 : 0);
  const std::size_t col_levels = col_depth + (values_on_cols ? 1 : 0);

  std::vector<AxisLeaf> row_leaves = collect_axis_leaves(result.rows, row_depth);
  std::vector<AxisLeaf> col_leaves = collect_axis_leaves(result.cols, col_depth);
  auto valid_or = validate_result_shape(table, result, row_leaves.size(), col_leaves.size());
  if (!valid_or) {
    return std::move(valid_or.error());
  }
  auto row_entries_or = collect_axis_entries(table, result.rows, result.row_subtotals, row_shape, "row");
  if (!row_entries_or) {
    return std::move(row_entries_or.error());
  }
  std::vector<AxisEntry> row_entries = row_entries_or.take();
  auto col_entries_or = collect_axis_entries(table, result.cols, result.col_subtotals, col_shape, "column");
  if (!col_entries_or) {
    return std::move(col_entries_or.error());
  }
  // An axis leaf without a data field is a label-only pivot, not a value column.
  std::vector<AxisEntry> col_entries = data_field_count == 0 ? std::vector<AxisEntry>{} : col_entries_or.take();
  const std::size_t data_cols = col_entries.size();

  // Compact form merges every row level into one physical column (Excel
  // indents nested keys); Tabular / Outline give each level its own column,
  // the Values level included. A compact pivot with no row levels whose
  // data fields head the columns has no row-label column at all.
  std::size_t row_header_cols = row_levels;
  if (row_levels == 0) {
    row_header_cols = compact && values_on_cols ? 0 : 1;
  } else if (compact) {
    row_header_cols = 1;
  }
  // Compact / Tabular / Outline: a pivot without column fields has one
  // header row; with them, a row above the column levels holds the
  // "Column Labels" placeholder (Compact) or the column level names
  // (Tabular / Outline). The English projection stacks the column levels
  // over a row of row-field names.
  std::size_t header_rows =
      legacy ? std::max<std::size_t>(col_levels, 1) + 1 : (col_depth == 0 ? std::size_t{1} : col_levels + 1);
  // Hidden field headers drop the caption row, unless a lone data field's
  // caption occupies its corner.
  bool caption_row = !legacy && col_depth > 0;
  if (caption_row && !table.show_headers() && !(data_field_count == 1 && row_depth > 0)) {
    caption_row = false;
    --header_rows;
  }
  // The classic layout adds a drop-zone row over a report without column fields.
  const std::size_t drop_zone_rows = !legacy && col_depth == 0 && table.grid_drop_zones() ? 1 : 0;
  header_rows += drop_zone_rows;
  // Compact form folds the grand-totals strip in axes that have no
  // hierarchy of their own: a no-column-fields pivot's per-row total
  // already lives in its data columns, and likewise a no-row-fields
  // pivot's per-column total. Tabular / Outline share that rule; the
  // English projection keeps emitting both strips.
  const bool emit_grand_totals_rows_strip =
      data_field_count > 0 && table.grand_totals_rows() && (legacy || col_depth > 0);
  const bool emit_grand_totals_cols_strip = table.grand_totals_cols() && (legacy || row_depth > 0);
  // The Values axis totals each data field on its own: one strip column
  // (or row) per field, named after it.
  const std::size_t total_strip_cols = emit_grand_totals_rows_strip ? (values_on_cols ? data_field_count : 1) : 0;
  const std::size_t total_strip_rows = emit_grand_totals_cols_strip ? (values_on_rows ? data_field_count : 1) : 0;
  // Page (report filter) fields are drawn above the report proper: one row
  // per field, then a blank row separating the block from the row/column
  // headers. Excel keeps both outside the pivot's `ref`: the report body
  // starts at the anchor and the block sits above it (pushing the body down
  // only when the anchor leaves no room).
  // `pageWrap` folds the fields into lines of that many (0 is one line), down
  // then over or over then down; each page column after the first sits one
  // empty column to the right. The block is at least `kPageHeaderCols` wide
  // even over a narrower report, because the selection cell has to land
  // somewhere.
  const std::size_t page_count = result.page_selections.size();
  const std::size_t page_lines =
      table.page_wrap() == 0 ? page_count : std::min<std::size_t>(table.page_wrap(), page_count);
  const std::size_t page_line_count = page_lines == 0 ? 0 : (page_count + page_lines - 1) / page_lines;
  const std::size_t page_grid_rows = table.page_over_then_down() ? page_line_count : page_lines;
  const std::size_t page_grid_cols = table.page_over_then_down() ? page_lines : page_line_count;
  const std::size_t page_rows = page_count == 0 ? 0 : page_grid_rows + 1;
  const std::size_t page_cols = page_count == 0 ? 0 : page_grid_cols * (kPageHeaderCols + 1) - 1;
  const std::size_t body_rows = header_rows + row_entries.size() + total_strip_rows;
  const std::size_t body_cols = row_header_cols + data_cols + total_strip_cols;
  const std::size_t total_rows = page_rows + body_rows;
  const std::size_t total_cols = std::max(body_cols, page_cols);

  if (total_rows > UINT32_MAX || total_cols > UINT32_MAX) {
    return make_error(FormulonErrorCode::kEvalPivotInvalid, "pivot layout: projected bounds exceed uint32_t",
                      "rows=" + std::to_string(total_rows) + " cols=" + std::to_string(total_cols));
  }
  auto cell_count_or = checked_mul_size_t(total_rows, total_cols);
  if (!cell_count_or) {
    return std::move(cell_count_or.error());
  }

  const std::uint32_t body_top = std::max(table.anchor_row(), static_cast<std::uint32_t>(page_rows));
  plan.cells_top = body_top - static_cast<std::uint32_t>(page_rows);
  plan.cell_count = cell_count_or.value();

  plan.left = table.anchor_col();
  plan.top = body_top;
  plan.header_top = plan.top + static_cast<std::uint32_t>(drop_zone_rows);
  plan.col_rows_top = plan.header_top + (caption_row ? 1U : 0U);
  plan.data_top = plan.top + static_cast<std::uint32_t>(header_rows);
  plan.data_left = plan.left + static_cast<std::uint32_t>(row_header_cols);
  plan.total_left = plan.data_left + static_cast<std::uint32_t>(data_cols);
  plan.row_header_row = plan.top + static_cast<std::uint32_t>(header_rows - 1);
  // Where the grand-totals-rows strip's own header sits: beside the
  // outermost column level (the row "Q1 集計" occupies) once a column
  // hierarchy exists, else beside the row-field names.
  plan.total_header_row = col_depth > 0 ? plan.col_rows_top : plan.row_header_row;
  plan.corner_extent = data_cols + total_strip_cols;

  plan.row_depth = row_depth;
  plan.col_depth = col_depth;
  plan.data_field_count = data_field_count;
  plan.tabular = tabular;
  plan.outline = outline;
  plan.multi_col_layout = multi_col_layout;
  plan.compact = compact;
  plan.legacy = legacy;
  plan.values_on_rows = values_on_rows;
  plan.values_on_cols = values_on_cols;
  plan.row_shape = std::move(row_shape);
  plan.col_shape = std::move(col_shape);
  plan.row_levels = row_levels;
  plan.col_levels = col_levels;
  plan.row_leaves = std::move(row_leaves);
  plan.col_leaves = std::move(col_leaves);
  plan.row_entries = std::move(row_entries);
  plan.col_entries = std::move(col_entries);
  plan.data_cols = data_cols;
  plan.row_header_cols = row_header_cols;
  plan.header_rows = header_rows;
  plan.caption_row = caption_row;
  plan.drop_zone_rows = drop_zone_rows;
  plan.emit_grand_totals_rows_strip = emit_grand_totals_rows_strip;
  plan.emit_grand_totals_cols_strip = emit_grand_totals_cols_strip;
  plan.total_strip_cols = total_strip_cols;
  plan.total_strip_rows = total_strip_rows;
  plan.page_rows = page_rows;
  plan.page_lines = page_lines;
  plan.total_rows = total_rows;
  plan.total_cols = total_cols;
  return plan;
}

std::string plan_subtotal_label(const LayoutPlan& plan, const std::string& group) {
  return plan.options.subtotal_prefix + group + plan.subtotal_suffix;
}

// A custom subtotal row puts its function word where the locale puts the
// automatic subtotal word: after the item, or before it in prefix locales.
std::string entry_subtotal_label(const LayoutPlan& plan, const AxisEntry& entry, const std::string& group) {
  if (!entry.subtotal_label.empty()) {
    return entry.subtotal_label;
  }
  if (entry.custom_fn.has_value()) {
    const std::string word(subtotal_function_label(
        static_cast<Aggregation>(static_cast<std::uint8_t>(*entry.custom_fn)), plan.options.profile));
    return plan.options.subtotal_prefix.empty() ? group + " " + word : word + " " + group;
  }
  return plan_subtotal_label(plan, group);
}

std::string data_total_label(const LayoutPlan& plan, std::size_t df) {
  return plan.options.data_total_prefix + data_field_name(plan.table, df) + plan.options.data_total_suffix;
}

// A field caption, or nothing while `showHeaders` hides them.
std::string field_caption(const LayoutPlan& plan, std::string text) {
  return plan.table.show_headers() ? std::move(text) : std::string();
}

// Compact form indents a level-`level` row label by `indent + 1` per level.
void set_compact_indent(const LayoutPlan& plan, PivotCells& cells, std::size_t level) {
  cells.cells.back().label_indent = static_cast<std::uint32_t>(level) * (plan.table.indent() + 1U);
}

// Display name of the row / column level at `level`: a field's display
// name, or the stored Values caption for the Values level.
std::string level_name(const LayoutPlan& plan, const std::vector<std::uint32_t>& order,
                       const std::optional<std::size_t>& values_at, std::size_t level) {
  const PivotTable& table = plan.table;
  if (values_at.has_value()) {
    if (level == *values_at) {
      return table.data_caption();
    }
    if (level > *values_at) {
      --level;
    }
  }
  if (level < order.size() && order[level] < table.fields().size()) {
    return pivot_field_display_name(table.fields()[order[level]]);
  }
  return std::string();
}

std::string row_level_name(const LayoutPlan& plan, std::size_t level) {
  return level < plan.row_levels ? level_name(plan, plan.table.row_field_order(), plan.row_shape.values_at, level)
                                 : std::string();
}

std::string col_level_name(const LayoutPlan& plan, std::size_t level) {
  return level < plan.col_levels ? level_name(plan, plan.table.col_field_order(), plan.col_shape.values_at, level)
                                 : std::string();
}

// The field at `shape`'s level `level`, or null at the Values level.
const PivotField* axis_level_field(const AxisShape& shape, std::size_t level) {
  if (shape.values_at.has_value()) {
    if (level == *shape.values_at) {
      return nullptr;
    }
    if (level > *shape.values_at) {
      --level;
    }
  }
  return level < shape.fields.size() ? shape.fields[level] : nullptr;
}

// Number format of the item at `entry.labels[level]`, for a cell drawing that label bare.
std::string label_format(const AxisEntry& entry, std::size_t level) {
  return level < entry.nodes.size() && entry.nodes[level] != nullptr ? entry.nodes[level]->number_format
                                                                     : std::string();
}

void emit_text_or_blank(PivotCells& cells, std::uint32_t row, std::uint32_t col, const std::string& text,
                        PivotCellKind kind, PivotCellRole role, std::uint32_t depth, std::string field_name = {},
                        std::string number_format = {}) {
  if (text.empty()) {
    append_cell(cells, row, col, Value::blank(), PivotCellKind::Blank, PivotCellRole::None, depth,
                std::move(field_name));
  } else {
    append_cell(cells, row, col, text_value(cells, text), kind, role, depth, std::move(field_name),
                std::move(number_format));
  }
}

// The data field a cell at (`row_entry`, `col_entry`) reports: whichever
// axis carries the Values level decides, and a single field is field 0.
std::size_t cell_data_field(const AxisEntry& row_entry, const AxisEntry& col_entry) {
  if (row_entry.data_field) {
    return *row_entry.data_field;
  }
  return col_entry.data_field.value_or(0);
}

// Label of a column with no column level above it: its data field's name,
// or nothing when the data fields sit on the rows.
std::string flat_col_header(const LayoutPlan& plan, std::size_t c_entry) {
  const AxisEntry& entry = plan.col_entries[c_entry];
  if (entry.data_field) {
    return data_field_name(plan.table, *entry.data_field);
  }
  return plan.data_field_count == 1 ? data_field_name(plan.table, 0) : std::string();
}

// The page block: field name, then the item it is showing. Both are
// header furniture, so neither earns a `PivotCellKind` of its own — the
// vocabulary is mirrored by the C ABI and the WASM enum, and a page cell
// is not something a renderer has to draw differently from a header.
// Every other cell of the block, the separator row included, is an explicit
// blank so the rendered extent is the rectangle Excel reports.
void emit_page_block(const LayoutPlan& plan, PivotCells& cells) {
  const std::vector<PivotPageSelection>& pages = plan.result.page_selections;
  if (plan.page_rows == 0) {
    return;
  }
  const std::size_t stride = kPageHeaderCols + 1;
  for (std::size_t r = 0; r < plan.page_rows; ++r) {
    const std::uint32_t row = cells.top + static_cast<std::uint32_t>(r);
    for (std::size_t c = 0; c < plan.total_cols; ++c) {
      const std::uint32_t col = plan.left + static_cast<std::uint32_t>(c);
      const std::size_t line = plan.table.page_over_then_down() ? r : c / stride;
      const std::size_t pos = plan.table.page_over_then_down() ? c / stride : r;
      const std::size_t i = line * plan.page_lines + pos;
      if (r + 1 == plan.page_rows || c % stride == kPageHeaderCols || pos >= plan.page_lines || i >= pages.size()) {
        append_blank(cells, row, col);
      } else if (c % stride == 0) {
        emit_text_or_blank(cells, row, col, field_caption(plan, pages[i].field_label), PivotCellKind::Header,
                           PivotCellRole::PageCaption, 0, pages[i].field_label);
      } else {
        append_cell(cells, row, col, text_value(cells, pages[i].item_label), PivotCellKind::Header,
                    PivotCellRole::PageValue, 0, pages[i].field_label);
      }
    }
  }
}

// Emits one column-level label row at `depth`. Excel shows a non-leaf
// level's label once per contiguous run and blanks the repeats; the leaf-
// most level always shows its own value. A subtotal column carries its
// label on its own group's level and blanks below it. The grand-totals-
// rows strip is left blank except on `total_header_row`.
void emit_col_hierarchy_row(const LayoutPlan& plan, PivotCells& cells, std::uint32_t row, std::size_t depth,
                            bool blank_repeats) {
  const PivotTable& table = plan.table;
  const std::vector<AxisEntry>& col_entries = plan.col_entries;
  const bool leaf_depth = depth + 1 >= plan.col_levels;
  for (std::size_t c_entry = 0; c_entry < col_entries.size(); ++c_entry) {
    const AxisEntry& entry = col_entries[c_entry];
    // A collapsed item's columns are labelled like a leaf's.
    const bool subtotal_here =
        entry.kind == EntryKind::Subtotal && !entry.heads_group && depth + 1 == entry.labels.size();
    std::string label;
    if (depth < entry.labels.size()) {
      label = entry.labels[depth];
    }
    const PivotField* field = axis_level_field(plan.col_shape, depth);
    if (subtotal_here) {
      label = entry_subtotal_label(plan, entry, label);
    } else if (blank_repeats && !leaf_depth && c_entry > 0 && (field == nullptr || !field->fill_down_labels)) {
      const std::vector<std::string>& prev = col_entries[c_entry - 1].labels;
      if (prev.size() > depth && entry.labels.size() > depth &&
          std::equal(prev.begin(), prev.begin() + static_cast<std::ptrdiff_t>(depth + 1), entry.labels.begin())) {
        label.clear();
      }
    }
    const std::uint32_t col = plan.data_left + static_cast<std::uint32_t>(c_entry);
    if (!subtotal_here && plan.col_shape.values_at == depth && entry.data_field) {
      // The Values level heads its column with the data field's name.
      emit_text_or_blank(cells, row, col, label, PivotCellKind::Header, PivotCellRole::DataHeader,
                         static_cast<std::uint32_t>(depth), data_field_name(table, *entry.data_field),
                         data_field_format(table, *entry.data_field));
      continue;
    }
    emit_text_or_blank(cells, row, col, label,
                       entry.kind == EntryKind::Subtotal ? PivotCellKind::ColSubtotal : PivotCellKind::ColLabel,
                       subtotal_here ? PivotCellRole::SubtotalLabel : PivotCellRole::ColItem,
                       static_cast<std::uint32_t>(depth), col_level_name(plan, depth),
                       subtotal_here ? std::string() : label_format(entry, depth));
  }
  if (row != plan.total_header_row) {
    for (std::size_t s = 0; s < plan.total_strip_cols; ++s) {
      append_blank(cells, row, plan.total_left + static_cast<std::uint32_t>(s));
    }
  }
}

// The corner of a pivot with column fields names its data field; with
// several, the Values level names them instead and the corner is blank.
void emit_corner(const LayoutPlan& plan, PivotCells& cells) {
  if (plan.data_field_count == 1 && plan.row_depth > 0) {
    append_df_cell(cells, plan.header_top, plan.left, text_value(cells, data_field_name(plan.table, 0)),
                   PivotCellKind::Header, PivotCellRole::FieldHeader, plan.table, 0);
  } else {
    append_blank(cells, plan.header_top, plan.left);
  }
}

// The `gridDropZones` row: the Values caption over the first data column when
// the data fields head the columns.
void emit_drop_zone_row(const LayoutPlan& plan, PivotCells& cells) {
  if (plan.drop_zone_rows == 0) {
    return;
  }
  const std::size_t width = plan.row_header_cols + plan.data_cols + plan.total_strip_cols;
  for (std::size_t c = 0; c < width; ++c) {
    const std::uint32_t col = plan.left + static_cast<std::uint32_t>(c);
    if (col == plan.data_left && plan.values_on_cols) {
      emit_text_or_blank(cells, plan.top, col, field_caption(plan, plan.table.data_caption()), PivotCellKind::Header,
                         PivotCellRole::FieldHeader, 0);
    } else {
      append_blank(cells, plan.top, col);
    }
  }
}

void emit_header_multi_col(const LayoutPlan& plan, PivotCells& cells) {
  const PivotTable& table = plan.table;
  const std::vector<AxisEntry>& col_entries = plan.col_entries;
  const std::uint32_t left = plan.left;
  const std::uint32_t top = plan.header_top;
  const std::uint32_t data_left = plan.data_left;
  const std::size_t row_header_cols = plan.row_header_cols;
  const std::size_t data_cols = plan.data_cols;
  const std::size_t col_levels = plan.col_levels;
  if (plan.col_depth == 0) {
    // One header row: the row level names, then the data field names.
    for (std::size_t d = 0; d < row_header_cols; ++d) {
      emit_text_or_blank(cells, top, left + static_cast<std::uint32_t>(d), field_caption(plan, row_level_name(plan, d)),
                         PivotCellKind::Header, PivotCellRole::FieldHeader, static_cast<std::uint32_t>(d));
    }
    for (std::size_t c_entry = 0; c_entry < data_cols; ++c_entry) {
      const std::size_t df = col_entries[c_entry].data_field.value_or(0);
      emit_text_or_blank(cells, top, data_left + static_cast<std::uint32_t>(c_entry), flat_col_header(plan, c_entry),
                         PivotCellKind::Header, PivotCellRole::DataHeader, 0, data_field_name(table, df),
                         data_field_format(table, df));
    }
  } else {
    // Caption row: the corner, then each column level's own name -- the
    // Values level under the stored caption -- where Compact draws its
    // single "Column Labels" placeholder.
    if (plan.caption_row) {
      emit_corner(plan, cells);
      for (std::size_t d = 1; d < row_header_cols; ++d) {
        append_blank(cells, top, left + static_cast<std::uint32_t>(d));
      }
      for (std::size_t i = 0; i < plan.corner_extent; ++i) {
        emit_text_or_blank(cells, top, data_left + static_cast<std::uint32_t>(i),
                           field_caption(plan, col_level_name(plan, i)), PivotCellKind::Header,
                           PivotCellRole::FieldHeader, 0);
      }
    }
    // The last column-level row also carries the row level names.
    for (std::size_t depth = 0; depth < col_levels; ++depth) {
      const std::uint32_t row = plan.col_rows_top + static_cast<std::uint32_t>(depth);
      for (std::size_t d = 0; d < row_header_cols; ++d) {
        emit_text_or_blank(cells, row, left + static_cast<std::uint32_t>(d),
                           depth + 1 == col_levels ? field_caption(plan, row_level_name(plan, d)) : std::string(),
                           PivotCellKind::Header, PivotCellRole::FieldHeader, static_cast<std::uint32_t>(d));
      }
      emit_col_hierarchy_row(plan, cells, row, depth, true);
    }
  }
}

void emit_header_compact(const LayoutPlan& plan, PivotCells& cells) {
  const PivotTable& table = plan.table;
  const PivotLayoutOptions& options = plan.options;
  const std::vector<AxisEntry>& col_entries = plan.col_entries;
  const std::uint32_t left = plan.left;
  const std::uint32_t top = plan.header_top;
  const std::uint32_t data_left = plan.data_left;
  const std::size_t row_header_cols = plan.row_header_cols;
  const std::size_t data_cols = plan.data_cols;
  const std::size_t col_levels = plan.col_levels;
  if (plan.col_depth == 0) {
    // One header row: the "Row Labels" placeholder, then the data field
    // names (blank when the data fields sit on the rows).
    if (row_header_cols > 0) {
      emit_text_or_blank(cells, top, left, field_caption(plan, options.row_labels_label), PivotCellKind::Header,
                         PivotCellRole::FieldHeader, 0);
    }
    for (std::size_t c_entry = 0; c_entry < data_cols; ++c_entry) {
      const std::size_t df = col_entries[c_entry].data_field.value_or(0);
      emit_text_or_blank(cells, top, data_left + static_cast<std::uint32_t>(c_entry), flat_col_header(plan, c_entry),
                         PivotCellKind::Header, PivotCellRole::DataHeader, 0, data_field_name(table, df),
                         data_field_format(table, df));
    }
  } else {
    // Caption row: the corner, then the "Column Labels" placeholder over
    // the first column. Without row fields the single data field's name
    // moves down to the data row instead (see the row loop below).
    if (plan.caption_row) {
      if (row_header_cols > 0) {
        emit_corner(plan, cells);
      }
      for (std::size_t i = 0; i < plan.corner_extent; ++i) {
        emit_text_or_blank(cells, top, data_left + static_cast<std::uint32_t>(i),
                           i == 0 ? field_caption(plan, options.column_labels_label) : std::string(),
                           PivotCellKind::Header, PivotCellRole::FieldHeader, 0);
      }
    }
    for (std::size_t depth = 0; depth < col_levels; ++depth) {
      const std::uint32_t row = plan.col_rows_top + static_cast<std::uint32_t>(depth);
      if (row_header_cols > 0 && (plan.row_levels == 0 || depth + 1 < col_levels)) {
        append_blank(cells, row, left);
      }
      emit_col_hierarchy_row(plan, cells, row, depth, true);
    }
    if (row_header_cols > 0 && plan.row_levels > 0) {
      emit_text_or_blank(cells, plan.row_header_row, left, field_caption(plan, options.row_labels_label),
                         PivotCellKind::Header, PivotCellRole::FieldHeader, 0);
    }
  }
}

void emit_header_legacy(const LayoutPlan& plan, PivotCells& cells) {
  // English projection: column levels stack from the top, each label
  // repeated per column, over a row of row level names.
  for (std::size_t depth = 0; depth < plan.row_header_cols; ++depth) {
    emit_text_or_blank(cells, plan.row_header_row, plan.left + static_cast<std::uint32_t>(depth),
                       field_caption(plan, row_level_name(plan, depth)), PivotCellKind::Header,
                       PivotCellRole::FieldHeader, static_cast<std::uint32_t>(depth));
  }
  if (plan.col_levels == 0) {
    for (std::size_t c_entry = 0; c_entry < plan.data_cols; ++c_entry) {
      emit_text_or_blank(cells, plan.top, plan.data_left + static_cast<std::uint32_t>(c_entry),
                         flat_col_header(plan, c_entry), PivotCellKind::ColLabel, PivotCellRole::DataHeader, 0);
    }
  }
  for (std::size_t depth = 0; depth < plan.col_levels; ++depth) {
    emit_col_hierarchy_row(plan, cells, plan.top + static_cast<std::uint32_t>(depth), depth, false);
  }
}

// Column hierarchy rows plus the row-level header, per layout mode.
void emit_header_block(const LayoutPlan& plan, PivotCells& cells) {
  emit_drop_zone_row(plan, cells);
  if (plan.multi_col_layout) {
    emit_header_multi_col(plan, cells);
  } else if (plan.compact) {
    emit_header_compact(plan, cells);
  } else {
    emit_header_legacy(plan, cells);
  }
}

// Row labels and data cells. Tabular tracks the previous leaf row's label
// so a parent path shared with the leaf row above is blanked, unless the
// parent's field fills its label down.
void emit_body_rows(const LayoutPlan& plan, PivotCells& cells) {
  const PivotTable& table = plan.table;
  const PivotResult& result = plan.result;
  const std::vector<AxisEntry>& row_entries = plan.row_entries;
  const std::vector<AxisEntry>& col_entries = plan.col_entries;
  const std::uint32_t left = plan.left;
  const std::uint32_t data_top = plan.data_top;
  const std::uint32_t data_left = plan.data_left;
  const std::size_t row_header_cols = plan.row_header_cols;
  const bool multi_col_layout = plan.multi_col_layout;
  const bool outline = plan.outline;
  const bool compact = plan.compact;
  const bool tabular = plan.tabular;
  std::vector<std::string> prev_row_labels(row_header_cols);
  std::vector<bool> prev_row_labels_filled(row_header_cols, false);
  for (std::size_t r_entry = 0; r_entry < row_entries.size(); ++r_entry) {
    const AxisEntry& entry = row_entries[r_entry];
    const std::uint32_t row = data_top + static_cast<std::uint32_t>(r_entry);
    if (entry.kind == EntryKind::Blank) {
      // Compact form writes the label cell of a blank row; the other layouts leave it out.
      for (std::size_t c = 0; c < row_header_cols; ++c) {
        append_cell(cells, row, left + static_cast<std::uint32_t>(c), Value::blank(), PivotCellKind::Blank,
                    compact ? PivotCellRole::BlankRow : PivotCellRole::None, 0);
        if (compact) {
          set_compact_indent(plan, cells, 0);
        }
      }
      for (std::size_t c_entry = 0; c_entry < col_entries.size(); ++c_entry) {
        append_empty_value_cell(cells, row, data_left + static_cast<std::uint32_t>(c_entry));
      }
      continue;
    }
    const std::vector<std::string>& labels = entry.labels;
    const bool subtotal = entry.kind == EntryKind::Subtotal;
    const std::size_t last_depth = labels.empty() ? 0 : labels.size() - 1;
    // Leading levels this row shares with the previous leaf row; a level
    // under a changed outer item starts a new block and shows its label.
    std::size_t shared = 0;
    while (shared < labels.size() && shared < row_header_cols && prev_row_labels_filled[shared] &&
           prev_row_labels[shared] == labels[shared]) {
      ++shared;
    }
    for (std::size_t depth = 0; depth < row_header_cols; ++depth) {
      std::string label;
      // The level whose item `label` shows as is, for its number format.
      std::optional<std::size_t> bare_level;
      PivotCellRole role = PivotCellRole::SubtotalLabel;
      if (multi_col_layout) {
        const PivotField* field = axis_level_field(plan.row_shape, depth);
        const bool fill_down = field != nullptr && field->fill_down_labels && depth < last_depth;
        if (subtotal || entry.kind == EntryKind::Group || outline) {
          // Subtotal and group rows, and every outline row, label only
          // their own level and the fill-down levels above it. A subtotal
          // takes the localized wording unless it doubles as the group's
          // heading row.
          if (depth == last_depth && depth < labels.size()) {
            label = labels[depth];
            if (subtotal && !entry.heads_group) {
              label = entry_subtotal_label(plan, entry, label);
            } else {
              bare_level = depth;
            }
          } else if (fill_down) {
            label = labels[depth];
            bare_level = depth;
          }
        } else if (depth < labels.size()) {
          // Tabular leaf rows repeat the path but blank a parent whose
          // whole path matches the previous leaf row's; the leaf-most level
          // always shows.
          const std::string& candidate = labels[depth];
          if (fill_down || depth + 1 == labels.size() || depth >= shared) {
            label = candidate;
            bare_level = depth;
          }
        }
      } else if (compact) {
        // Compact form shows each row's own level in the single label
        // column; a subtotal heading its group is the bare group label.
        if (subtotal && !entry.subtotal_label.empty()) {
          label = entry.subtotal_label;
        } else if (subtotal && !entry.heads_group && !labels.empty()) {
          label = entry_subtotal_label(plan, entry, labels.back());
        } else if (!labels.empty()) {
          label = labels.back();
          bare_level = last_depth;
        } else if (plan.row_levels == 0 && plan.col_depth > 0 && r_entry == 0) {
          // No-row-fields pivot: the data field's name heads the single
          // data row.
          label = data_field_name(table, 0);
          role = PivotCellRole::FieldHeader;
        }
      } else if (depth < labels.size()) {
        label = labels[depth];
        if (subtotal && depth == last_depth && !entry.heads_group) {
          label = entry_subtotal_label(plan, entry, label);
        } else {
          bare_level = depth;
        }
      }
      const PivotCellKind kind = subtotal ? PivotCellKind::RowSubtotal : PivotCellKind::RowLabel;
      // Compact folds every level into one column, so the cell reports the
      // row's hierarchy level rather than its physical column.
      const std::size_t level = compact ? last_depth : depth;
      if (bare_level.has_value()) {
        role = plan.row_shape.values_at == *bare_level ? PivotCellRole::DataRowLabel : PivotCellRole::RowItem;
      }
      if (label.empty()) {
        append_cell(cells, row, left + static_cast<std::uint32_t>(depth), Value::blank(), kind, PivotCellRole::None,
                    static_cast<std::uint32_t>(level), row_level_name(plan, level));
      } else {
        append_cell(cells, row, left + static_cast<std::uint32_t>(depth), text_value(cells, label), kind, role,
                    static_cast<std::uint32_t>(level), row_level_name(plan, level),
                    bare_level.has_value() ? label_format(entry, *bare_level) : std::string());
        if (compact && role != PivotCellRole::FieldHeader) {
          set_compact_indent(plan, cells, level);
        }
      }
      if (tabular && entry.kind == EntryKind::Leaf && depth < labels.size()) {
        prev_row_labels[depth] = labels[depth];
        prev_row_labels_filled[depth] = true;
      }
    }
    if (entry.kind == EntryKind::Group) {
      // A heading row without its subtotal still spans the data area as blanks.
      for (std::size_t c_entry = 0; c_entry < col_entries.size(); ++c_entry) {
        append_empty_value_cell(cells, row, data_left + static_cast<std::uint32_t>(c_entry));
      }
      continue;
    }
    for (std::size_t c_entry = 0; c_entry < col_entries.size(); ++c_entry) {
      const AxisEntry& col_entry = col_entries[c_entry];
      const std::size_t df = cell_data_field(entry, col_entry);
      const std::uint32_t col = data_left + static_cast<std::uint32_t>(c_entry);
      const Value* value = nullptr;
      PivotCellKind kind = PivotCellKind::Data;
      if (subtotal) {
        const RowSubtotal& row_subtotal = result.row_subtotals[entry.subtotal_index];
        kind = PivotCellKind::RowSubtotal;
        if (col_entry.kind == EntryKind::Subtotal &&
            col_entry.subtotal_index < row_subtotal.col_subtotal_values.size() &&
            df < row_subtotal.col_subtotal_values[col_entry.subtotal_index].size()) {
          value = &row_subtotal.col_subtotal_values[col_entry.subtotal_index][df];
        } else if (col_entry.kind == EntryKind::Leaf && col_entry.leaf_index < row_subtotal.col_values.size() &&
                   df < row_subtotal.col_values[col_entry.leaf_index].size()) {
          value = &row_subtotal.col_values[col_entry.leaf_index][df];
        } else if (df < row_subtotal.values.size()) {
          value = &row_subtotal.values[df];
        }
      } else if (col_entry.kind == EntryKind::Subtotal) {
        const ColSubtotal& col_subtotal = result.col_subtotals[col_entry.subtotal_index];
        kind = PivotCellKind::ColSubtotal;
        if (entry.leaf_index < col_subtotal.values.size() && df < col_subtotal.values[entry.leaf_index].size()) {
          value = &col_subtotal.values[entry.leaf_index][df];
        }
      } else {
        value = &result.values[entry.leaf_index][col_entry.leaf_index][df];
      }
      append_value_cell(cells, row, col, value != nullptr ? *value : Value::blank(), kind, table, df);
    }
  }
}

// Row totals: the strip at the right, from the evaluator's re-aggregated
// per-row totals so the projection remains cache-only.
void emit_row_totals_strip(const LayoutPlan& plan, PivotCells& cells) {
  if (!plan.emit_grand_totals_rows_strip) {
    return;
  }
  const PivotTable& table = plan.table;
  const PivotResult& result = plan.result;
  const std::vector<AxisEntry>& row_entries = plan.row_entries;
  const std::vector<AxisLeaf>& col_leaves = plan.col_leaves;
  const std::uint32_t data_top = plan.data_top;
  const std::uint32_t total_left = plan.total_left;
  const std::size_t total_strip_cols = plan.total_strip_cols;
  const bool values_on_cols = plan.values_on_cols;
  for (std::size_t r_entry = 0; r_entry < row_entries.size(); ++r_entry) {
    const AxisEntry& entry = row_entries[r_entry];
    const std::uint32_t row = data_top + static_cast<std::uint32_t>(r_entry);
    if (entry.kind == EntryKind::Group || entry.kind == EntryKind::Blank) {
      for (std::size_t s = 0; s < total_strip_cols; ++s) {
        append_empty_value_cell(cells, row, total_left + static_cast<std::uint32_t>(s));
      }
      continue;
    }
    for (std::size_t s = 0; s < total_strip_cols; ++s) {
      const std::size_t df = entry.data_field.value_or(values_on_cols ? s : 0);
      const std::uint32_t col = total_left + static_cast<std::uint32_t>(s);
      if (entry.kind == EntryKind::Subtotal) {
        // `RowSubtotal::values[df]` aggregates the group across the whole
        // column axis, which is exactly what this strip states.
        const RowSubtotal& subtotal = result.row_subtotals[entry.subtotal_index];
        append_value_cell(cells, row, col, df < subtotal.values.size() ? subtotal.values[df] : Value::blank(),
                          PivotCellKind::RowSubtotal, table, df);
        continue;
      }
      // Use the evaluator's per-row-leaf re-aggregation rather than
      // summing this row's cells: a non-additive function (Average/Max/
      // Min/StdDev/Var) cannot be recovered from the cell aggregates.
      const std::size_t leaf = entry.leaf_index;
      if (leaf < result.row_leaf_totals.size() && df < result.row_leaf_totals[leaf].size() &&
          !result.row_leaf_totals[leaf][df].is_blank()) {
        append_value_cell(cells, row, col, result.row_leaf_totals[leaf][df], PivotCellKind::GrandTotal, table, df);
        continue;
      }
      if (total_is_derived(table, df)) {
        append_value_cell(cells, row, col, Value::blank(), PivotCellKind::GrandTotal, table, df);
        continue;
      }
      // Every row of this strip states a total: a source error wins, an
      // all-numeric row sums, and an aggregate that cannot be summed
      // reports blank rather than leaving a hole in the total column.
      double sum = 0.0;
      bool summable = true;
      bool any_number = false;
      const Value* source_error = nullptr;
      for (std::size_t c_leaf = 0; c_leaf < col_leaves.size(); ++c_leaf) {
        const Value& v = result.values[leaf][c_leaf][df];
        if (v.is_error()) {
          source_error = &v;
          break;
        }
        if (v.is_number()) {
          sum += v.as_number();
          any_number = true;
        } else if (!v.is_blank()) {
          summable = false;
        }
      }
      Value total = Value::blank();
      if (source_error != nullptr) {
        total = *source_error;
      } else if (summable && any_number) {
        total = Value::number(sum);
      }
      append_value_cell(cells, row, col, total, PivotCellKind::GrandTotal, table, df);
    }
  }
  for (std::size_t s = 0; s < total_strip_cols; ++s) {
    const std::size_t df = values_on_cols ? s : 0;
    append_df_cell(cells, plan.total_header_row, total_left + static_cast<std::uint32_t>(s),
                   text_value(cells, values_on_cols ? data_total_label(plan, s) : plan.grand_total_label),
                   PivotCellKind::Header, PivotCellRole::GrandTotalLabel, table, df);
  }
}

// Column totals: the strip at the bottom.
void emit_col_totals_strip(const LayoutPlan& plan, PivotCells& cells) {
  const PivotTable& table = plan.table;
  const PivotResult& result = plan.result;
  const std::vector<AxisEntry>& row_entries = plan.row_entries;
  const std::vector<AxisEntry>& col_entries = plan.col_entries;
  const std::vector<AxisLeaf>& row_leaves = plan.row_leaves;
  const std::vector<AxisLeaf>& col_leaves = plan.col_leaves;
  const std::uint32_t left = plan.left;
  const std::uint32_t data_top = plan.data_top;
  const std::uint32_t data_left = plan.data_left;
  const std::uint32_t total_left = plan.total_left;
  const std::size_t row_header_cols = plan.row_header_cols;
  const std::size_t total_strip_cols = plan.total_strip_cols;
  for (std::size_t b = 0; b < plan.total_strip_rows; ++b) {
    const std::optional<std::size_t> row_df = plan.values_on_rows ? std::optional<std::size_t>(b) : std::nullopt;
    const std::uint32_t total_row = data_top + static_cast<std::uint32_t>(row_entries.size() + b);
    append_cell(cells, total_row, left, text_value(cells, row_df ? data_total_label(plan, b) : plan.grand_total_label),
                PivotCellKind::GrandTotal, PivotCellRole::GrandTotalLabel, 0);
    if (plan.compact) {
      set_compact_indent(plan, cells, 0);
    }
    // Tabular / Outline leave the remaining row-label columns blank so the
    // rendered grid is rectangular.
    if (plan.multi_col_layout) {
      for (std::size_t d = 1; d < row_header_cols; ++d) {
        append_cell(cells, total_row, left + static_cast<std::uint32_t>(d), Value::blank(), PivotCellKind::GrandTotal,
                    PivotCellRole::None, static_cast<std::uint32_t>(d));
      }
    }
    for (std::size_t c_entry = 0; c_entry < col_entries.size(); ++c_entry) {
      const AxisEntry& col_entry = col_entries[c_entry];
      const std::size_t df = row_df.value_or(col_entry.data_field.value_or(0));
      const std::uint32_t col = data_left + static_cast<std::uint32_t>(c_entry);
      const bool col_subtotal = col_entry.kind == EntryKind::Subtotal;
      // Prefer the evaluator's re-aggregated totals: summing per-row
      // aggregates is wrong for Average, Max, Min, StdDev and Var.
      if (plan.col_depth == 0 && !col_subtotal && df < result.grand_totals.size() &&
          !result.grand_totals[df].is_blank()) {
        append_value_cell(cells, total_row, col, result.grand_totals[df], PivotCellKind::GrandTotal, table, df);
        continue;
      }
      if (!col_subtotal && col_entry.leaf_index < result.col_leaf_totals.size() &&
          df < result.col_leaf_totals[col_entry.leaf_index].size() &&
          !result.col_leaf_totals[col_entry.leaf_index][df].is_blank()) {
        append_value_cell(cells, total_row, col, result.col_leaf_totals[col_entry.leaf_index][df],
                          PivotCellKind::GrandTotal, table, df);
        continue;
      }
      const PivotCellKind kind = col_subtotal ? PivotCellKind::ColSubtotal : PivotCellKind::GrandTotal;
      if (total_is_derived(table, df)) {
        const ColSubtotal* subtotal = col_subtotal ? &result.col_subtotals[col_entry.subtotal_index] : nullptr;
        append_value_cell(cells, total_row, col,
                          subtotal != nullptr && df < subtotal->total.size() ? subtotal->total[df] : Value::blank(),
                          kind, table, df);
        continue;
      }
      // Grand-total row x subtotal column: non-Sum takes the evaluator's 2-D combination.
      if (col_subtotal && df < table.data_fields().size()) {
        const ColSubtotal& subtotal = result.col_subtotals[col_entry.subtotal_index];
        const Aggregation aggregation =
            subtotal.aggregation.has_value() ? *subtotal.aggregation : table.data_fields()[df].aggregation;
        if ((aggregation != Aggregation::Sum || result.sum_totals_as_state) && df < subtotal.total.size() &&
            !subtotal.total[df].is_blank()) {
          append_value_cell(cells, total_row, col, subtotal.total[df], kind, table, df);
          continue;
        }
      }
      double sum = 0.0;
      bool numeric = true;
      bool any_number = false;
      const Value* source_error = nullptr;
      for (std::size_t r_leaf = 0; r_leaf < row_leaves.size(); ++r_leaf) {
        const Value* value = nullptr;
        if (col_subtotal) {
          const ColSubtotal& subtotal = result.col_subtotals[col_entry.subtotal_index];
          if (r_leaf < subtotal.values.size() && df < subtotal.values[r_leaf].size()) {
            value = &subtotal.values[r_leaf][df];
          }
        } else {
          value = &result.values[r_leaf][col_entry.leaf_index][df];
        }
        if (value == nullptr) {
          continue;
        }
        if (value->is_error()) {
          source_error = value;
          break;
        }
        if (value->is_number()) {
          sum += value->as_number();
          any_number = true;
        } else if (!value->is_blank()) {
          numeric = false;
        }
      }
      Value total = Value::blank();
      if (source_error != nullptr) {
        total = *source_error;
      } else if (numeric && any_number) {
        total = Value::number(sum);
      }
      append_value_cell(cells, total_row, col, total, kind, table, df);
    }
    for (std::size_t s = 0; s < total_strip_cols; ++s) {
      const std::size_t df = row_df.value_or(plan.values_on_cols ? s : 0);
      Value v = Value::blank();
      if (df < result.grand_totals.size() && !result.grand_totals[df].is_blank()) {
        v = result.grand_totals[df];
      } else if (df == 0 && !result.grand_total.is_blank()) {
        v = result.grand_total;
      } else if (!total_is_derived(table, df)) {
        double sum = 0.0;
        bool numeric = true;
        bool any_number = false;
        for (std::size_t r_leaf = 0; r_leaf < row_leaves.size(); ++r_leaf) {
          for (std::size_t c_leaf = 0; c_leaf < col_leaves.size(); ++c_leaf) {
            const Value& cell = result.values[r_leaf][c_leaf][df];
            if (cell.is_number()) {
              sum += cell.as_number();
              any_number = true;
            } else if (!cell.is_blank()) {
              numeric = false;
            }
          }
        }
        if (numeric && any_number) {
          v = Value::number(sum);
        }
      }
      append_value_cell(cells, total_row, total_left + static_cast<std::uint32_t>(s), v, PivotCellKind::GrandTotal,
                        table, df);
    }
  }
}

// True when row entry `entry` lies inside the block of `node`, the item at
// level `level`, below the item's own label row; its own subtotal does not.
bool inside_block(const AxisEntry& entry, std::size_t level, const AxisHierarchyNode* node) {
  return entry.nodes.size() > level + 1 && entry.nodes[level] == node;
}

// `mergeItem` in tabular form: an outer item's label spans the rows of its
// block when it has two or more, and a subtotal or grand-total label spans
// the label columns from its own to the last.
void emit_merges(const LayoutPlan& plan, PivotCells& cells) {
  if (!plan.tabular || !plan.table.merge_item()) {
    return;
  }
  const std::vector<AxisEntry>& entries = plan.row_entries;
  const std::size_t width = plan.row_header_cols;
  const auto col_at = [&](std::size_t c) { return plan.left + static_cast<std::uint32_t>(c); };
  const auto row_at = [&](std::size_t r) { return plan.data_top + static_cast<std::uint32_t>(r); };
  for (std::size_t r = 0; r < entries.size(); ++r) {
    const AxisEntry& entry = entries[r];
    for (std::size_t level = 0; level + 1 < entry.nodes.size() && level < width; ++level) {
      const AxisHierarchyNode* node = entry.nodes[level];
      if (node == nullptr || (r > 0 && inside_block(entries[r - 1], level, node))) {
        continue;
      }
      std::size_t end = r + 1;
      while (end < entries.size() && inside_block(entries[end], level, node)) {
        ++end;
      }
      if (end - r >= 2) {
        cells.merges.push_back(PivotRect{row_at(r), col_at(level), row_at(end - 1), col_at(level)});
      }
    }
    const std::size_t own = entry.labels.empty() ? 0 : entry.labels.size() - 1;
    if (entry.kind == EntryKind::Subtotal && !entry.heads_group && own + 1 < width) {
      cells.merges.push_back(PivotRect{row_at(r), col_at(own), row_at(r), col_at(width - 1)});
    }
  }
  for (std::size_t b = 0; width >= 2 && b < plan.total_strip_rows; ++b) {
    cells.merges.push_back(
        PivotRect{row_at(entries.size() + b), col_at(0), row_at(entries.size() + b), col_at(width - 1)});
  }
}

// Appends the items on `entry`'s path, outermost first, skipping the Values level.
Expected<void, Error> append_item_refs(const AxisEntry& entry, const std::vector<std::uint32_t>& field_order,
                                       std::vector<PivotItemRef>& out, std::deque<std::string>& text_storage) {
  std::size_t depth = 0;
  for (const AxisHierarchyNode* node : entry.nodes) {
    if (node == nullptr) {
      continue;
    }
    if (!node->identity.has_value()) {
      return make_error(FormulonErrorCode::kEvalPivotInvalid, "pivot cell item path: item has no identity",
                        "label=" + node->label);
    }
    PivotItemRef ref;
    ref.field_index = field_order[depth++];
    ref.identity = *node->identity;
    if (ref.identity.is_text()) {
      text_storage.emplace_back(ref.identity.as_text());
      ref.identity = Value::text(text_storage.back());
    }
    ref.is_group_item = node->identity_is_group_item;
    out.push_back(std::move(ref));
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<PivotCells, Error> layout(const PivotTable& table, const PivotResult& result,
                                   const PivotLayoutOptions& options) {
  auto plan_or = build_layout_plan(table, result, options);
  if (!plan_or) {
    return std::move(plan_or.error());
  }
  const LayoutPlan plan = plan_or.take();

  PivotCells cells;
  cells.top = plan.cells_top;
  cells.left = plan.left;
  cells.rows = static_cast<std::uint32_t>(plan.total_rows);
  cells.cols = static_cast<std::uint32_t>(plan.total_cols);
  cells.cells.reserve(plan.cell_count);

  emit_page_block(plan, cells);
  emit_header_block(plan, cells);
  emit_body_rows(plan, cells);
  emit_row_totals_strip(plan, cells);
  emit_col_totals_strip(plan, cells);
  emit_merges(plan, cells);
  return cells;
}

Expected<CellItemPath, Error> cell_item_path(const PivotTable& table, const PivotResult& result,
                                             const PivotLayoutOptions& options, std::uint32_t row, std::uint32_t col) {
  auto plan_or = build_layout_plan(table, result, options);
  if (!plan_or) {
    return std::move(plan_or.error());
  }
  const LayoutPlan plan = plan_or.take();
  const auto not_a_value_cell = [&]() {
    return make_error(FormulonErrorCode::kInvalidArgument, "pivot cell item path: not a value cell of the report",
                      "row=" + std::to_string(row) + " col=" + std::to_string(col));
  };
  if (row < plan.data_top || col < plan.data_left) {
    return not_a_value_cell();
  }
  const std::size_t r = row - plan.data_top;
  const std::size_t c = col - plan.data_left;
  if (r >= plan.row_entries.size() + plan.total_strip_rows || c >= plan.data_cols + plan.total_strip_cols) {
    return not_a_value_cell();
  }
  // Outside the body the cell sits in a grand-total strip, which lists no items on that axis.
  const AxisEntry* row_entry = r < plan.row_entries.size() ? &plan.row_entries[r] : nullptr;
  const AxisEntry* col_entry = c < plan.data_cols ? &plan.col_entries[c] : nullptr;
  if (row_entry != nullptr && (row_entry->kind == EntryKind::Group || row_entry->kind == EntryKind::Blank)) {
    return not_a_value_cell();
  }
  const bool row_subtotal = row_entry != nullptr && row_entry->kind == EntryKind::Subtotal;
  const bool col_subtotal = col_entry != nullptr && col_entry->kind == EntryKind::Subtotal;
  CellItemPath path;
  if (row_subtotal) {
    path.kind = PivotCellKind::RowSubtotal;
  } else if (col_subtotal) {
    path.kind = PivotCellKind::ColSubtotal;
  } else if (row_entry != nullptr && col_entry != nullptr) {
    path.kind = PivotCellKind::Data;
  } else {
    path.kind = PivotCellKind::GrandTotal;
  }
  if (row_entry != nullptr) {
    auto rows_or = append_item_refs(*row_entry, table.row_field_order(), path.row_items, path.text_storage);
    if (!rows_or) {
      return std::move(rows_or.error());
    }
  }
  if (col_entry != nullptr) {
    auto cols_or = append_item_refs(*col_entry, table.col_field_order(), path.col_items, path.text_storage);
    if (!cols_or) {
      return std::move(cols_or.error());
    }
  }
  return path;
}

}  // namespace formulon::pivot
