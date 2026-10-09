//
// Implementation of the PivotResult -> grid-cell projection.

#include "pivot/pivot_layout.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "pivot/field_lookup.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
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

std::string data_field_name(const PivotTable& table, std::size_t index) {
  if (index >= table.data_fields().size()) {
    return {};
  }
  return table.data_fields()[index].name;
}

std::string data_field_format(const PivotTable& table, std::size_t index) {
  if (index >= table.data_fields().size()) {
    return {};
  }
  return table.data_fields()[index].number_format;
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
                 std::uint32_t depth, std::string field_name = {}, std::string number_format = {}) {
  PivotCell cell;
  cell.row = row;
  cell.col = col;
  cell.value = value;
  cell.kind = kind;
  cell.depth = depth;
  cell.field_name = std::move(field_name);
  cell.number_format = std::move(number_format);
  cells.cells.push_back(std::move(cell));
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
};

/// One row or column of the projected grid.
struct AxisEntry {
  /// Display path, outermost first. The Values level, when it sits on this
  /// axis, contributes the data-field name at its position.
  std::vector<std::string> labels;
  EntryKind kind = EntryKind::Leaf;
  std::size_t leaf_index = 0;      ///< Leaf: index into the evaluated leaves.
  std::size_t subtotal_index = 0;  ///< Subtotal: index into the result's subtotals.
  /// The data field this entry carries when Values sits on this axis.
  std::optional<std::size_t> data_field;
  /// Subtotal over a group that contains the Values level: Excel names it
  /// "<group> <data field>" in place of the localized subtotal wording.
  std::string subtotal_label;
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
    counts[depth] = !field.subtotal_fns.empty() ? field.subtotal_fns.size() : (field.default_subtotal ? 1U : 0U);
  }
  return counts;
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
  std::size_t depth = 0;                     ///< Real fields on the axis.
  std::optional<std::size_t> values_at;      ///< Level the Values pseudo-field occupies, if on this axis.
  bool group_rows = false;                   ///< Compact / outline row axis: groups open with their own row.
  std::vector<bool> subtotal_first;          ///< Per depth: the subtotal leads its group.
  std::vector<std::size_t> subtotal_counts;  ///< Per depth: subtotals the evaluator emitted.
};

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
        if (depth == shape_.depth) {
          push(EntryKind::Leaf, df, leaf);
        } else {
          if (shape_.group_rows) {
            push(EntryKind::Group, df);
          }
          level(nodes, depth, df, leaf);
        }
        path_.pop_back();
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

  void node(const AxisHierarchyNode& node, std::size_t depth, std::optional<std::size_t> data_field) {
    path_.push_back(node.label);
    // A group holding the Values level has no single value to show on its
    // own row, so Excel opens it with a label-only row and moves its
    // subtotals below it, one per data field.
    const bool values_below = !data_field && shape_.values_at.has_value() && *shape_.values_at > depth;
    if (shape_.group_rows && values_below) {
      push(EntryKind::Group, std::nullopt);
    }
    if (node.children.empty()) {
      level({}, depth + 1, data_field, leaf_cursor_++);
      path_.pop_back();
      return;
    }
    const std::size_t group_begin = entries_.size();
    level(node.children, depth + 1, data_field, 0);
    const std::size_t subtotal_begin = entries_.size();
    const std::size_t count = depth < shape_.subtotal_counts.size() ? shape_.subtotal_counts[depth] : 0;
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
      }
    }
    const bool first = depth < shape_.subtotal_first.size() && shape_.subtotal_first[depth];
    if (first && !values_below) {
      std::rotate(entries_.begin() + static_cast<std::ptrdiff_t>(group_begin),
                  entries_.begin() + static_cast<std::ptrdiff_t>(subtotal_begin), entries_.end());
    }
    path_.pop_back();
  }

  const PivotTable& table_;
  const AxisShape& shape_;
  std::size_t leaf_cursor_ = 0;
  std::size_t subtotal_cursor_ = 0;
  std::vector<std::string> path_;
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

}  // namespace

Expected<PivotCells, Error> layout(const PivotTable& table, const PivotResult& result,
                                   const PivotLayoutOptions& options) {
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
  const std::string subtotal_suffix = options.subtotal_suffix.empty() && options.subtotal_prefix.empty()
                                          ? std::string(" ") + options.grand_total_label
                                          : options.subtotal_suffix;
  const auto subtotal_label = [&](const std::string& group) {
    return options.subtotal_prefix + group + subtotal_suffix;
  };
  const auto data_total_label = [&](std::size_t df) {
    return options.data_total_prefix + data_field_name(table, df) + options.data_total_suffix;
  };

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
  AxisShape col_shape;
  col_shape.depth = col_depth;
  if (values_on_cols) {
    col_shape.values_at = std::min(table.col_values_position().value_or(col_depth), col_depth);
  }
  col_shape.subtotal_counts = subtotal_counts_for_axis(table, table.col_field_order());
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
  const std::vector<AxisEntry> row_entries = row_entries_or.take();
  auto col_entries_or = collect_axis_entries(table, result.cols, result.col_subtotals, col_shape, "column");
  if (!col_entries_or) {
    return std::move(col_entries_or.error());
  }
  const std::vector<AxisEntry> col_entries = col_entries_or.take();
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
  const std::size_t header_rows =
      legacy ? std::max<std::size_t>(col_levels, 1) + 1 : (col_depth == 0 ? std::size_t{1} : col_levels + 1);
  // Compact form folds the grand-totals strip in axes that have no
  // hierarchy of their own: a no-column-fields pivot's per-row total
  // already lives in its data columns, and likewise a no-row-fields
  // pivot's per-column total. Tabular / Outline share that rule; the
  // English projection keeps emitting both strips.
  const bool emit_grand_totals_rows_strip = table.grand_totals_rows() && (legacy || col_depth > 0);
  const bool emit_grand_totals_cols_strip = table.grand_totals_cols() && (legacy || row_depth > 0);
  // The Values axis totals each data field on its own: one strip column
  // (or row) per field, named after it.
  const std::size_t total_strip_cols = emit_grand_totals_rows_strip ? (values_on_cols ? data_field_count : 1) : 0;
  const std::size_t total_strip_rows = emit_grand_totals_cols_strip ? (values_on_rows ? data_field_count : 1) : 0;
  // Page (report filter) fields are drawn above the report proper: one row
  // per field, then a blank row separating the block from the row/column
  // headers. Excel counts both inside the pivot's own `ref`, so they take
  // the top of the projected grid and the report body starts below them.
  // The block is at least `kPageHeaderCols` wide even over a report
  // narrower than that, because the selection cell has to land somewhere.
  const std::size_t page_rows = result.page_selections.empty() ? 0 : result.page_selections.size() + 1;
  const std::size_t body_rows = header_rows + row_entries.size() + total_strip_rows;
  const std::size_t body_cols = row_header_cols + data_cols + total_strip_cols;
  const std::size_t total_rows = page_rows + body_rows;
  const std::size_t total_cols = (page_rows > 0 && body_cols < kPageHeaderCols) ? kPageHeaderCols : body_cols;

  if (total_rows > UINT32_MAX || total_cols > UINT32_MAX) {
    return make_error(FormulonErrorCode::kEvalPivotInvalid, "pivot layout: projected bounds exceed uint32_t",
                      "rows=" + std::to_string(total_rows) + " cols=" + std::to_string(total_cols));
  }
  auto cell_count_or = checked_mul_size_t(total_rows, total_cols);
  if (!cell_count_or) {
    return std::move(cell_count_or.error());
  }

  PivotCells cells;
  cells.top = table.anchor_row();
  cells.left = table.anchor_col();
  cells.rows = static_cast<std::uint32_t>(total_rows);
  cells.cols = static_cast<std::uint32_t>(total_cols);
  cells.cells.reserve(cell_count_or.value());

  const std::uint32_t left = cells.left;
  const std::uint32_t top = cells.top + static_cast<std::uint32_t>(page_rows);
  const std::uint32_t data_top = top + static_cast<std::uint32_t>(header_rows);
  const std::uint32_t data_left = left + static_cast<std::uint32_t>(row_header_cols);
  const std::uint32_t total_left = data_left + static_cast<std::uint32_t>(data_cols);
  const std::uint32_t row_header_row = top + static_cast<std::uint32_t>(header_rows - 1);
  // Where the grand-totals-rows strip's own header sits: beside the
  // outermost column level (the row "Q1 集計" occupies) once a column
  // hierarchy exists, else beside the row-field names.
  const std::uint32_t total_header_row = col_depth > 0 ? top + 1 : row_header_row;

  // The page block: field name, then the item it is showing. Both are
  // header furniture, so neither earns a `PivotCellKind` of its own — the
  // vocabulary is mirrored by the C ABI and the WASM enum, and a page cell
  // is not something a renderer has to draw differently from a header.
  // Trailing columns and the separator row are explicit blanks so the
  // rendered extent is the rectangle Excel reports.
  for (std::size_t i = 0; i < result.page_selections.size(); ++i) {
    const PivotPageSelection& page = result.page_selections[i];
    const std::uint32_t row = cells.top + static_cast<std::uint32_t>(i);
    append_cell(cells, row, left, text_value(cells, page.field_label), PivotCellKind::Header, 0, page.field_label);
    append_cell(cells, row, left + 1, text_value(cells, page.item_label), PivotCellKind::Header, 0, page.field_label);
    for (std::size_t col = kPageHeaderCols; col < total_cols; ++col) {
      append_cell(cells, row, left + static_cast<std::uint32_t>(col), Value::blank(), PivotCellKind::Blank, 0);
    }
  }
  if (page_rows > 0) {
    const std::uint32_t row = cells.top + static_cast<std::uint32_t>(result.page_selections.size());
    for (std::size_t col = 0; col < total_cols; ++col) {
      append_cell(cells, row, left + static_cast<std::uint32_t>(col), Value::blank(), PivotCellKind::Blank, 0);
    }
  }

  // Display name of the row / column level at `level`: a field's display
  // name, or the stored Values caption for the Values level.
  auto level_name = [&](const std::vector<std::uint32_t>& order, const std::optional<std::size_t>& values_at,
                        std::size_t level) {
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
  };
  auto row_level_name = [&](std::size_t level) {
    return level < row_levels ? level_name(table.row_field_order(), row_shape.values_at, level) : std::string();
  };
  auto col_level_name = [&](std::size_t level) {
    return level < col_levels ? level_name(table.col_field_order(), col_shape.values_at, level) : std::string();
  };
  auto emit_text_or_blank = [&](std::uint32_t row, std::uint32_t col, const std::string& text, PivotCellKind kind,
                                std::uint32_t depth, std::string field_name = {}, std::string number_format = {}) {
    if (text.empty()) {
      append_cell(cells, row, col, Value::blank(), PivotCellKind::Blank, depth, std::move(field_name));
    } else {
      append_cell(cells, row, col, text_value(cells, text), kind, depth, std::move(field_name),
                  std::move(number_format));
    }
  };
  // The data field a cell at (`row_entry`, `col_entry`) reports: whichever
  // axis carries the Values level decides, and a single field is field 0.
  auto cell_data_field = [](const AxisEntry& row_entry, const AxisEntry& col_entry) -> std::size_t {
    if (row_entry.data_field) {
      return *row_entry.data_field;
    }
    return col_entry.data_field.value_or(0);
  };
  // Label of a column with no column level above it: its data field's name,
  // or nothing when the data fields sit on the rows.
  auto flat_col_header = [&](std::size_t c_entry) {
    const AxisEntry& entry = col_entries[c_entry];
    if (entry.data_field) {
      return data_field_name(table, *entry.data_field);
    }
    return data_field_count == 1 ? data_field_name(table, 0) : std::string();
  };

  // Emits one column-level label row at `depth`. Excel shows a non-leaf
  // level's label once per contiguous run and blanks the repeats; the leaf-
  // most level always shows its own value. A subtotal column carries its
  // label on its own group's level and blanks below it. The grand-totals-
  // rows strip is left blank except on `total_header_row`.
  auto emit_col_hierarchy_row = [&](std::uint32_t row, std::size_t depth, bool blank_repeats) {
    const bool leaf_depth = depth + 1 >= col_levels;
    for (std::size_t c_entry = 0; c_entry < col_entries.size(); ++c_entry) {
      const AxisEntry& entry = col_entries[c_entry];
      const bool subtotal_here = entry.kind == EntryKind::Subtotal && depth + 1 == entry.labels.size();
      std::string label;
      if (depth < entry.labels.size()) {
        label = entry.labels[depth];
      }
      if (subtotal_here) {
        label = entry.subtotal_label.empty() ? subtotal_label(label) : entry.subtotal_label;
      } else if (blank_repeats && !leaf_depth && c_entry > 0) {
        const std::vector<std::string>& prev = col_entries[c_entry - 1].labels;
        if (prev.size() > depth && entry.labels.size() > depth &&
            std::equal(prev.begin(), prev.begin() + static_cast<std::ptrdiff_t>(depth + 1), entry.labels.begin())) {
          label.clear();
        }
      }
      const std::uint32_t col = data_left + static_cast<std::uint32_t>(c_entry);
      if (!subtotal_here && col_shape.values_at == depth && entry.data_field) {
        // The Values level heads its column with the data field's name.
        emit_text_or_blank(row, col, label, PivotCellKind::Header, static_cast<std::uint32_t>(depth),
                           data_field_name(table, *entry.data_field), data_field_format(table, *entry.data_field));
        continue;
      }
      emit_text_or_blank(row, col, label,
                         entry.kind == EntryKind::Subtotal ? PivotCellKind::ColSubtotal : PivotCellKind::ColLabel,
                         static_cast<std::uint32_t>(depth), col_level_name(depth));
    }
    if (row != total_header_row) {
      for (std::size_t s = 0; s < total_strip_cols; ++s) {
        append_cell(cells, row, total_left + static_cast<std::uint32_t>(s), Value::blank(), PivotCellKind::Blank, 0);
      }
    }
  };
  // The corner of a pivot with column fields names its data field; with
  // several, the Values level names them instead and the corner is blank.
  auto emit_corner = [&]() {
    if (data_field_count == 1 && row_depth > 0) {
      append_cell(cells, top, left, text_value(cells, data_field_name(table, 0)), PivotCellKind::Header, 0,
                  data_field_name(table, 0), data_field_format(table, 0));
    } else {
      append_cell(cells, top, left, Value::blank(), PivotCellKind::Blank, 0);
    }
  };
  const std::size_t corner_extent = data_cols + total_strip_cols;

  if (multi_col_layout) {
    if (col_depth == 0) {
      // One header row: the row level names, then the data field names.
      for (std::size_t d = 0; d < row_header_cols; ++d) {
        emit_text_or_blank(top, left + static_cast<std::uint32_t>(d), row_level_name(d), PivotCellKind::Header,
                           static_cast<std::uint32_t>(d));
      }
      for (std::size_t c_entry = 0; c_entry < data_cols; ++c_entry) {
        const std::size_t df = col_entries[c_entry].data_field.value_or(0);
        emit_text_or_blank(top, data_left + static_cast<std::uint32_t>(c_entry), flat_col_header(c_entry),
                           PivotCellKind::Header, 0, data_field_name(table, df), data_field_format(table, df));
      }
    } else {
      // Corner row: the corner, then each column level's own name -- the
      // Values level under the stored caption -- where Compact draws its
      // single "Column Labels" placeholder.
      emit_corner();
      for (std::size_t d = 1; d < row_header_cols; ++d) {
        append_cell(cells, top, left + static_cast<std::uint32_t>(d), Value::blank(), PivotCellKind::Blank, 0);
      }
      for (std::size_t i = 0; i < corner_extent; ++i) {
        emit_text_or_blank(top, data_left + static_cast<std::uint32_t>(i), col_level_name(i), PivotCellKind::Header, 0);
      }
      // The last column-level row also carries the row level names.
      for (std::size_t depth = 0; depth < col_levels; ++depth) {
        const std::uint32_t row = top + 1 + static_cast<std::uint32_t>(depth);
        for (std::size_t d = 0; d < row_header_cols; ++d) {
          emit_text_or_blank(row, left + static_cast<std::uint32_t>(d),
                             depth + 1 == col_levels ? row_level_name(d) : std::string(), PivotCellKind::Header,
                             static_cast<std::uint32_t>(d));
        }
        emit_col_hierarchy_row(row, depth, true);
      }
    }
  } else if (compact) {
    if (col_depth == 0) {
      // One header row: the "Row Labels" placeholder, then the data field
      // names (blank when the data fields sit on the rows).
      if (row_header_cols > 0) {
        append_cell(cells, top, left, text_value(cells, options.row_labels_label), PivotCellKind::Header, 0);
      }
      for (std::size_t c_entry = 0; c_entry < data_cols; ++c_entry) {
        const std::size_t df = col_entries[c_entry].data_field.value_or(0);
        emit_text_or_blank(top, data_left + static_cast<std::uint32_t>(c_entry), flat_col_header(c_entry),
                           PivotCellKind::Header, 0, data_field_name(table, df), data_field_format(table, df));
      }
    } else {
      // Corner row: the corner, then the "Column Labels" placeholder over
      // the first column. Without row fields the single data field's name
      // moves down to the data row instead (see the row loop below).
      if (row_header_cols > 0) {
        emit_corner();
      }
      for (std::size_t i = 0; i < corner_extent; ++i) {
        emit_text_or_blank(top, data_left + static_cast<std::uint32_t>(i),
                           i == 0 ? options.column_labels_label : std::string(), PivotCellKind::Header, 0);
      }
      for (std::size_t depth = 0; depth < col_levels; ++depth) {
        const std::uint32_t row = top + 1 + static_cast<std::uint32_t>(depth);
        if (row_header_cols > 0 && (row_levels == 0 || depth + 1 < col_levels)) {
          append_cell(cells, row, left, Value::blank(), PivotCellKind::Blank, 0);
        }
        emit_col_hierarchy_row(row, depth, true);
      }
      if (row_header_cols > 0 && row_levels > 0) {
        append_cell(cells, row_header_row, left, text_value(cells, options.row_labels_label), PivotCellKind::Header, 0);
      }
    }
  } else {
    // English projection: column levels stack from the top, each label
    // repeated per column, over a row of row level names.
    for (std::size_t depth = 0; depth < row_header_cols; ++depth) {
      append_cell(cells, row_header_row, left + static_cast<std::uint32_t>(depth),
                  text_value(cells, row_level_name(depth)), PivotCellKind::Header, static_cast<std::uint32_t>(depth));
    }
    if (col_levels == 0) {
      for (std::size_t c_entry = 0; c_entry < data_cols; ++c_entry) {
        emit_text_or_blank(top, data_left + static_cast<std::uint32_t>(c_entry), flat_col_header(c_entry),
                           PivotCellKind::ColLabel, 0);
      }
    }
    for (std::size_t depth = 0; depth < col_levels; ++depth) {
      emit_col_hierarchy_row(top + static_cast<std::uint32_t>(depth), depth, false);
    }
  }

  // Row labels and data cells. Tabular tracks the previous leaf row's label
  // per column so a parent shared with the row above is blanked.
  std::vector<std::string> prev_row_labels(row_header_cols);
  std::vector<bool> prev_row_labels_filled(row_header_cols, false);
  for (std::size_t r_entry = 0; r_entry < row_entries.size(); ++r_entry) {
    const AxisEntry& entry = row_entries[r_entry];
    const std::uint32_t row = data_top + static_cast<std::uint32_t>(r_entry);
    const std::vector<std::string>& labels = entry.labels;
    const bool subtotal = entry.kind == EntryKind::Subtotal;
    const std::size_t last_depth = labels.empty() ? 0 : labels.size() - 1;
    for (std::size_t depth = 0; depth < row_header_cols; ++depth) {
      std::string label;
      if (multi_col_layout) {
        if (subtotal || entry.kind == EntryKind::Group || outline) {
          // Subtotal and group rows, and every outline row, label only
          // their own level. A tabular subtotal takes the localized
          // wording; outline's doubles as the group's header row.
          if (depth == last_depth && depth < labels.size()) {
            label = labels[depth];
            if (subtotal && !entry.subtotal_label.empty()) {
              label = entry.subtotal_label;
            } else if (subtotal && tabular) {
              label = subtotal_label(label);
            }
          }
        } else if (depth < labels.size()) {
          // Tabular leaf rows repeat the path but blank a parent equal to
          // the previous row's; the leaf-most level always shows.
          const std::string& candidate = labels[depth];
          if (depth + 1 == labels.size() || !prev_row_labels_filled[depth] || prev_row_labels[depth] != candidate) {
            label = candidate;
          }
        }
      } else if (compact) {
        // Compact form shows each row's own level in the single label
        // column; a subtotal there is the bare group label, since it heads
        // its group, unless it totals one data field of the group.
        if (subtotal && !entry.subtotal_label.empty()) {
          label = entry.subtotal_label;
        } else if (!labels.empty()) {
          label = labels.back();
        } else if (row_levels == 0 && col_depth > 0 && r_entry == 0) {
          // No-row-fields pivot: the data field's name heads the single
          // data row.
          label = data_field_name(table, 0);
        }
      } else if (depth < labels.size()) {
        label = labels[depth];
        if (subtotal && depth == last_depth) {
          label = entry.subtotal_label.empty() ? subtotal_label(label) : entry.subtotal_label;
        }
      }
      const PivotCellKind kind = subtotal ? PivotCellKind::RowSubtotal : PivotCellKind::RowLabel;
      if (label.empty()) {
        append_cell(cells, row, left + static_cast<std::uint32_t>(depth), Value::blank(), kind,
                    static_cast<std::uint32_t>(depth), row_level_name(depth));
      } else {
        append_cell(cells, row, left + static_cast<std::uint32_t>(depth), text_value(cells, label), kind,
                    static_cast<std::uint32_t>(depth), row_level_name(depth));
      }
      if (tabular && entry.kind == EntryKind::Leaf && depth < labels.size()) {
        prev_row_labels[depth] = labels[depth];
        prev_row_labels_filled[depth] = true;
      }
    }
    if (entry.kind == EntryKind::Group) {
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
      if (value != nullptr) {
        append_cell(cells, row, col, reify_value(cells, *value), kind, 0, data_field_name(table, df),
                    data_field_format(table, df));
      }
    }
  }

  // Row totals: the strip at the right, from the evaluator's re-aggregated
  // per-row totals so the projection remains cache-only.
  if (emit_grand_totals_rows_strip) {
    for (std::size_t r_entry = 0; r_entry < row_entries.size(); ++r_entry) {
      const AxisEntry& entry = row_entries[r_entry];
      if (entry.kind == EntryKind::Group) {
        continue;
      }
      const std::uint32_t row = data_top + static_cast<std::uint32_t>(r_entry);
      for (std::size_t s = 0; s < total_strip_cols; ++s) {
        const std::size_t df = entry.data_field.value_or(values_on_cols ? s : 0);
        const std::uint32_t col = total_left + static_cast<std::uint32_t>(s);
        if (entry.kind == EntryKind::Subtotal) {
          // `RowSubtotal::values[df]` aggregates the group across the whole
          // column axis, which is exactly what this strip states.
          const RowSubtotal& subtotal = result.row_subtotals[entry.subtotal_index];
          const Value total = df < subtotal.values.size() ? reify_value(cells, subtotal.values[df]) : Value::blank();
          append_cell(cells, row, col, total, PivotCellKind::RowSubtotal, 0, data_field_name(table, df),
                      data_field_format(table, df));
          continue;
        }
        // Use the evaluator's per-row-leaf re-aggregation rather than
        // summing this row's cells: a non-additive function (Average/Max/
        // Min/StdDev/Var) cannot be recovered from the cell aggregates.
        const std::size_t leaf = entry.leaf_index;
        if (leaf < result.row_leaf_totals.size() && df < result.row_leaf_totals[leaf].size() &&
            !result.row_leaf_totals[leaf][df].is_blank()) {
          append_cell(cells, row, col, reify_value(cells, result.row_leaf_totals[leaf][df]), PivotCellKind::GrandTotal,
                      0, data_field_name(table, df), data_field_format(table, df));
          continue;
        }
        // Every row of this strip states a total: a source error wins, an
        // all-numeric row sums, and an aggregate that cannot be summed
        // reports blank rather than leaving a hole in the total column.
        double sum = 0.0;
        bool summable = true;
        const Value* source_error = nullptr;
        for (std::size_t c_leaf = 0; c_leaf < col_leaves.size(); ++c_leaf) {
          const Value& v = result.values[leaf][c_leaf][df];
          if (v.is_error()) {
            source_error = &v;
            break;
          }
          if (v.is_number()) {
            sum += v.as_number();
          } else if (!v.is_blank()) {
            summable = false;
          }
        }
        Value total = Value::blank();
        if (source_error != nullptr) {
          total = reify_value(cells, *source_error);
        } else if (summable) {
          total = Value::number(sum);
        }
        append_cell(cells, row, col, total, PivotCellKind::GrandTotal, 0, data_field_name(table, df),
                    data_field_format(table, df));
      }
    }
    for (std::size_t s = 0; s < total_strip_cols; ++s) {
      const std::size_t df = values_on_cols ? s : 0;
      append_cell(cells, total_header_row, total_left + static_cast<std::uint32_t>(s),
                  text_value(cells, values_on_cols ? data_total_label(s) : options.grand_total_label),
                  PivotCellKind::Header, 0, data_field_name(table, df), data_field_format(table, df));
    }
  }

  // Column totals: the strip at the bottom.
  for (std::size_t b = 0; b < total_strip_rows; ++b) {
    const std::optional<std::size_t> row_df = values_on_rows ? std::optional<std::size_t>(b) : std::nullopt;
    const std::uint32_t total_row = data_top + static_cast<std::uint32_t>(row_entries.size() + b);
    append_cell(cells, total_row, left, text_value(cells, row_df ? data_total_label(b) : options.grand_total_label),
                PivotCellKind::GrandTotal, 0);
    // Tabular / Outline leave the remaining row-label columns blank so the
    // rendered grid is rectangular.
    if (multi_col_layout) {
      for (std::size_t d = 1; d < row_header_cols; ++d) {
        append_cell(cells, total_row, left + static_cast<std::uint32_t>(d), Value::blank(), PivotCellKind::GrandTotal,
                    static_cast<std::uint32_t>(d));
      }
    }
    for (std::size_t c_entry = 0; c_entry < col_entries.size(); ++c_entry) {
      const AxisEntry& col_entry = col_entries[c_entry];
      const std::size_t df = row_df.value_or(col_entry.data_field.value_or(0));
      const std::uint32_t col = data_left + static_cast<std::uint32_t>(c_entry);
      const bool col_subtotal = col_entry.kind == EntryKind::Subtotal;
      // Prefer the evaluator's re-aggregated totals: summing per-row
      // aggregates is wrong for Average, Max, Min, StdDev and Var.
      if (col_depth == 0 && !col_subtotal && df < result.grand_totals.size() && !result.grand_totals[df].is_blank()) {
        append_cell(cells, total_row, col, reify_value(cells, result.grand_totals[df]), PivotCellKind::GrandTotal, 0,
                    data_field_name(table, df), data_field_format(table, df));
        continue;
      }
      if (!col_subtotal && col_entry.leaf_index < result.col_leaf_totals.size() &&
          df < result.col_leaf_totals[col_entry.leaf_index].size() &&
          !result.col_leaf_totals[col_entry.leaf_index][df].is_blank()) {
        append_cell(cells, total_row, col, reify_value(cells, result.col_leaf_totals[col_entry.leaf_index][df]),
                    PivotCellKind::GrandTotal, 0, data_field_name(table, df), data_field_format(table, df));
        continue;
      }
      const PivotCellKind kind = col_subtotal ? PivotCellKind::ColSubtotal : PivotCellKind::GrandTotal;
      double sum = 0.0;
      bool numeric = true;
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
          append_cell(cells, total_row, col, reify_value(cells, *value), kind, 0, data_field_name(table, df),
                      data_field_format(table, df));
          numeric = false;
          break;
        }
        if (value->is_number()) {
          sum += value->as_number();
        } else if (!value->is_blank()) {
          numeric = false;
        }
      }
      if (numeric) {
        append_cell(cells, total_row, col, Value::number(sum), kind, 0, data_field_name(table, df),
                    data_field_format(table, df));
      }
    }
    for (std::size_t s = 0; s < total_strip_cols; ++s) {
      const std::size_t df = row_df.value_or(values_on_cols ? s : 0);
      Value v = Value::blank();
      if (df < result.grand_totals.size() && !result.grand_totals[df].is_blank()) {
        v = reify_value(cells, result.grand_totals[df]);
      } else if (df == 0 && !result.grand_total.is_blank()) {
        v = reify_value(cells, result.grand_total);
      } else {
        double sum = 0.0;
        bool numeric = true;
        for (std::size_t r_leaf = 0; r_leaf < row_leaves.size(); ++r_leaf) {
          for (std::size_t c_leaf = 0; c_leaf < col_leaves.size(); ++c_leaf) {
            const Value& cell = result.values[r_leaf][c_leaf][df];
            if (cell.is_number()) {
              sum += cell.as_number();
            } else if (!cell.is_blank()) {
              numeric = false;
            }
          }
        }
        if (numeric) {
          v = Value::number(sum);
        }
      }
      append_cell(cells, total_row, total_left + static_cast<std::uint32_t>(s), v, PivotCellKind::GrandTotal, 0,
                  data_field_name(table, df), data_field_format(table, df));
    }
  }

  return cells;
}

}  // namespace formulon::pivot
