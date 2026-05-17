// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the PivotResult -> grid-cell projection.

#include "pivot/pivot_layout.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

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

struct RowEntry {
  AxisLeaf leaf;
  bool subtotal = false;
  std::size_t subtotal_index = 0;
};

struct ColEntry {
  AxisLeaf leaf;
  bool subtotal = false;
  std::size_t subtotal_index = 0;
};

std::string field_display_name(const PivotField& field) {
  if (!field.custom_name.empty()) {
    return field.custom_name;
  }
  return field.source_name;
}

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

void collect_row_leaves_impl(const RowHierarchyNode& node, std::vector<std::string>& path,
                             std::vector<AxisLeaf>& leaves) {
  path.push_back(node.label);
  if (node.children.empty()) {
    leaves.push_back({path});
  } else {
    for (const RowHierarchyNode& child : node.children) {
      collect_row_leaves_impl(child, path, leaves);
    }
  }
  path.pop_back();
}

std::vector<AxisLeaf> collect_row_leaves(const std::vector<RowHierarchyNode>& roots, std::size_t depth) {
  std::vector<AxisLeaf> leaves;
  if (depth == 0) {
    leaves.push_back(AxisLeaf{});
    return leaves;
  }
  std::vector<std::string> path;
  for (const RowHierarchyNode& root : roots) {
    collect_row_leaves_impl(root, path, leaves);
  }
  return leaves;
}

bool labels_equal(const std::vector<std::string>& a, const std::vector<std::string>& b) {
  return a.size() == b.size() && std::equal(a.begin(), a.end(), b.begin());
}

void collect_row_entries_impl(const RowHierarchyNode& node, std::vector<std::string>& path,
                              const std::vector<RowSubtotal>& subtotals, std::size_t& subtotal_cursor,
                              std::vector<RowEntry>& rows, bool subtotal_first) {
  path.push_back(node.label);
  if (node.children.empty()) {
    rows.push_back({AxisLeaf{path}, false, 0});
  } else {
    // Capture the subtotal slot up-front so `subtotal_first` mode can
    // emit it ahead of the children while keeping the cursor in
    // monotonic lock-step with the evaluator-produced sequence.
    const bool has_subtotal =
        subtotal_cursor < subtotals.size() && labels_equal(subtotals[subtotal_cursor].labels, path);
    const std::size_t subtotal_idx = has_subtotal ? subtotal_cursor : 0;
    if (has_subtotal && subtotal_first) {
      rows.push_back({AxisLeaf{path}, true, subtotal_idx});
      ++subtotal_cursor;
    }
    for (const RowHierarchyNode& child : node.children) {
      collect_row_entries_impl(child, path, subtotals, subtotal_cursor, rows, subtotal_first);
    }
    if (has_subtotal && !subtotal_first) {
      rows.push_back({AxisLeaf{path}, true, subtotal_idx});
      ++subtotal_cursor;
    }
  }
  path.pop_back();
}

std::vector<RowEntry> collect_row_entries(const PivotResult& result, std::size_t depth, bool include_subtotals,
                                          bool subtotal_first) {
  if (depth == 0) {
    return {RowEntry{}};
  }
  if (!include_subtotals) {
    std::vector<RowEntry> entries;
    for (AxisLeaf& leaf : collect_row_leaves(result.rows, depth)) {
      entries.push_back({std::move(leaf), false, 0});
    }
    return entries;
  }
  std::vector<RowEntry> entries;
  std::vector<std::string> path;
  std::size_t subtotal_cursor = 0;
  for (const RowHierarchyNode& root : result.rows) {
    collect_row_entries_impl(root, path, result.row_subtotals, subtotal_cursor, entries, subtotal_first);
  }
  return entries;
}

void collect_col_leaves_impl(const ColHierarchyNode& node, std::vector<std::string>& path,
                             std::vector<AxisLeaf>& leaves) {
  path.push_back(node.label);
  if (node.children.empty()) {
    leaves.push_back({path});
  } else {
    for (const ColHierarchyNode& child : node.children) {
      collect_col_leaves_impl(child, path, leaves);
    }
  }
  path.pop_back();
}

std::vector<AxisLeaf> collect_col_leaves(const std::vector<ColHierarchyNode>& roots, std::size_t depth) {
  std::vector<AxisLeaf> leaves;
  if (depth == 0) {
    leaves.push_back(AxisLeaf{});
    return leaves;
  }
  std::vector<std::string> path;
  for (const ColHierarchyNode& root : roots) {
    collect_col_leaves_impl(root, path, leaves);
  }
  return leaves;
}

void collect_col_entries_impl(const ColHierarchyNode& node, std::vector<std::string>& path,
                              const std::vector<ColSubtotal>& subtotals, std::size_t& subtotal_cursor,
                              std::vector<ColEntry>& cols) {
  path.push_back(node.label);
  if (node.children.empty()) {
    cols.push_back({AxisLeaf{path}, false, 0});
  } else {
    for (const ColHierarchyNode& child : node.children) {
      collect_col_entries_impl(child, path, subtotals, subtotal_cursor, cols);
    }
    if (subtotal_cursor < subtotals.size() && labels_equal(subtotals[subtotal_cursor].labels, path)) {
      cols.push_back({AxisLeaf{path}, true, subtotal_cursor});
      ++subtotal_cursor;
    }
  }
  path.pop_back();
}

std::vector<ColEntry> collect_col_entries(const PivotResult& result, std::size_t depth, bool include_subtotals) {
  if (depth == 0) {
    return {ColEntry{}};
  }
  if (!include_subtotals) {
    std::vector<ColEntry> entries;
    for (AxisLeaf& leaf : collect_col_leaves(result.cols, depth)) {
      entries.push_back({std::move(leaf), false, 0});
    }
    return entries;
  }
  std::vector<ColEntry> entries;
  std::vector<std::string> path;
  std::size_t subtotal_cursor = 0;
  for (const ColHierarchyNode& root : result.cols) {
    collect_col_entries_impl(root, path, result.col_subtotals, subtotal_cursor, entries);
  }
  return entries;
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
  // into Excel-style rendering AND the pivot has row fields with no
  // column hierarchy. Pivots with column fields fall back to the
  // compact (single row-label column) shape that the workbook oracle
  // already verifies for the col-field cases; widening tabular/outline
  // to col-field pivots is a separate scope.
  const bool tabular = locale_opted_in && table.layout() == PivotLayout::Tabular && row_depth > 0 && col_depth == 0;
  const bool outline = locale_opted_in && table.layout() == PivotLayout::Outline && row_depth > 0 && col_depth == 0;
  const bool multi_col_layout = tabular || outline;
  const bool compact = locale_opted_in && !multi_col_layout;
  const std::string subtotal_suffix =
      options.subtotal_suffix.empty() ? std::string(" ") + options.grand_total_label : options.subtotal_suffix;

  std::vector<AxisLeaf> row_leaves = collect_row_leaves(result.rows, row_depth);
  std::vector<AxisLeaf> col_leaves = collect_col_leaves(result.cols, col_depth);
  const bool include_row_subtotals = !result.row_subtotals.empty();
  // Outline places the subtotal row ABOVE the group it summarises (so
  // the parent label / value share one row). Compact also emits the
  // subtotal at the top of the group because Excel's compact form
  // shows the parent label and its sum on the same row. Tabular keeps
  // the legacy English ordering: subtotal BELOW the group.
  const bool subtotal_first = compact || outline;
  std::vector<RowEntry> row_entries = collect_row_entries(result, row_depth, include_row_subtotals, subtotal_first);
  const bool include_col_subtotals = !result.col_subtotals.empty();
  std::vector<ColEntry> col_entries = collect_col_entries(result, col_depth, include_col_subtotals);

  auto valid_or = validate_result_shape(table, result, row_leaves.size(), col_leaves.size());
  if (!valid_or) {
    return valid_or.error();
  }

  auto data_cols_or = checked_mul_size_t(col_entries.size(), data_field_count);
  if (!data_cols_or) {
    return data_cols_or.error();
  }
  const std::size_t data_cols = data_cols_or.value();
  // Compact form merges every row-field level into the same physical
  // column (Excel renders nested keys via indentation, not via
  // adjacent columns), so the row-label side always occupies one
  // column regardless of `row_depth`. Tabular / Outline give each row
  // field its own column so they consume `row_depth` columns.
  const std::size_t row_header_cols =
      (compact || row_depth == 0) ? std::size_t{1} : (multi_col_layout ? row_depth : row_depth);
  const std::size_t col_header_rows = col_depth == 0 ? 1 : col_depth;
  const std::size_t data_field_header_rows = data_field_count > 1 ? 1 : 0;
  // Compact form folds the row-field-name row away: when there are no
  // column fields the placeholder + data-field-name share a single
  // header row, and when column fields exist the row-field placeholder
  // lives on the last column-header row instead of an extra row of its
  // own. Tabular / Outline use one header row carrying the per-column
  // row-field display names + the data-field display name(s).
  // English mode preserves the legacy `col_header_rows +
  // data_field_header_rows + 1` shape.
  const std::size_t header_rows =
      compact ? (col_depth == 0 ? std::size_t{1} : col_header_rows + 1)
              : (multi_col_layout ? std::size_t{1} : col_header_rows + data_field_header_rows + 1);
  // Compact form folds the grand-totals strip in axes that have no
  // hierarchy of their own: a no-column-fields pivot's per-row total
  // already lives in the single data column, so the right-hand grand-
  // totals-rows strip would be redundant. Same for a no-row-fields
  // pivot's per-col total and the bottom grand-totals-cols row. Excel
  // suppresses both strips in that situation; tabular/outline follow
  // the same suppression rule (they share Compact's no-redundant-strip
  // policy). The legacy English layout below keeps emitting them.
  const bool emit_grand_totals_rows_strip =
      table.grand_totals_rows() && (!(compact || multi_col_layout) || col_depth > 0);
  const bool emit_grand_totals_cols_strip =
      table.grand_totals_cols() && (!(compact || multi_col_layout) || row_depth > 0);
  const std::size_t total_rows = header_rows + row_entries.size() + (emit_grand_totals_cols_strip ? 1 : 0);
  const std::size_t total_cols = row_header_cols + data_cols + (emit_grand_totals_rows_strip ? data_field_count : 0);

  if (total_rows > UINT32_MAX || total_cols > UINT32_MAX) {
    return make_error(FormulonErrorCode::kEvalPivotInvalid, "pivot layout: projected bounds exceed uint32_t",
                      "rows=" + std::to_string(total_rows) + " cols=" + std::to_string(total_cols));
  }

  PivotCells cells;
  cells.top = table.anchor_row();
  cells.left = table.anchor_col();
  cells.rows = static_cast<std::uint32_t>(total_rows);
  cells.cols = static_cast<std::uint32_t>(total_cols);
  cells.cells.reserve(total_rows * total_cols);

  const std::uint32_t top = cells.top;
  const std::uint32_t left = cells.left;
  const std::uint32_t data_top = top + static_cast<std::uint32_t>(header_rows);
  const std::uint32_t data_left = left + static_cast<std::uint32_t>(row_header_cols);
  const std::uint32_t row_header_row = top + static_cast<std::uint32_t>(header_rows - 1);

  if (multi_col_layout) {
    // Tabular / Outline header: one row containing the per-row-field
    // display names followed by the data-field display name(s). Only
    // exercised for row-only pivots (col_depth == 0 enforced above).
    for (std::size_t d = 0; d < row_depth; ++d) {
      std::string label;
      if (table.row_field_order()[d] < table.fields().size()) {
        label = field_display_name(table.fields()[table.row_field_order()[d]]);
      }
      append_cell(cells, top, left + static_cast<std::uint32_t>(d), text_value(cells, std::move(label)),
                  PivotCellKind::Header, static_cast<std::uint32_t>(d));
    }
    for (std::size_t df = 0; df < data_field_count; ++df) {
      const std::uint32_t col = data_left + static_cast<std::uint32_t>(df);
      append_cell(cells, top, col, text_value(cells, data_field_name(table, df)), PivotCellKind::Header, 0,
                  data_field_name(table, df), data_field_format(table, df));
    }
  } else if (compact) {
    // Compact form: a single "Row Labels" placeholder occupies the
    // corner instead of the per-row-field display names. With column
    // fields the data-field name slides into the top-left corner and a
    // "Column Labels" placeholder sits over the first column-leaf slot;
    // without column fields the data-field name(s) span the header row
    // directly, so no separate data-field-header row is emitted.
    if (col_depth == 0) {
      // Anchor placeholder.
      append_cell(cells, top, left, text_value(cells, options.row_labels_label), PivotCellKind::Header, 0);
      // Data-field column headers (one per value field).
      for (std::size_t df = 0; df < data_field_count; ++df) {
        const std::uint32_t col = data_left + static_cast<std::uint32_t>(df);
        append_cell(cells, top, col, text_value(cells, data_field_name(table, df)), PivotCellKind::Header, 0,
                    data_field_name(table, df), data_field_format(table, df));
      }
    } else {
      // Corner row: when the pivot has row fields the corner carries
      // the data-field name and the placeholder sits over the first
      // column-leaf slot. Without row fields the corner stays blank
      // and the data-field name slides down to the data row instead
      // (emitted by the row-label loop below).
      // (Multi-data-field combined with column fields is not exercised
      // by the workbook oracle, so we emit only the first data-field
      // name in the corner; sibling slots stay blank.)
      if (row_depth > 0) {
        append_cell(cells, top, left, text_value(cells, data_field_name(table, 0)), PivotCellKind::Header, 0,
                    data_field_name(table, 0), data_field_format(table, 0));
      } else {
        // Without row fields the corner stays blank, but Excel still
        // emits a placeholder cell at (top, left) so consumers can
        // round-trip the rendered grid without losing the anchor.
        append_cell(cells, top, left, Value::blank(), PivotCellKind::Blank, 0);
      }
      if (!options.column_labels_label.empty()) {
        append_cell(cells, top, data_left, text_value(cells, options.column_labels_label), PivotCellKind::Header, 0);
      }
      // Remaining corner slots are explicit blanks so the rendered
      // grid lines up with Excel's reported extent.
      const std::size_t corner_extent = data_cols + (emit_grand_totals_rows_strip ? data_field_count : 0);
      for (std::size_t i = 1; i < corner_extent; ++i) {
        append_cell(cells, top, data_left + static_cast<std::uint32_t>(i), Value::blank(), PivotCellKind::Blank, 0);
      }
      // Column hierarchy labels start one row below the corner.
      for (std::size_t depth = 0; depth < col_header_rows; ++depth) {
        const std::uint32_t row = top + 1 + static_cast<std::uint32_t>(depth);
        if (row_depth == 0) {
          // Row-label column is otherwise empty in this row; emit an
          // explicit blank so the rendered extent matches Excel.
          append_cell(cells, row, left, Value::blank(), PivotCellKind::Blank, 0);
        }
        for (std::size_t c_entry = 0; c_entry < col_entries.size(); ++c_entry) {
          const ColEntry& entry = col_entries[c_entry];
          const AxisLeaf& leaf = entry.leaf;
          std::string label;
          if (depth < leaf.labels.size()) {
            label = leaf.labels[depth];
          }
          if (entry.subtotal && depth + 1 == leaf.labels.size()) {
            label += subtotal_suffix;
          }
          std::string field_name;
          if (depth < col_depth && table.col_field_order()[depth] < table.fields().size()) {
            field_name = field_display_name(table.fields()[table.col_field_order()[depth]]);
          }
          for (std::size_t df = 0; df < data_field_count; ++df) {
            const std::uint32_t col = data_left + static_cast<std::uint32_t>(c_entry * data_field_count + df);
            append_cell(cells, row, col, text_value(cells, label),
                        entry.subtotal ? PivotCellKind::ColSubtotal : PivotCellKind::ColLabel,
                        static_cast<std::uint32_t>(depth), field_name);
          }
        }
      }
      // Row labels placeholder lives on the last header row when there
      // are row fields. Without row fields the data-field name will go
      // there instead, handled by the row-label loop below.
      if (row_depth > 0) {
        append_cell(cells, row_header_row, left, text_value(cells, options.row_labels_label), PivotCellKind::Header, 0);
      }
    }
  } else {
    // Legacy English layout: row-field display names occupy the last
    // header row, column hierarchy labels span `col_header_rows`, and a
    // dedicated data-field-header row appears whenever the pivot
    // carries more than one value field.

    // Row field headers.
    for (std::size_t depth = 0; depth < row_header_cols; ++depth) {
      std::string label;
      if (depth < row_depth && table.row_field_order()[depth] < table.fields().size()) {
        label = field_display_name(table.fields()[table.row_field_order()[depth]]);
      }
      append_cell(cells, row_header_row, left + static_cast<std::uint32_t>(depth), text_value(cells, std::move(label)),
                  PivotCellKind::Header, static_cast<std::uint32_t>(depth));
    }

    // Column hierarchy labels. Labels are repeated per leaf/data-field
    // slot; consumers can visually merge adjacent equal labels if
    // desired.
    for (std::size_t depth = 0; depth < col_header_rows; ++depth) {
      const std::uint32_t row = top + static_cast<std::uint32_t>(depth);
      for (std::size_t c_entry = 0; c_entry < col_entries.size(); ++c_entry) {
        const ColEntry& entry = col_entries[c_entry];
        const AxisLeaf& leaf = entry.leaf;
        std::string label;
        if (depth < leaf.labels.size()) {
          label = leaf.labels[depth];
        }
        if (entry.subtotal && depth + 1 == leaf.labels.size()) {
          label += subtotal_suffix;
        } else if (col_depth == 0 && data_field_count == 1) {
          label = data_field_name(table, 0);
        } else if (col_depth == 0) {
          label = options.values_label;
        }
        std::string field_name;
        if (depth < col_depth && table.col_field_order()[depth] < table.fields().size()) {
          field_name = field_display_name(table.fields()[table.col_field_order()[depth]]);
        }
        for (std::size_t df = 0; df < data_field_count; ++df) {
          const std::uint32_t col = data_left + static_cast<std::uint32_t>(c_entry * data_field_count + df);
          append_cell(cells, row, col, text_value(cells, label),
                      entry.subtotal ? PivotCellKind::ColSubtotal : PivotCellKind::ColLabel,
                      static_cast<std::uint32_t>(depth), field_name);
        }
      }
    }

    // Data-field header row when the pivot has multiple value fields.
    if (data_field_count > 1) {
      const std::uint32_t row = top + static_cast<std::uint32_t>(col_header_rows);
      for (std::size_t c_entry = 0; c_entry < col_entries.size(); ++c_entry) {
        for (std::size_t df = 0; df < data_field_count; ++df) {
          const std::uint32_t col = data_left + static_cast<std::uint32_t>(c_entry * data_field_count + df);
          append_cell(cells, row, col, text_value(cells, data_field_name(table, df)), PivotCellKind::Header, 0,
                      data_field_name(table, df), data_field_format(table, df));
        }
      }
    }
  }

  // Row labels, row subtotal rows, and data cells.
  // For Tabular / Outline we track the per-depth "last visible label"
  // so a leaf row that shares a parent with the row above it can emit
  // a blank cell in the parent column instead of repeating the parent
  // value. Compact / English layouts ignore this state.
  std::vector<std::string> prev_row_labels(row_header_cols);
  std::vector<bool> prev_row_labels_filled(row_header_cols, false);
  std::size_t row_leaf_index = 0;
  for (std::size_t r_entry = 0; r_entry < row_entries.size(); ++r_entry) {
    const RowEntry& entry = row_entries[r_entry];
    const std::uint32_t row = data_top + static_cast<std::uint32_t>(r_entry);
    const AxisLeaf& leaf = entry.leaf;
    for (std::size_t depth = 0; depth < row_header_cols; ++depth) {
      std::string label;
      bool emit_blank = false;
      if (multi_col_layout) {
        const std::size_t leaf_last_depth = leaf.labels.empty() ? 0 : leaf.labels.size() - 1;
        if (entry.subtotal) {
          // Subtotal row carries the parent-most label in the parent's
          // own depth column; all other row-label columns stay blank.
          // Tabular appends the " 集計" suffix; Outline shows the bare
          // parent label because the same row also doubles as the
          // parent-group header.
          if (depth == leaf_last_depth && depth < leaf.labels.size()) {
            label = leaf.labels[depth];
            if (tabular) {
              label += subtotal_suffix;
            }
          } else {
            emit_blank = true;
          }
        } else if (outline) {
          // Outline leaf rows: only the leaf-most column carries a
          // label; the parent columns stay blank because the parent
          // value already appeared on the subtotal row that sits above
          // this group.
          if (depth == leaf_last_depth && depth < leaf.labels.size()) {
            label = leaf.labels[depth];
          } else {
            emit_blank = true;
          }
        } else {
          // Tabular leaf rows: repeat the leaf's label path across the
          // row-label columns, but blank a column whose value equals
          // the previous row's value in that column (Excel renders the
          // parent label only on its first child row of each group).
          if (depth < leaf.labels.size()) {
            const std::string& candidate = leaf.labels[depth];
            if (depth + 1 == leaf.labels.size()) {
              // Leaf-most column always shows the leaf value.
              label = candidate;
            } else if (prev_row_labels_filled[depth] && prev_row_labels[depth] == candidate) {
              emit_blank = true;
            } else {
              label = candidate;
            }
          } else {
            emit_blank = true;
          }
        }
      } else if (compact && !leaf.labels.empty()) {
        // Compact form collapses every hierarchy level into a single
        // physical column and shows the leaf-most key on each row.
        // Parent subtotal rows carry a shorter label path (only up to
        // the parent) so the back element naturally selects the
        // parent's own name.
        label = leaf.labels.back();
      } else if (depth < leaf.labels.size()) {
        label = leaf.labels[depth];
      }
      if (!multi_col_layout) {
        if (entry.subtotal && depth + 1 == leaf.labels.size() && !compact) {
          // Compact form hides the localized " 集計" / " Grand Total"
          // suffix on row subtotals: the parent-group row already
          // carries the bare group label and Excel relies on the
          // following indented child rows for visual disambiguation.
          label += subtotal_suffix;
        } else if (compact && row_depth == 0 && col_depth > 0 && depth == 0 && r_entry == 0) {
          // No-row-fields compact pivot: the data-field name sits on
          // the row-label column of the single implicit data row, since
          // there is no separate grand-totals row to host it.
          label = data_field_name(table, 0);
        }
      }
      std::string field_name;
      if (depth < row_depth && table.row_field_order()[depth] < table.fields().size()) {
        field_name = field_display_name(table.fields()[table.row_field_order()[depth]]);
      }
      if (emit_blank) {
        append_cell(cells, row, left + static_cast<std::uint32_t>(depth), Value::blank(),
                    entry.subtotal ? PivotCellKind::RowSubtotal : PivotCellKind::RowLabel,
                    static_cast<std::uint32_t>(depth), std::move(field_name));
      } else {
        append_cell(cells, row, left + static_cast<std::uint32_t>(depth), text_value(cells, label),
                    entry.subtotal ? PivotCellKind::RowSubtotal : PivotCellKind::RowLabel,
                    static_cast<std::uint32_t>(depth), std::move(field_name));
      }
      // Track only Tabular's leaf-row label-suppression state. Outline
      // never repeats parent labels, and subtotal rows do not change
      // the "previous parent value" sliding window.
      if (multi_col_layout && tabular && !entry.subtotal) {
        if (depth < leaf.labels.size()) {
          prev_row_labels[depth] = leaf.labels[depth];
          prev_row_labels_filled[depth] = true;
        }
      }
    }
    if (entry.subtotal) {
      const RowSubtotal& subtotal = result.row_subtotals[entry.subtotal_index];
      std::size_t col_leaf_index = 0;
      for (std::size_t c_entry = 0; c_entry < col_entries.size(); ++c_entry) {
        const ColEntry& col_entry = col_entries[c_entry];
        for (std::size_t df = 0; df < data_field_count; ++df) {
          const std::uint32_t col = data_left + static_cast<std::uint32_t>(c_entry * data_field_count + df);
          const Value* value = nullptr;
          if (col_entry.subtotal && col_entry.subtotal_index < subtotal.col_subtotal_values.size() &&
              df < subtotal.col_subtotal_values[col_entry.subtotal_index].size()) {
            value = &subtotal.col_subtotal_values[col_entry.subtotal_index][df];
          } else if (!col_entry.subtotal && col_leaf_index < subtotal.col_values.size() &&
                     df < subtotal.col_values[col_leaf_index].size()) {
            value = &subtotal.col_values[col_leaf_index][df];
          } else if (df < subtotal.values.size()) {
            value = &subtotal.values[df];
          }
          if (value != nullptr) {
            append_cell(cells, row, col, reify_value(cells, *value), PivotCellKind::RowSubtotal, 0,
                        data_field_name(table, df), data_field_format(table, df));
          }
        }
        if (!col_entry.subtotal) {
          ++col_leaf_index;
        }
      }
    } else {
      std::size_t col_leaf_index = 0;
      for (std::size_t c_entry = 0; c_entry < col_entries.size(); ++c_entry) {
        const ColEntry& col_entry = col_entries[c_entry];
        for (std::size_t df = 0; df < data_field_count; ++df) {
          const std::uint32_t col = data_left + static_cast<std::uint32_t>(c_entry * data_field_count + df);
          if (col_entry.subtotal) {
            const ColSubtotal& subtotal = result.col_subtotals[col_entry.subtotal_index];
            if (row_leaf_index < subtotal.values.size() && df < subtotal.values[row_leaf_index].size()) {
              append_cell(cells, row, col, reify_value(cells, subtotal.values[row_leaf_index][df]),
                          PivotCellKind::ColSubtotal, 0, data_field_name(table, df), data_field_format(table, df));
            }
          } else {
            append_cell(cells, row, col, reify_value(cells, result.values[row_leaf_index][col_leaf_index][df]),
                        PivotCellKind::Data, 0, data_field_name(table, df), data_field_format(table, df));
          }
        }
        if (!col_entry.subtotal) {
          ++col_leaf_index;
        }
      }
      ++row_leaf_index;
    }
  }

  // Row totals: one total strip at the right side, computed from the
  // already-evaluated data matrix so the projection remains cache-only.
  if (emit_grand_totals_rows_strip) {
    const std::uint32_t total_left = data_left + static_cast<std::uint32_t>(data_cols);
    std::size_t total_r_leaf = 0;
    for (std::size_t r_entry = 0; r_entry < row_entries.size(); ++r_entry) {
      if (row_entries[r_entry].subtotal) {
        continue;
      }
      const std::uint32_t row = data_top + static_cast<std::uint32_t>(r_entry);
      for (std::size_t df = 0; df < data_field_count; ++df) {
        double sum = 0.0;
        bool numeric = true;
        for (std::size_t c_leaf = 0; c_leaf < col_leaves.size(); ++c_leaf) {
          const Value& v = result.values[total_r_leaf][c_leaf][df];
          if (v.is_error()) {
            append_cell(cells, row, total_left + static_cast<std::uint32_t>(df), reify_value(cells, v),
                        PivotCellKind::GrandTotal, 0, data_field_name(table, df), data_field_format(table, df));
            numeric = false;
            break;
          }
          if (v.is_number()) {
            sum += v.as_number();
          } else if (!v.is_blank()) {
            numeric = false;
          }
        }
        if (numeric) {
          append_cell(cells, row, total_left + static_cast<std::uint32_t>(df), Value::number(sum),
                      PivotCellKind::GrandTotal, 0, data_field_name(table, df), data_field_format(table, df));
        }
      }
      ++total_r_leaf;
    }
    for (std::size_t df = 0; df < data_field_count; ++df) {
      append_cell(cells, row_header_row, total_left + static_cast<std::uint32_t>(df),
                  text_value(cells, options.grand_total_label), PivotCellKind::Header, 0, data_field_name(table, df),
                  data_field_format(table, df));
    }
  }

  // Column totals: one total row at the bottom.
  if (emit_grand_totals_cols_strip) {
    const std::uint32_t total_row = data_top + static_cast<std::uint32_t>(row_entries.size());
    append_cell(cells, total_row, left, text_value(cells, options.grand_total_label), PivotCellKind::GrandTotal, 0);
    // Tabular / Outline give each row field its own column; the grand
    // total row carries the localized label in the first column and
    // leaves the remaining row-label columns blank so the rendered
    // grid is rectangular.
    if (multi_col_layout) {
      for (std::size_t d = 1; d < row_header_cols; ++d) {
        append_cell(cells, total_row, left + static_cast<std::uint32_t>(d), Value::blank(), PivotCellKind::GrandTotal,
                    static_cast<std::uint32_t>(d));
      }
    }
    std::size_t total_col_leaf_index = 0;
    for (std::size_t c_entry = 0; c_entry < col_entries.size(); ++c_entry) {
      const ColEntry& col_entry = col_entries[c_entry];
      for (std::size_t df = 0; df < data_field_count; ++df) {
        const std::uint32_t col = data_left + static_cast<std::uint32_t>(c_entry * data_field_count + df);
        // Prefer the evaluator's pre-aggregated grand total when the
        // pivot has no column hierarchy: summing the per-row data here
        // would be wrong for non-additive aggregations (Average, Max,
        // Min, StdDev, Var) because each row already contains the
        // per-group aggregate, not raw source values. The evaluator
        // re-applies the aggregation over all surviving records and
        // stores the result in `result.grand_totals[df]`.
        if (col_depth == 0 && !col_entry.subtotal && df < result.grand_totals.size() &&
            !result.grand_totals[df].is_blank()) {
          append_cell(cells, total_row, col, reify_value(cells, result.grand_totals[df]), PivotCellKind::GrandTotal, 0,
                      data_field_name(table, df), data_field_format(table, df));
          continue;
        }
        double sum = 0.0;
        bool numeric = true;
        for (std::size_t r_leaf = 0; r_leaf < row_leaves.size(); ++r_leaf) {
          const Value* value = nullptr;
          if (col_entry.subtotal) {
            const ColSubtotal& subtotal = result.col_subtotals[col_entry.subtotal_index];
            if (r_leaf < subtotal.values.size() && df < subtotal.values[r_leaf].size()) {
              value = &subtotal.values[r_leaf][df];
            }
          } else {
            value = &result.values[r_leaf][total_col_leaf_index][df];
          }
          if (value == nullptr) {
            continue;
          }
          const Value& v = *value;
          if (v.is_error()) {
            append_cell(cells, total_row, col, reify_value(cells, v),
                        col_entry.subtotal ? PivotCellKind::ColSubtotal : PivotCellKind::GrandTotal, 0,
                        data_field_name(table, df), data_field_format(table, df));
            numeric = false;
            break;
          }
          if (v.is_number()) {
            sum += v.as_number();
          } else if (!v.is_blank()) {
            numeric = false;
          }
        }
        if (numeric) {
          append_cell(cells, total_row, col, Value::number(sum),
                      col_entry.subtotal ? PivotCellKind::ColSubtotal : PivotCellKind::GrandTotal, 0,
                      data_field_name(table, df), data_field_format(table, df));
        }
      }
      if (!col_entry.subtotal) {
        ++total_col_leaf_index;
      }
    }
    if (emit_grand_totals_rows_strip) {
      const std::uint32_t total_left = data_left + static_cast<std::uint32_t>(data_cols);
      for (std::size_t df = 0; df < data_field_count; ++df) {
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
        append_cell(cells, total_row, total_left + static_cast<std::uint32_t>(df), v, PivotCellKind::GrandTotal, 0,
                    data_field_name(table, df), data_field_format(table, df));
      }
    }
  }

  return cells;
}

}  // namespace formulon::pivot
