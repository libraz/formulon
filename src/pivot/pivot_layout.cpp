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
                              std::vector<RowEntry>& rows) {
  path.push_back(node.label);
  if (node.children.empty()) {
    rows.push_back({AxisLeaf{path}, false, 0});
  } else {
    for (const RowHierarchyNode& child : node.children) {
      collect_row_entries_impl(child, path, subtotals, subtotal_cursor, rows);
    }
    if (subtotal_cursor < subtotals.size() && labels_equal(subtotals[subtotal_cursor].labels, path)) {
      rows.push_back({AxisLeaf{path}, true, subtotal_cursor});
      ++subtotal_cursor;
    }
  }
  path.pop_back();
}

std::vector<RowEntry> collect_row_entries(const PivotResult& result, std::size_t depth, bool include_subtotals) {
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
    collect_row_entries_impl(root, path, result.row_subtotals, subtotal_cursor, entries);
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

  std::vector<AxisLeaf> row_leaves = collect_row_leaves(result.rows, row_depth);
  std::vector<AxisLeaf> col_leaves = collect_col_leaves(result.cols, col_depth);
  const bool include_row_subtotals = !result.row_subtotals.empty();
  std::vector<RowEntry> row_entries = collect_row_entries(result, row_depth, include_row_subtotals);
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
  const std::size_t row_header_cols = row_depth == 0 ? 1 : row_depth;
  const std::size_t col_header_rows = col_depth == 0 ? 1 : col_depth;
  const std::size_t data_field_header_rows = data_field_count > 1 ? 1 : 0;
  const std::size_t header_rows = col_header_rows + data_field_header_rows + 1;
  const std::size_t total_rows = header_rows + row_entries.size() + (table.grand_totals_cols() ? 1 : 0);
  const std::size_t total_cols = row_header_cols + data_cols + (table.grand_totals_rows() ? data_field_count : 0);

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

  // Row field headers.
  for (std::size_t depth = 0; depth < row_header_cols; ++depth) {
    std::string label;
    if (depth < row_depth && table.row_field_order()[depth] < table.fields().size()) {
      label = field_display_name(table.fields()[table.row_field_order()[depth]]);
    }
    append_cell(cells, row_header_row, left + static_cast<std::uint32_t>(depth), text_value(cells, std::move(label)),
                PivotCellKind::Header, static_cast<std::uint32_t>(depth));
  }

  // Column hierarchy labels. Labels are repeated per leaf/data-field slot;
  // consumers can visually merge adjacent equal labels if desired.
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
        label += " " + options.grand_total_label;
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

  // Row labels, row subtotal rows, and data cells.
  std::size_t row_leaf_index = 0;
  for (std::size_t r_entry = 0; r_entry < row_entries.size(); ++r_entry) {
    const RowEntry& entry = row_entries[r_entry];
    const std::uint32_t row = data_top + static_cast<std::uint32_t>(r_entry);
    const AxisLeaf& leaf = entry.leaf;
    for (std::size_t depth = 0; depth < row_header_cols; ++depth) {
      std::string label;
      if (depth < leaf.labels.size()) {
        label = leaf.labels[depth];
      }
      if (entry.subtotal && depth + 1 == leaf.labels.size()) {
        label += " " + options.grand_total_label;
      }
      std::string field_name;
      if (depth < row_depth && table.row_field_order()[depth] < table.fields().size()) {
        field_name = field_display_name(table.fields()[table.row_field_order()[depth]]);
      }
      append_cell(cells, row, left + static_cast<std::uint32_t>(depth), text_value(cells, std::move(label)),
                  entry.subtotal ? PivotCellKind::RowSubtotal : PivotCellKind::RowLabel,
                  static_cast<std::uint32_t>(depth), std::move(field_name));
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
  if (table.grand_totals_rows()) {
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
  if (table.grand_totals_cols()) {
    const std::uint32_t total_row = data_top + static_cast<std::uint32_t>(row_entries.size());
    append_cell(cells, total_row, left, text_value(cells, options.grand_total_label), PivotCellKind::GrandTotal, 0);
    std::size_t total_col_leaf_index = 0;
    for (std::size_t c_entry = 0; c_entry < col_entries.size(); ++c_entry) {
      const ColEntry& col_entry = col_entries[c_entry];
      for (std::size_t df = 0; df < data_field_count; ++df) {
        const std::uint32_t col = data_left + static_cast<std::uint32_t>(c_entry * data_field_count + df);
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
    if (table.grand_totals_rows()) {
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
