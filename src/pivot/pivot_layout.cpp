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
#include "pivot/pivot_layout_internal.h"
#include "pivot/pivot_locale.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/value_order.h"
#include "utils/checked_mul.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "value.h"

namespace formulon::pivot {
using namespace detail;

namespace {

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
  RETURN_IF_ERROR(validate_result_shape(table, result, row_leaves.size(), col_leaves.size()));
  ASSIGN_OR_RETURN(std::vector<AxisEntry> row_entries,
                   collect_axis_entries(table, result.rows, result.row_subtotals, row_shape, "row"));
  ASSIGN_OR_RETURN(std::vector<AxisEntry> col_entries_all,
                   collect_axis_entries(table, result.cols, result.col_subtotals, col_shape, "column"));
  // Column fields without a data field still lay out their items and total
  // column, over empty value cells; without column fields there is nothing to lay out.
  const bool valueless = data_field_count == 0 && col_depth > 0;
  std::vector<AxisEntry> col_entries =
      data_field_count == 0 && !valueless ? std::vector<AxisEntry>{} : std::move(col_entries_all);
  const std::size_t data_cols = col_entries.size();
  // Without row fields either, such a report is its caption and column-item rows alone.
  if (valueless && !legacy && row_levels == 0) {
    row_entries.clear();
  }
  // `gridDropZones` keeps the row-label column of a report without row fields.
  const bool grid_without_rows = !legacy && table.grid_drop_zones() && row_levels == 0 && data_field_count > 0;

  // Compact form merges every row level into one physical column (Excel
  // indents nested keys); Tabular / Outline give each level its own column,
  // the Values level included. A compact pivot with no row levels whose
  // data fields head the columns has no row-label column at all, nor has a
  // report without row or data fields.
  std::size_t row_header_cols = row_levels;
  if (row_levels == 0) {
    if (grid_without_rows) {
      row_header_cols = 1;
    } else {
      row_header_cols = (compact && values_on_cols) || (valueless && !legacy) ? 0 : 1;
    }
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
  // The classic layout adds a drop-zone row over a report without column
  // fields, except one holding a lone data field and nothing else.
  const std::size_t drop_zone_rows =
      !legacy && col_depth == 0 && table.grid_drop_zones() && !(grid_without_rows && data_field_count == 1) ? 1 : 0;
  header_rows += drop_zone_rows;
  // Compact form folds the grand-totals strip in axes that have no
  // hierarchy of their own: a no-column-fields pivot's per-row total
  // already lives in its data columns, and likewise a no-row-fields
  // pivot's per-column total. Tabular / Outline share that rule; the
  // English projection keeps emitting both strips.
  const bool emit_grand_totals_rows_strip =
      (data_field_count > 0 || valueless) && table.grand_totals_rows() && (legacy || col_depth > 0);
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
  ASSIGN_OR_RETURN(std::size_t cell_count, checked_mul_size_t(total_rows, total_cols));

  const std::uint32_t body_top = std::max(table.anchor_row(), static_cast<std::uint32_t>(page_rows));
  plan.cells_top = body_top - static_cast<std::uint32_t>(page_rows);
  plan.cell_count = cell_count;

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
  plan.valueless = valueless;
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
  plan.grid_without_rows = grid_without_rows;
  plan.emit_grand_totals_rows_strip = emit_grand_totals_rows_strip;
  plan.emit_grand_totals_cols_strip = emit_grand_totals_cols_strip;
  plan.total_strip_cols = total_strip_cols;
  plan.total_strip_rows = total_strip_rows;
  plan.page_rows = page_rows;
  plan.page_lines = page_lines;
  plan.total_rows = total_rows;
  plan.total_cols = total_cols;
  plan.body_cols = body_cols;
  return plan;
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
  ASSIGN_OR_RETURN(const LayoutPlan plan, build_layout_plan(table, result, options));

  PivotCells cells;
  cells.top = plan.cells_top;
  cells.left = plan.left;
  cells.rows = static_cast<std::uint32_t>(plan.total_rows);
  cells.cols = static_cast<std::uint32_t>(plan.total_cols);
  cells.body_top = plan.top;
  cells.body_rows = static_cast<std::uint32_t>(plan.total_rows - plan.page_rows);
  cells.body_cols = static_cast<std::uint32_t>(plan.body_cols);
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
  ASSIGN_OR_RETURN(const LayoutPlan plan, build_layout_plan(table, result, options));
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
    RETURN_IF_ERROR(append_item_refs(*row_entry, table.row_field_order(), path.row_items, path.text_storage));
  }
  if (col_entry != nullptr) {
    RETURN_IF_ERROR(append_item_refs(*col_entry, table.col_field_order(), path.col_items, path.text_storage));
  }
  return path;
}

}  // namespace formulon::pivot
