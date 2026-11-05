//
// Types and entry points shared by the pivot layout translation units:
// axis planning (`pivot_layout_axis.cpp`), cell emission
// (`pivot_layout_emit.cpp`) and plan construction (`pivot_layout.cpp`).

#ifndef FORMULON_PIVOT_PIVOT_LAYOUT_INTERNAL_H_
#define FORMULON_PIVOT_PIVOT_LAYOUT_INTERNAL_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "pivot/pivot_layout.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon::pivot::detail {

struct AxisLeaf {
  std::vector<std::string> labels;
};

/// Columns a page-field header row occupies: the field name and the item it
/// is showing. Excel leaves the rest of the row empty however wide the
/// report below it is.
constexpr std::size_t kPageHeaderCols = 2;

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
  /// Column fields without a data field: items and totals lay out over empty value cells.
  bool valueless = false;

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
  /// `gridDropZones` over a report with data fields and an empty row axis.
  bool grid_without_rows = false;
  bool emit_grand_totals_rows_strip = false;
  bool emit_grand_totals_cols_strip = false;
  std::size_t total_strip_cols = 0;
  std::size_t total_strip_rows = 0;
  std::size_t page_rows = 0;
  std::size_t page_lines = 0;  ///< Page fields per column (down-then-over) or row (over-then-down).
  std::size_t total_rows = 0;
  std::size_t total_cols = 0;
  std::size_t body_cols = 0;  ///< Columns of the report body, the page block excluded.
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

// Axis planning (pivot_layout_axis.cpp).
std::vector<AxisLeaf> collect_axis_leaves(const std::vector<AxisHierarchyNode>& roots, std::size_t depth);
std::vector<std::size_t> subtotal_counts_for_axis(const PivotTable& table,
                                                  const std::vector<std::uint32_t>& field_order);
std::vector<const PivotField*> fields_for_axis(const PivotTable& table, const std::vector<std::uint32_t>& field_order);
std::vector<std::vector<SubtotalFn>> subtotal_fns_for_axis(const PivotTable& table,
                                                           const std::vector<std::uint32_t>& field_order);

/// Projects one axis hierarchy into its entries. `axis` names the axis in
/// diagnostics. Instantiated for `RowSubtotal` and `ColSubtotal`.
template <typename Subtotal>
Expected<std::vector<AxisEntry>, Error> collect_axis_entries(const PivotTable& table,
                                                             const std::vector<AxisHierarchyNode>& roots,
                                                             const std::vector<Subtotal>& subtotals,
                                                             const AxisShape& shape, const std::string& axis);

// Cell emission (pivot_layout_emit.cpp).
void emit_page_block(const LayoutPlan& plan, PivotCells& cells);
void emit_header_block(const LayoutPlan& plan, PivotCells& cells);
void emit_body_rows(const LayoutPlan& plan, PivotCells& cells);
void emit_row_totals_strip(const LayoutPlan& plan, PivotCells& cells);
void emit_col_totals_strip(const LayoutPlan& plan, PivotCells& cells);
void emit_merges(const LayoutPlan& plan, PivotCells& cells);

}  // namespace formulon::pivot::detail

#endif  // FORMULON_PIVOT_PIVOT_LAYOUT_INTERNAL_H_
