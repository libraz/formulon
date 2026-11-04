//
// Pivot layout projection. This layer converts an evaluated `PivotResult`
// into absolute grid cells that a UI can draw with the ordinary cell
// renderer, each tagged with the role that decides its sheet format. It does
// not write into the owning Sheet's cell store.

#ifndef FORMULON_PIVOT_PIVOT_LAYOUT_H_
#define FORMULON_PIVOT_PIVOT_LAYOUT_H_

#include <cstdint>
#include <deque>
#include <optional>
#include <string>
#include <vector>

#include "excel_profile.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon::pivot {

enum class PivotCellKind : std::uint8_t {
  Header = 0,
  RowLabel = 1,
  ColLabel = 2,
  Data = 3,
  RowSubtotal = 4,
  ColSubtotal = 5,
  GrandTotal = 6,
  Blank = 7,
};

/// Per-locale label overrides for the pivot grid.
///
/// Most of these are consumed by the projection in `layout`, but
/// `blank_item_label` names an axis item and so has to be resolved while
/// the hierarchy is built: `pivot::evaluate` bakes it into
/// `PivotResult`'s node labels, which is what both the projection and
/// GETPIVOTDATA's exact-match walk read.
///
/// All fields except `grand_total_label` / `values_label` /
/// `blank_item_label` / `data_total_prefix` default to empty.
/// `values_label` is the caption a newly created pivot stores; an existing
/// pivot draws the caption it carries.
/// When `row_labels_label` (or `column_labels_label`) is empty the
/// projection keeps the historical English layout: the row-field (resp.
/// column-field) display name occupies the corresponding header cell.
/// When non-empty, the projection switches to Excel's "compact form"
/// placeholder layout used by Excel 365 and below, where the row-field
/// name is replaced by a single localized "Row Labels" placeholder.
///
/// Subtotal rows / columns label a group as `subtotal_prefix + group +
/// subtotal_suffix` (fr-FR puts its word first). When both are empty the
/// projection falls back to `" " + grand_total_label`. With several data
/// fields, each one's grand total on the Values axis is labelled
/// `data_total_prefix + data field name + data_total_suffix`.
///
/// `blank_item_label` names an axis group with no source value, which
/// must be named because an empty label can be neither drawn nor
/// addressed by GETPIVOTDATA. The label is applied after the axis is
/// ordered: Excel sorts the blank group last by its source value, not
/// first by this text. `all_pages_label` / `multiple_items_label` name a
/// page field's selection when no single item identifies it, and are
/// resolved alongside `blank_item_label` for the same reason. The
/// per-locale values come from `pivot_layout_options_for`; do not
/// hard-code them at a use site.
struct PivotLayoutOptions {
  std::string grand_total_label = "Grand Total";  ///< A table's `grand_total_caption` overrides it.
  std::string values_label = "Values";
  std::string blank_item_label = "(blank)";               ///< Label of an axis item with no value.
  std::string all_pages_label = "(All)";                  ///< Page field showing every item.
  std::string multiple_items_label = "(Multiple Items)";  ///< Page field showing several items.
  std::string row_labels_label;                           ///< e.g. "行ラベル"; empty disables.
  std::string column_labels_label;                        ///< e.g. "列ラベル"; empty disables.
  std::string subtotal_prefix;                            ///< e.g. "Total " (fr-FR); placed before the group label.
  std::string subtotal_suffix;                            ///< e.g. " 集計"; empty falls back to grand_total_label.
  std::string data_total_prefix = "Total ";        ///< Grand total of one of several data fields, e.g. "全体の ".
  std::string data_total_suffix;                   ///< e.g. "汇总" (zh-CN); placed after the data-field name.
  ExcelProfile profile = default_excel_profile();  ///< Collates item text while the pivot is evaluated.
};

/// What a projected cell is in the report Excel writes, which decides how its
/// sheet cell is formatted. `PivotCellKind` stays the coarser vocabulary the
/// bindings expose.
enum class PivotCellRole : std::uint8_t {
  None = 0,         ///< Structural blank: Excel writes no sheet cell.
  PageCaption,      ///< A page field's name.
  PageValue,        ///< The item a page field shows.
  FieldHeader,      ///< A field button: the row / column labels caption, a tabular field
                    ///< name, the Values caption, or a lone data field's caption.
  DataHeader,       ///< A data field's caption heading a value column.
  ColItem,          ///< A column item label.
  RowItem,          ///< A row item label.
  DataRowLabel,     ///< A data field's caption heading a row (Values on the rows).
  SubtotalLabel,    ///< A subtotal label, on either axis.
  GrandTotalLabel,  ///< A grand-total label, on either axis.
  DataValue,        ///< A value cell, empty ones included.
  SubtotalValue,    ///< A subtotal value cell.
  GrandTotalValue,  ///< A grand-total value cell.
  BlankRow,         ///< The label cell of an inserted blank row.
};

struct PivotCell {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  Value value = Value::blank();
  PivotCellKind kind = PivotCellKind::Blank;
  PivotCellRole role = PivotCellRole::None;
  std::uint32_t depth = 0;
  /// Indent of a left-aligned compact-form row label; unset for every other cell.
  std::optional<std::uint32_t> label_indent;
  std::string field_name;
  std::string number_format;
};

struct PivotCells {
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  std::vector<PivotCell> cells;
  /// Cell ranges Excel merges under `mergeItem`, in absolute sheet coordinates.
  std::vector<PivotRect> merges;

  // Backing storage for Text values emitted by the layout layer. This
  // includes synthesized labels and Text values copied from `PivotResult`,
  // so `PivotCells` can outlive the evaluated result it was projected from.
  std::deque<std::string> text_storage;
};

/// One axis item on the path to a value cell.
struct PivotItemRef {
  std::uint32_t field_index = 0;  ///< Index into `PivotTable::fields()`.
  /// The node's `AxisHierarchyNode::identity`: the record value, or the group
  /// item number when `is_group_item`.
  Value identity = Value::blank();
  bool is_group_item = false;
};

/// The axis items a value cell of the projected grid aggregates over.
///
/// Each list runs from the outermost field down to the cell's level: a leaf
/// lists every field of its axis, a subtotal stops at the field it totals, and
/// a grand total lists none. The Values pseudo-level is never listed.
struct CellItemPath {
  PivotCellKind kind = PivotCellKind::Data;  ///< Data, RowSubtotal, ColSubtotal or GrandTotal.
  std::vector<PivotItemRef> row_items;
  std::vector<PivotItemRef> col_items;
  /// Backing storage for Text identities, so the path outlives the result.
  std::deque<std::string> text_storage;
};

/// Projects an evaluated pivot result into absolute sheet coordinates.
///
/// The report body starts at `table.anchor_row()` / `anchor_col()`; the
/// page-field block, when present, sits above it. Missing row or column axes are represented as one
/// implicit leaf so the data area remains rectangular.
///
/// The table's display options shape the grid: `show_headers`,
/// `grid_drop_zones`, `page_wrap` / `page_over_then_down`, the error and
/// missing captions, `indent` (compact label indent) and `merge_item`
/// (tabular label merges).
Expected<PivotCells, Error> layout(const PivotTable& table, const PivotResult& result,
                                   const PivotLayoutOptions& options = PivotLayoutOptions{});

/// The axis items behind the value cell `layout` places at absolute (`row`,
/// `col`). Fails with `kInvalidArgument` for a cell outside the report and for
/// a header, label or blank cell, and with `kEvalPivotInvalid` when an item on
/// the path carries no identity.
Expected<CellItemPath, Error> cell_item_path(const PivotTable& table, const PivotResult& result,
                                             const PivotLayoutOptions& options, std::uint32_t row, std::uint32_t col);

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_PIVOT_LAYOUT_H_
