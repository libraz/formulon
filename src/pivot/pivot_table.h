// Copyright 2026 libraz. Licensed under the MIT License.
//
// `PivotTable` is the workbook-level pivot definition: identity, cache
// binding, field configuration, layout, anchor location, transient slicer
// state, and the most recent evaluation result. One pivot table is owned
// by exactly one sheet (the sheet supplies the sheet identity); it points
// at a workbook-owned `PivotCache` by id. See
// backup/plans/15-pivot-and-advanced.md §15.1.

#ifndef FORMULON_PIVOT_PIVOT_TABLE_H_
#define FORMULON_PIVOT_PIVOT_TABLE_H_

#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "pivot/pivot_result.h"
#include "pivot/pivot_types.h"

namespace formulon::pivot {

/// One pivot table definition + most recent evaluation result.
///
/// Move-only. Designed to be held by the owning `Sheet` via
/// `std::unique_ptr<PivotTable>` so the address is stable for as long as
/// the sheet retains the pivot.
class PivotTable {
 public:
  PivotTable() = default;

  PivotTable(const PivotTable&) = delete;
  PivotTable& operator=(const PivotTable&) = delete;
  PivotTable(PivotTable&&) = default;
  PivotTable& operator=(PivotTable&&) = default;

  // Identity / cache binding -------------------------------------------------

  const std::string& name() const { return name_; }
  void set_name(std::string n) { name_ = std::move(n); }

  std::uint32_t pivot_cache_id() const { return pivot_cache_id_; }
  void set_pivot_cache_id(std::uint32_t id) { pivot_cache_id_ = id; }

  // Field configuration ------------------------------------------------------

  const std::vector<PivotField>& fields() const { return fields_; }
  std::vector<PivotField>& mutable_fields() { return fields_; }

  /// Data-field entries from `<dataFields>`. Keyed by display name for
  /// GETPIVOTDATA; one source `PivotField` may have multiple data-field
  /// entries (e.g. Sum + Average of the same column).
  const std::vector<PivotDataField>& data_fields() const { return data_fields_; }
  std::vector<PivotDataField>& mutable_data_fields() { return data_fields_; }

  /// Document-order indices into `fields()` of the row-axis fields,
  /// captured from `<rowFields>`. Used by the pivot evaluator to walk
  /// the row hierarchy.
  const std::vector<std::uint32_t>& row_field_order() const { return row_field_order_; }
  std::vector<std::uint32_t>& mutable_row_field_order() { return row_field_order_; }

  /// Document-order indices into `fields()` of the column-axis fields,
  /// captured from `<colFields>`.
  const std::vector<std::uint32_t>& col_field_order() const { return col_field_order_; }
  std::vector<std::uint32_t>& mutable_col_field_order() { return col_field_order_; }

  // Layout / location --------------------------------------------------------

  PivotLayout layout() const { return layout_; }
  void set_layout(PivotLayout l) { layout_ = l; }

  std::uint32_t anchor_row() const { return anchor_row_; }
  std::uint32_t anchor_col() const { return anchor_col_; }
  std::uint32_t span_rows() const { return span_rows_; }
  std::uint32_t span_cols() const { return span_cols_; }

  void set_anchor(std::uint32_t row, std::uint32_t col, std::uint32_t rows, std::uint32_t cols) {
    anchor_row_ = row;
    anchor_col_ = col;
    span_rows_ = rows;
    span_cols_ = cols;
  }

  /// True iff `(row, col)` is inside the pivot's bounds.
  bool contains(std::uint32_t row, std::uint32_t col) const noexcept {
    return row >= anchor_row_ && row < anchor_row_ + span_rows_ && col >= anchor_col_ && col < anchor_col_ + span_cols_;
  }

  // Transient slicer-applied filters ----------------------------------------
  //
  // Kept separate from `PivotField::items` (which is the XML-defined manual
  // filter state). Folding slicer state into `items` would make the
  // `evaluator -> XML round-trip` path overwrite the authored definition.
  const std::vector<PivotFilter>& active_filters() const { return active_filters_; }
  std::vector<PivotFilter>& mutable_active_filters() { return active_filters_; }

  // Most-recent evaluation result -------------------------------------------
  //
  // `last_result_` is `mutable` because it is a logical-const memoisation
  // slot: GETPIVOTDATA observes a `PivotTable` through a `const Workbook&`
  // (`EvalContext::workbook()`) but still needs to refresh the result
  // cache on demand if the OOXML reader did not pre-compute it. Marking
  // the field `mutable` lets the lazy form refresh through the const
  // accessor without `const_cast` gymnastics; the public surface remains
  // const-correct because callers can only read the result through the
  // `last_result()` accessor.

  const std::optional<PivotResult>& last_result() const { return last_result_; }
  /// Returns the underlying `optional` for mutation. Callable through a
  /// const reference because the result cache is logical-const
  /// memoisation (see the `mutable` rationale on `last_result_`).
  std::optional<PivotResult>& mutable_last_result() const { return last_result_; }

  // Grand totals layout flags -----------------------------------------------

  bool grand_totals_rows() const { return grand_totals_rows_; }
  bool grand_totals_cols() const { return grand_totals_cols_; }
  void set_grand_totals(bool rows, bool cols) {
    grand_totals_rows_ = rows;
    grand_totals_cols_ = cols;
  }

 private:
  std::string name_;
  std::uint32_t pivot_cache_id_ = 0;
  std::vector<PivotField> fields_;
  std::vector<PivotDataField> data_fields_;
  std::vector<std::uint32_t> row_field_order_;
  std::vector<std::uint32_t> col_field_order_;
  PivotLayout layout_ = PivotLayout::Compact;
  std::uint32_t anchor_row_ = 0;
  std::uint32_t anchor_col_ = 0;
  std::uint32_t span_rows_ = 0;
  std::uint32_t span_cols_ = 0;
  bool grand_totals_rows_ = true;
  bool grand_totals_cols_ = true;
  std::vector<PivotFilter> active_filters_;
  // Logical-const memoisation slot for the most recent evaluation result.
  // GETPIVOTDATA refreshes this through `const PivotTable&` so the field
  // must be `mutable`; see the docstring on `last_result()` /
  // `mutable_last_result()` above.
  mutable std::optional<PivotResult> last_result_;
};

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_PIVOT_TABLE_H_
