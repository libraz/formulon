// Copyright 2026 libraz. Licensed under the MIT License.
//
// Workbook sheet model. A `Sheet` owns a display name and a row-sparse,
// column-dense cell store keyed by 0-based row index. Excel sheets reach
// 1,048,576 rows by 16,384 columns but are overwhelmingly row-sparse, so
// a hash map of rows keeps memory proportional to populated rows while
// per-row dense vectors keep contiguous-range iteration (e.g. `SUM(A1:A100)`
// or OOXML row serialisation) cheap and cache-friendly.

#ifndef FORMULON_SHEET_H_
#define FORMULON_SHEET_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

#include "cell.h"
#include "value.h"

namespace formulon {

/// A single worksheet inside a `Workbook`.
///
/// Owns the worksheet's display name and its cell store. The cell store is
/// a `unordered_map<row, vector<Cell>>` where the inner vector is grown on
/// demand to cover the highest column touched in that row; columns not yet
/// touched are absent from the vector. Rows that have never been touched
/// are absent from the map.
class Sheet {
 public:
  /// Excel 365 maximum row count (rows are addressable as 0..kMaxRows-1).
  static constexpr std::uint32_t kMaxRows = 1048576U;

  /// Excel 365 maximum column count (columns are addressable as
  /// 0..kMaxCols-1, mapping to A..XFD).
  static constexpr std::uint32_t kMaxCols = 16384U;

  /// Builds a sheet with the given display name. The name is adopted
  /// verbatim; callers are expected to supply a valid Excel sheet name
  /// (name validation will live in the workbook layer once it is wired up).
  explicit Sheet(std::string name) : name_(std::move(name)) {}

  /// Current display name of the sheet.
  const std::string& name() const noexcept { return name_; }

  /// Replaces the display name.
  void set_name(std::string name) { name_ = std::move(name); }

  /// Stores a literal value at `(row, col)`.
  ///
  /// The cell's `formula_text` is reset to empty and `cached_value` is set
  /// to `v`. If the row's vector is shorter than `col + 1`, it is grown
  /// with default-constructed `Cell` instances (empty formula, blank
  /// cached value). If the row is not yet present in the map, it is
  /// created. `row` must satisfy `row < kMaxRows` and `col < kMaxCols`;
  /// out-of-range coordinates trip a debug assert. Callers (parser, OOXML
  /// reader) are responsible for validating coordinates before invoking
  /// the storage layer.
  void set_cell_value(std::uint32_t row, std::uint32_t col, Value v);

  /// Stores a formula at `(row, col)`.
  ///
  /// The cell's `formula_text` is replaced with `formula` (move-stored as-is;
  /// no validation that it begins with `=` — the parser owns that contract)
  /// and `cached_value` is reset to `Value::blank()` until the evaluator
  /// populates a result. Growth and bounds semantics match `set_cell_value`.
  void set_cell_formula(std::uint32_t row, std::uint32_t col, std::string formula);

  /// Returns a non-owning pointer to the cell at `(row, col)`, or `nullptr`
  /// when the coordinate is not in storage.
  ///
  /// "Not in storage" means either the row has never been touched (row
  /// missing from the map) or the row's vector has not been grown to cover
  /// `col`. A returned pointer may still reference a default-constructed
  /// `Cell` (empty formula and blank cached value): this happens when the
  /// column was implicitly created while growing the row vector to cover a
  /// later column. Callers that need to distinguish "explicitly blank" from
  /// "implicitly default" should check
  /// `cell->formula_text.empty() && cell->cached_value.is_blank()`.
  ///
  /// The returned pointer is invalidated by any mutation that grows the
  /// row's vector or rehashes the row map (i.e. by any subsequent
  /// `set_cell_value` / `set_cell_formula` call); callers must not retain
  /// it across mutations.
  const Cell* cell_at(std::uint32_t row, std::uint32_t col) const noexcept;

  /// Convenience predicate equivalent to `cell_at(row, col) != nullptr`.
  bool has_cell(std::uint32_t row, std::uint32_t col) const noexcept;

  /// Total number of stored `Cell` slots across all populated rows.
  ///
  /// Counts every slot in every populated row's vector, including
  /// implicitly default-constructed cells created by growth. Useful for
  /// tests and as a coarse memory-footprint indicator.
  std::size_t cell_count() const noexcept;

  /// Read-only access to the underlying row map.
  ///
  /// Exposed so consumers (e.g. the OOXML writer) can iterate populated
  /// rows in their own order without paying for an intermediate copy. The
  /// reference is invalidated by mutating Sheet operations.
  const std::unordered_map<std::uint32_t, std::vector<Cell>>& rows() const noexcept { return rows_; }

 private:
  std::string name_;
  std::unordered_map<std::uint32_t, std::vector<Cell>> rows_;
};

}  // namespace formulon

#endif  // FORMULON_SHEET_H_
