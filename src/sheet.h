// Copyright 2026 libraz. Licensed under the MIT License.
//
// Workbook sheet model. A `Sheet` owns a display name and a row-sparse,
// column-dense cell store keyed by 0-based row index. Excel sheets reach
// 1,048,576 rows by 16,384 columns but are overwhelmingly row-sparse, so
// a hash map of rows keeps memory proportional to populated rows while
// per-row dense vectors keep contiguous-range iteration (e.g. `SUM(A1:A100)`
// or OOXML row serialisation) cheap and cache-friendly.
//
// Sheets also own dynamic-array spill regions: when a formula returns a
// `Value::Array` it spills into adjacent cells (the formula owns the anchor;
// the rest are "phantoms"). Spill regions live on the sheet, are heap-owned
// (outliving the per-evaluation arena), and are eagerly invalidated when an
// underlying cell mutates. See `SpillRegion` and the `*_spill` member
// functions below for the full contract.

#ifndef FORMULON_SHEET_H_
#define FORMULON_SHEET_H_

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

#include "cell.h"
#include "value.h"

namespace formulon {

/// A registered spill region produced by a dynamic-array formula.
///
/// `cells` is row-major (size = `rows * cols`) and holds heap-owned `Value`
/// copies. Any `Text` cell's payload is interned in `owned_strings` so the
/// `SpillRegion`'s lifetime is independent of the per-evaluation arena that
/// produced the original `ArrayValue`. The order of strings in
/// `owned_strings` is insertion order (left-to-right, top-to-bottom across
/// `cells`), which is purely an implementation detail; consumers must not
/// rely on it.
struct SpillRegion {
  std::uint32_t anchor_row = 0;
  std::uint32_t anchor_col = 0;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  std::vector<Value> cells;
  // Backing store for Text values: a `Text` entry in `cells` is a string_view
  // into one of these strings. Kept as a separate vector so the `cells`
  // vector itself remains a dense array of `Value` (which is trivially
  // copyable).
  std::vector<std::string> owned_strings;
};

/// Hash for `CellAddress` suitable for `std::unordered_map`.
///
/// Excel addresses cap at row < 2^21 and col < 2^14 so a simple
/// `(row * 31) + col` mix is collision-free in the usable range and faster
/// than a generic 64-bit splat. The function is `noexcept` because the
/// underlying field accesses cannot throw.
struct CellAddressHash {
  std::size_t operator()(CellAddress a) const noexcept {
    std::size_t h = static_cast<std::size_t>(a.row);
    h = h * 31U + static_cast<std::size_t>(a.col);
    return h;
  }
};

// Private spill-table type defined in sheet.cpp. Only the unique_ptr<>
// declared in `Sheet` needs to know it exists.
struct SpillTable;

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
  ///
  /// Defined out-of-line because `spill_table_` (a `unique_ptr` to the
  /// forward-declared `SpillTable`) must see the complete deleter type at
  /// the point any constructor body is generated, including the implicit
  /// member-cleanup paths the compiler emits even under `-fno-exceptions`.
  explicit Sheet(std::string name);

  // Move-only. The explicit declarations are required because
  // `spill_table_` is a `unique_ptr<SpillTable>` to a forward-declared type;
  // the special members must be defined out-of-line where `SpillTable` is
  // complete. See sheet.cpp for the defaulted implementations.
  Sheet(const Sheet&) = delete;
  Sheet& operator=(const Sheet&) = delete;
  Sheet(Sheet&&) noexcept;
  Sheet& operator=(Sheet&&) noexcept;
  ~Sheet();

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
  ///
  /// If `(row, col)` is currently a phantom of a registered spill region,
  /// that region is eagerly cleared before the write proceeds: writing to a
  /// phantom semantically mutates the spill area, so the spill must be
  /// dropped. The spill anchor's own `cached_value` is left untouched and
  /// will be refreshed (or surface `#SPILL!`) on the next evaluation pass.
  void set_cell_value(std::uint32_t row, std::uint32_t col, Value v);

  /// Stores a formula at `(row, col)`.
  ///
  /// The cell's `formula_text` is replaced with `formula` (move-stored as-is;
  /// no validation that it begins with `=` — the parser owns that contract)
  /// and `cached_value` is reset to `Value::blank()` until the evaluator
  /// populates a result. Growth and bounds semantics match `set_cell_value`.
  ///
  /// Eagerly clears any spill region covering `(row, col)` as a phantom; see
  /// `set_cell_value` for the rationale.
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
  /// `cell_at` is intentionally narrow: it does not consult the spill
  /// table. Phantom cells of a spill region therefore return `nullptr` here
  /// (or whatever literal was stored before the spill was committed).
  /// Callers that need the spill-aware effective value should use
  /// `resolve_cell_value` instead.
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

  // ---------------------------------------------------------------------------
  // Dynamic-array spill API
  // ---------------------------------------------------------------------------

  /// Returns the spill region anchored at `(row, col)`, or `nullptr` when
  /// no region is anchored there. The returned pointer is valid until the
  /// next mutating call to the spill API (`commit_spill`, `clear_spill`)
  /// or to a cell-mutating call that triggers eager invalidation.
  const SpillRegion* spill_region_at_anchor(std::uint32_t row, std::uint32_t col) const noexcept;

  /// Returns the spill region whose phantom area covers `(row, col)`, or
  /// `nullptr` when no region covers it. Returns `nullptr` for the
  /// region's anchor cell itself: only phantoms are tracked in the
  /// reverse map. Use `spill_region_at_anchor` to look up by anchor.
  const SpillRegion* spill_region_covering(std::uint32_t row, std::uint32_t col) const noexcept;

  /// Returns the spill-aware effective value of `(row, col)`:
  ///
  ///   1. If the cell is a phantom of a spill region, returns the
  ///      corresponding row-major cell value from that region.
  ///   2. Else if a literal/formula `Cell` is stored at `(row, col)`,
  ///      returns its `cached_value`.
  ///   3. Else returns `Value::blank()`.
  ///
  /// The anchor cell of a spill region falls into case 2 (its
  /// `cached_value` was set by `commit_spill` to the region's first
  /// row-major cell), so anchors round-trip correctly without a special
  /// case in the caller.
  Value resolve_cell_value(std::uint32_t row, std::uint32_t col) const noexcept;

  /// Registers a spill region anchored at `(anchor_row, anchor_col)` with
  /// the given dimensions and row-major cell payload.
  ///
  /// Behaviour:
  ///
  ///   * Any existing region anchored at `(anchor_row, anchor_col)` is
  ///     cleared first (including its reverse-map entries).
  ///   * The would-be footprint is checked: every cell in the rectangle
  ///     except the anchor is "occupied" if (a) `cell_at` returns non-null
  ///     with a non-blank `cached_value` or non-empty `formula_text`, or
  ///     (b) it is already covered by another spill region. On any
  ///     occupied cell the anchor's `cached_value` is set to
  ///     `#SPILL!`, no region is registered, and the function returns
  ///     `false`. Pre-existing literals are preserved.
  ///   * On success, `cells` is deep-copied into the region (Text payloads
  ///     interned in `owned_strings`), reverse entries are written for each
  ///     phantom (the anchor itself is excluded from the reverse map), the
  ///     anchor's `cached_value` is set to `cells[0]`, and the function
  ///     returns `true`.
  ///   * A degenerate 1x1 region (`rows == cols == 1`) is accepted: no
  ///     phantoms are registered and only the anchor's `cached_value` is
  ///     written.
  ///   * Returns `false` (without side effect on the spill table) when
  ///     `rows == 0 || cols == 0`, when `cells.size() != rows * cols`, or
  ///     when the footprint would overflow `kMaxRows` / `kMaxCols`.
  bool commit_spill(std::uint32_t anchor_row, std::uint32_t anchor_col, std::uint32_t rows, std::uint32_t cols,
                    std::vector<Value> cells);

  /// Clears the spill region anchored at `(anchor_row, anchor_col)`.
  ///
  /// Removes every reverse-map entry for the region's phantoms and drops
  /// the by-anchor entry. Does not modify the anchor cell's
  /// `cached_value`; callers that want to overwrite the anchor (e.g. when
  /// the formula is being deleted) must do so separately.
  ///
  /// No-op when no region is anchored at `(anchor_row, anchor_col)`.
  void clear_spill(std::uint32_t anchor_row, std::uint32_t anchor_col) noexcept;

 private:
  std::string name_;
  std::unordered_map<std::uint32_t, std::vector<Cell>> rows_;
  // Lazily allocated: most sheets do not host any spill regions, so the
  // table is only materialised on the first `commit_spill` call.
  std::unique_ptr<SpillTable> spill_table_;
};

}  // namespace formulon

#endif  // FORMULON_SHEET_H_
