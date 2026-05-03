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
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include "cell.h"
#include "cf/cf_types.h"
#include "value.h"

namespace formulon {

namespace pivot {
class PivotTable;
}  // namespace pivot

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

/// A merged cell rectangle on a sheet.
///
/// Coordinates are 0-based and inclusive on both ends, matching the OOXML
/// `<mergeCell ref="A1:B2">` shape after conversion. The merge reader/writer
/// bundle owns this list; the evaluator treats merged ranges as a display
/// concern and does not consult them. The schema may be extended by later
/// bundles, but the field names declared here are stable.
struct MergeRange {
  std::uint32_t first_row = 0;
  std::uint32_t last_row = 0;
  std::uint32_t first_col = 0;
  std::uint32_t last_col = 0;
};

/// A hyperlink anchored on a single cell.
///
/// `target` is the resolved URL or in-workbook target (e.g. `Sheet2!A1`),
/// `display` is the surface text shown in the cell when distinct from the
/// cell's stored value, and `tooltip` populates the OOXML `tooltip` attribute.
/// `rid` is the source rels id from the sheet part, retained verbatim so
/// round-trip writes can preserve relationship ordering. The schema may be
/// extended by later bundles, but the field names declared here are stable.
struct Hyperlink {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::string target;
  std::string display;
  std::string tooltip;
  std::string rid;
};

/// A cell-anchored comment / threaded-comment payload.
///
/// The schema is intentionally minimal at this point and will grow as the
/// comment reader/writer bundle wires up rich-text runs, threading metadata,
/// and modern ("threaded") comment ids. The field names declared here are
/// stable; new fields are appended.
struct CellComment {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::string author;
  std::string text;
};

/// A cell-range data-validation rule.
///
/// Schema deliberately empty at this point; the validation reader/writer
/// bundle expands this with type / formula1 / formula2 / errorTitle /
/// errorMessage / promptTitle / prompt fields. Existing fields (none) stay
/// stable; new fields are appended.
struct DataValidation {};

/// Per-sheet view state (zoom, frozen panes, tab visibility).
///
/// Mirrors the OOXML `<sheetView>` and `<sheetPr>` attributes that survive
/// a save/load round-trip without affecting evaluation. The schema may be
/// extended by later bundles, but the field names declared here are stable.
struct SheetView {
  /// Default zoom percentage when no `<sheetView zoomScale="...">` was
  /// stored. Matches Excel's "100%" default.
  static constexpr std::uint32_t kDefaultZoomScale = 100U;

  std::uint32_t zoom_scale = kDefaultZoomScale;
  std::uint32_t freeze_rows = 0;  // 0 = no row freeze
  std::uint32_t freeze_cols = 0;  // 0 = no column freeze
  bool tab_hidden = false;
};

/// Layout overrides for a contiguous column span.
///
/// Mirrors the OOXML `<col min="..." max="..." width="..." hidden="..."
/// outlineLevel="...">` shape. Both endpoints are 0-based and inclusive. The
/// schema may be extended by later bundles, but the field names declared
/// here are stable.
struct ColumnLayout {
  std::uint32_t first = 0;  // 0-based, inclusive
  std::uint32_t last = 0;   // 0-based, inclusive
  double width = 0.0;       // in OOXML character-width units
  bool hidden = false;
  std::uint8_t outline_level = 0;
};

/// Layout override for a single row.
///
/// Mirrors the OOXML `<row r="..." ht="..." hidden="..." outlineLevel="...">`
/// attributes that override the default row metrics. The schema may be
/// extended by later bundles, but the field names declared here are stable.
struct RowLayout {
  std::uint32_t row = 0;  // 0-based
  double height = 0.0;    // in points
  bool hidden = false;
  std::uint8_t outline_level = 0;
};

/// Aggregate of per-sheet layout overrides (column spans + row overrides).
///
/// Both lists are empty by default; the OOXML reader populates them from
/// `<cols>` and `<row>` entries that carry non-default attributes. The
/// schema may be extended by later bundles, but the field names declared
/// here are stable.
struct SheetLayout {
  std::vector<ColumnLayout> columns;
  std::vector<RowLayout> row_overrides;
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

  /// Updates only the cached `Value` of an existing cell at `(row, col)`.
  ///
  /// The cell's `formula_text` is preserved exactly — this is the post-
  /// evaluation update path used by the recalc engine, which has just
  /// computed a fresh value for a formula cell and needs to write it back
  /// without erasing the formula. If the cell does not yet exist, it is
  /// created as a plain literal (empty `formula_text`, `cached_value = v`),
  /// growing the row vector exactly as `set_cell_value` would.
  ///
  /// Text-payload lifetime: when `v.is_text()`, the bytes are deep-copied
  /// into the cell's own `cached_text_owned` storage and the stored
  /// `cached_value` is rewritten to reference that internal buffer. This
  /// is what lets the recalc engine reset its per-evaluation `Arena`
  /// between cells without dangling any Text scalar previously committed
  /// here. The `set_cell_value` path retains the caller-owns lifetime
  /// contract (used by the OOXML reader where the SST owns the bytes and
  /// outlives the cell).
  ///
  /// Unlike `set_cell_value`, this method does NOT eagerly clear any
  /// spill region covering `(row, col)`: cached-value updates are not
  /// considered structural mutations of the cell, so phantoms continue to
  /// reflect their owning anchor's array. The recalc engine separately
  /// commits any new spill before reaching this method via
  /// `EvalContext::dispatch_array_result`.
  void set_cell_cached_value(std::uint32_t row, std::uint32_t col, Value v);

  /// Stores a phonetic-kana annotation on the cell at `(row, col)`.
  ///
  /// Used by the OOXML reader after a Text cell is committed to copy the
  /// `<rPh>` payload from the cell's source `<si>` (SST) or `<is>`
  /// (inline string) into the cell's `phonetic_text` field. PHONETIC
  /// reads this field directly via the lazy dispatch path; the writer
  /// emits it back as `<rPh>` inside the `<is>` block on save.
  ///
  /// If the cell does not yet exist at `(row, col)` it is created as a
  /// default-constructed slot (empty `formula_text`, blank
  /// `cached_value`); growth and bounds semantics match
  /// `set_cell_value`. Passing an empty `phonetic` clears any previous
  /// annotation. This method does NOT touch the cell's stored value or
  /// formula, and does NOT eagerly clear any covering spill region.
  void set_cell_phonetic(std::uint32_t row, std::uint32_t col, std::string_view phonetic);

  /// Stores the cellXfs index for the cell at `(row, col)`. The cell must
  /// already exist (created via `set_cell_value` / `set_cell_formula`); on
  /// an absent cell this method is a no-op. `xf_index = 0` references the
  /// workbook's default cellXf and is the on-disk default Excel writes when
  /// no `s=` attribute is present.
  void set_cell_xf_index(std::uint32_t row, std::uint32_t col, std::uint32_t xf_index);

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

  // ---------------------------------------------------------------------------
  // Pivot tables anchored on this sheet
  // ---------------------------------------------------------------------------
  //
  // A worksheet may host any number of pivot tables; each is owned via
  // `unique_ptr` so its address is stable for the lifetime of the sheet.
  // The OOXML reader populates this list at workbook-load time; the pivot
  // evaluator reads from it (anchor lookup, result-cache refresh) and
  // GETPIVOTDATA's lazy form drives an on-demand evaluation through it.

  /// Read-only access to the pivot tables anchored on this sheet, in
  /// document-discovery (insertion) order.
  const std::vector<std::unique_ptr<pivot::PivotTable>>& pivot_tables() const noexcept { return pivot_tables_; }

  /// Mutable access to the pivot table list. Exposed so the evaluator (or
  /// GETPIVOTDATA's on-demand refresh path) can reach the underlying
  /// `PivotTable` and refresh its cached `last_result_`. The list itself
  /// is rarely mutated post-load.
  std::vector<std::unique_ptr<pivot::PivotTable>>& mutable_pivot_tables() noexcept { return pivot_tables_; }

  /// Appends a pivot table anchored on this sheet. Ownership transfers to
  /// the sheet; the returned reference (via the unique_ptr) is stable
  /// because each table is heap-allocated. No validation is performed
  /// here; callers are responsible for ensuring the table's anchor is
  /// inside the sheet bounds and references a valid cache id.
  void add_pivot_table(std::unique_ptr<pivot::PivotTable> table);

  // ---------------------------------------------------------------------------
  // Conditional formats attached to this sheet
  // ---------------------------------------------------------------------------
  //
  // Each entry is one `<conditionalFormatting>` block from the OOXML sheet
  // part: an sqref union plus a list of rules. Stored by value (the model
  // is plain POD; no result cache, no stable-address requirement). The
  // OOXML reader populates this list at workbook-load time; the writer
  // round-trips it; the evaluator (future PR) walks it per-cell or per-
  // viewport to compute matches.

  /// Read-only access to the conditional-format blocks attached to this
  /// sheet, in document-discovery (insertion) order. Across blocks the
  /// rules' `priority` is workbook-global; consumers that need
  /// priority-ordered evaluation must merge across blocks themselves.
  const std::vector<cf::ConditionalFormat>& conditional_formats() const noexcept { return conditional_formats_; }

  /// Mutable access to the conditional-format list. Exposed so the OOXML
  /// reader (and the future editing API) can append, replace, or
  /// reorder blocks without an extra accessor pair per field.
  std::vector<cf::ConditionalFormat>& mutable_conditional_formats() noexcept { return conditional_formats_; }

  // ---------------------------------------------------------------------------
  // Merged cell rectangles
  // ---------------------------------------------------------------------------

  /// Read-only access to the merged-cell rectangles attached to this sheet.
  /// Populated by the OOXML reader; empty until the corresponding format
  /// support lands.
  const std::vector<MergeRange>& merges() const noexcept { return merges_; }

  /// Mutable access to the merged-cell rectangle list. Exposed so the
  /// OOXML reader and future editing API can append, replace, or reorder
  /// entries without an extra accessor pair per field.
  std::vector<MergeRange>& mutable_merges() noexcept { return merges_; }

  // ---------------------------------------------------------------------------
  // Hyperlinks
  // ---------------------------------------------------------------------------

  /// Read-only access to the hyperlinks attached to this sheet. Populated
  /// by the OOXML reader; empty until the corresponding format support
  /// lands.
  const std::vector<Hyperlink>& hyperlinks() const noexcept { return hyperlinks_; }

  /// Mutable access to the hyperlink list. Exposed so the OOXML reader
  /// and future editing API can append, replace, or reorder entries
  /// without an extra accessor pair per field.
  std::vector<Hyperlink>& mutable_hyperlinks() noexcept { return hyperlinks_; }

  // ---------------------------------------------------------------------------
  // Cell comments
  // ---------------------------------------------------------------------------

  /// Read-only access to the cell comments attached to this sheet.
  /// Populated by the OOXML reader; empty until the corresponding format
  /// support lands.
  const std::vector<CellComment>& comments() const noexcept { return comments_; }

  /// Mutable access to the cell-comment list. Exposed so the OOXML reader
  /// and future editing API can append, replace, or reorder entries
  /// without an extra accessor pair per field.
  std::vector<CellComment>& mutable_comments() noexcept { return comments_; }

  // ---------------------------------------------------------------------------
  // Data validations
  // ---------------------------------------------------------------------------

  /// Read-only access to the data-validation rules attached to this sheet.
  /// Populated by the OOXML reader; empty until the corresponding format
  /// support lands.
  const std::vector<DataValidation>& validations() const noexcept { return validations_; }

  /// Mutable access to the data-validation list. Exposed so the OOXML
  /// reader and future editing API can append, replace, or reorder
  /// entries without an extra accessor pair per field.
  std::vector<DataValidation>& mutable_validations() noexcept { return validations_; }

  // ---------------------------------------------------------------------------
  // Sheet view state
  // ---------------------------------------------------------------------------

  /// Read-only access to the sheet's view state (zoom, frozen panes, tab
  /// visibility). Populated by the OOXML reader; carries default values
  /// until the corresponding format support lands.
  const SheetView& view() const noexcept { return view_; }

  /// Mutable access to the sheet's view state. Exposed so the OOXML
  /// reader and future editing API can update view fields without an
  /// extra accessor pair per field.
  SheetView& mutable_view() noexcept { return view_; }

  // ---------------------------------------------------------------------------
  // Sheet layout (column spans + row overrides)
  // ---------------------------------------------------------------------------

  /// Read-only access to the sheet's layout overrides. Populated by the
  /// OOXML reader; empty until the corresponding format support lands.
  const SheetLayout& layout() const noexcept { return layout_; }

  /// Mutable access to the sheet's layout overrides. Exposed so the
  /// OOXML reader and future editing API can append, replace, or
  /// reorder entries without an extra accessor pair per field.
  SheetLayout& mutable_layout() noexcept { return layout_; }

 private:
  std::string name_;
  std::unordered_map<std::uint32_t, std::vector<Cell>> rows_;
  // Lazily allocated: most sheets do not host any spill regions, so the
  // table is only materialised on the first `commit_spill` call.
  std::unique_ptr<SpillTable> spill_table_;
  // Pivot tables anchored on this sheet. Empty by default; populated by
  // the OOXML reader at workbook-load time. Heap-owned so addresses stay
  // stable across vector reallocations.
  std::vector<std::unique_ptr<pivot::PivotTable>> pivot_tables_;
  // Conditional-format blocks attached to this sheet. Empty by default;
  // populated by the OOXML reader from `<conditionalFormatting>` blocks.
  std::vector<cf::ConditionalFormat> conditional_formats_;
  // Merged-cell rectangles. Empty by default; populated by the OOXML
  // reader once the corresponding format support lands.
  std::vector<MergeRange> merges_;
  // Hyperlinks anchored on cells in this sheet. Empty by default;
  // populated by the OOXML reader once the corresponding format support
  // lands.
  std::vector<Hyperlink> hyperlinks_;
  // Cell comments / threaded comments. Empty by default; populated by the
  // OOXML reader once the corresponding format support lands.
  std::vector<CellComment> comments_;
  // Data-validation rules. Empty by default; populated by the OOXML
  // reader once the corresponding format support lands.
  std::vector<DataValidation> validations_;
  // Per-sheet view state (zoom, frozen panes, tab visibility). Defaults
  // are populated inline; the OOXML reader overwrites them once the
  // corresponding format support lands.
  SheetView view_;
  // Per-sheet layout overrides (column spans + row overrides). Empty by
  // default; populated by the OOXML reader once the corresponding format
  // support lands.
  SheetLayout layout_;
};

}  // namespace formulon

#endif  // FORMULON_SHEET_H_
