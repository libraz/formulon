// Copyright 2026 libraz. Licensed under the MIT License.
//
// Top-level workbook model. The current surface owns a vector of `Sheet`
// instances and exposes a `save()` method that serialises the workbook to
// a `.xlsx` byte stream via the OOXML writer slice. Defined-name and
// table metadata are preserved here as passive round-trip state so the
// reader/writer pipeline can carry them through unchanged; named-range
// resolution at evaluation time and structured-reference parsing arrive
// in Phase 4. Styles, shared strings, pivots and the full cell store
// will be layered on in follow-up work (see backup/plans/04-xlsx-io.md).

#ifndef FORMULON_WORKBOOK_H_
#define FORMULON_WORKBOOK_H_

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <string_view>
#include <vector>

#include "io/defined_names.h"
#include "io/passthrough_part.h"
#include "io/tables_reader.h"
#include "io/workbook_kind.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {

namespace eval {
class FunctionRegistry;
class RecalcEngine;
struct IterativeOptions;
struct RecalcStats;
struct SchedulerConfig;
struct SchedulerStats;
}  // namespace eval

namespace pivot {
class PivotCache;
}  // namespace pivot

/// Move-only container representing a spreadsheet workbook.
///
/// Instances are constructed via the `create()` factory, which returns a
/// default workbook containing a single `"Sheet1"`. Additional sheets will
/// be exposed through explicit mutation APIs in follow-up work; callers
/// currently only mutate the one sheet's display name.
class Workbook {
 public:
  /// Factory for the default workbook: a single sheet named `"Sheet1"`.
  static Workbook create();

  /// Factory for a workbook with no sheets at all.
  ///
  /// Designed for I/O paths (notably `io::read_ooxml`) that need to
  /// reconstruct a workbook from an external description and append
  /// sheets in source-defined order; using `create()` would require the
  /// reader to mutate the implicit `"Sheet1"` placeholder before adding
  /// the real sheets, which is awkward and makes empty-archive handling
  /// ambiguous. Most callers should still use `create()`; this factory
  /// is intentionally separate so the "default Sheet1" guarantee of
  /// `create()` stays unsurprising.
  ///
  /// Move-only contract is identical to `create()`. Note that a
  /// zero-sheet workbook is invalid input for `save()` — Excel rejects
  /// empty sheet lists, and so do we.
  static Workbook create_empty();

  Workbook(const Workbook&) = delete;
  Workbook& operator=(const Workbook&) = delete;
  Workbook(Workbook&&) noexcept;
  Workbook& operator=(Workbook&&) noexcept;
  ~Workbook();

  /// Number of sheets in the workbook. Always at least 1 for
  /// `create()`-constructed instances.
  std::size_t sheet_count() const noexcept { return sheets_.size(); }

  /// Immutable access to the sheet at `index`. `index` must be `<
  /// sheet_count()`.
  const Sheet& sheet(std::size_t index) const { return sheets_[index]; }

  /// Mutable access to the sheet at `index`. `index` must be `<
  /// sheet_count()`.
  Sheet& sheet(std::size_t index) { return sheets_[index]; }

  /// Appends a new sheet with display name `name` and returns a reference to
  /// it. The Workbook retains ownership and returned references are
  /// invalidated by subsequent `add_sheet` calls (which may reallocate the
  /// underlying vector). Duplicate names are not rejected at this layer;
  /// OOXML-level name validation lives at the I/O boundary.
  Sheet& add_sheet(std::string name);

  /// Case-insensitive lookup by display name. Matches using
  /// `strings::case_insensitive_eq` (ASCII-fold), so `"SHEET2"` and
  /// `"Sheet2"` locate the same sheet — consistent with Excel's
  /// case-insensitive sheet-name semantics. Returns `nullptr` when no sheet
  /// matches. A linear scan is used; workbooks typically carry O(1)–O(10)
  /// sheets, so a hash index is premature.
  const Sheet* sheet_by_name(std::string_view name) const noexcept;

  /// Returns the 0-based index of the sheet whose display name matches
  /// `name` (case-insensitive, ASCII-fold). Returns `static_cast<size_t>(-1)`
  /// when no sheet matches. Linear scan, like `sheet_by_name`.
  std::size_t sheet_index_by_name(std::string_view name) const noexcept;

  /// Serialises the workbook to an in-memory `.xlsx` byte stream. Delegates
  /// to `io::write_ooxml`; see that function's documentation for the exact
  /// set of OOXML parts emitted by the empty-workbook writer slice.
  Expected<std::vector<std::uint8_t>, Error> save() const;

  // ---------------------------------------------------------------------------
  // Recalc-engine integration
  // ---------------------------------------------------------------------------
  //
  // Mutation APIs that participate in the dependency graph live at the
  // workbook layer so the embedded `RecalcEngine` can register / dirty
  // cells in lockstep with the underlying `Sheet` storage. Direct
  // mutation through `sheet(i).set_cell_*` still works, but it bypasses
  // the dep graph: callers that want incremental recalc must route their
  // edits through `set_cell_value` / `set_cell_formula` here.

  /// Stores a literal `value` at `(row, col)` on `sheet_index`.
  ///
  /// Side effects:
  ///   * Drops any dep-graph edges previously owned by the cell (the cell
  ///     is no longer a formula).
  ///   * Marks the cell dirty so any dependent formulas re-evaluate on
  ///     the next `recalc()` pass.
  ///   * Marks every existing dependent dirty as well, eagerly — the
  ///     recalc pass would discover them via BFS anyway, but eager marking
  ///     keeps the dirty set inspectable between mutations.
  ///
  /// Returns `kInvalidArgument` when `sheet_index >= sheet_count()`.
  Expected<void, Error> set_cell_value(std::size_t sheet_index, std::uint32_t row, std::uint32_t col, Value value);

  /// Stores a formula at `(row, col)` on `sheet_index` and registers its
  /// dependencies with the recalc engine.
  ///
  /// `formula` is the raw Excel-style formula text. Whether or not it begins
  /// with `=` is not checked here — the parser owns that contract — but
  /// callers typically pass the leading `=`. The cell's `cached_value` is
  /// reset to blank; a subsequent `recalc()` populates it.
  ///
  /// Side effects:
  ///   * Parses the formula in a transient arena to extract its
  ///     dependencies; the parser's diagnostics are NOT propagated. Parse
  ///     failures simply leave the cell with no outgoing dep edges (it
  ///     will surface `#NAME?` at evaluation time).
  ///   * Re-registers the cell with the dep graph and volatile tracker.
  ///   * Marks the cell and its existing dependents dirty.
  ///
  /// Returns `kInvalidArgument` when `sheet_index >= sheet_count()`.
  Expected<void, Error> set_cell_formula(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                         std::string formula);

  /// Drives a full incremental recalc using the embedded recalc engine.
  ///
  /// `registry` supplies the function dispatch table — typically
  /// `eval::default_registry()`. Returns the engine's `RecalcStats` on
  /// success.
  Expected<eval::RecalcStats, Error> recalc(const eval::FunctionRegistry& registry);

  /// Returns the embedded recalc engine. Exposed for tests and debug
  /// tooling; lifetime is tied to the workbook.
  const eval::RecalcEngine& recalc_engine() const noexcept { return *engine_; }

  /// Mutable access to the embedded recalc engine. Used by the parallel
  /// scheduler entry point in `eval/scheduler.{h,cpp}` to drive the same
  /// dirty-set / dep-graph state the single-threaded `recalc()` consumes.
  /// External callers should prefer `recalc()` / `recalc_parallel()`; the
  /// raw engine handle is provided for the few code paths (the scheduler,
  /// targeted tests) that need direct access.
  eval::RecalcEngine& recalc_engine() noexcept { return *engine_; }

  /// Drives a parallel incremental recalc using the embedded engine.
  ///
  /// Wrapper around `eval::recalc_parallel(*this, registry, cfg, stats)`.
  /// Callers MUST NOT race two `recalc_parallel` invocations on the same
  /// workbook — the scheduler does not own a workbook-level lock.
  ///
  /// `cfg` and `stats` are forward-declared types; callers wanting the
  /// default config / no-stats shape should pass an explicit
  /// `eval::SchedulerConfig{}` (header users that include
  /// `eval/scheduler.h` for the full definitions can construct it
  /// inline).
  Expected<void, Error> recalc_parallel(const eval::FunctionRegistry& registry, const eval::SchedulerConfig& cfg,
                                        eval::SchedulerStats* stats);

  /// Sets iterative-calc options for the workbook's recalc engine.
  /// Defaults match Excel's out-of-the-box settings:
  /// `{ enabled = false, max_iterations = 100, max_change = 0.001 }`.
  /// Setting `enabled = true` opts in to circular-reference resolution;
  /// circular SCCs that previously surfaced `#REF!` will then be solved
  /// via fixed-point iteration up to `max_iterations` passes.
  void set_iterative_options(eval::IterativeOptions opts);

  /// Returns the active iterative-calc options.
  const eval::IterativeOptions& iterative_options() const noexcept;

  // ---------------------------------------------------------------------------
  // Passive round-trip metadata (Bundle 2.4)
  // ---------------------------------------------------------------------------
  //
  // Defined names and tables are preserved so the OOXML writer slice can
  // emit them back unchanged. Neither participates in the dep graph at
  // this layer; named-range / structured-reference resolution at
  // evaluation time arrives in Phase 4.

  /// Read-only access to the workbook's defined-name list (in source
  /// declaration order).
  const std::vector<io::DefinedName>& defined_names() const noexcept { return defined_names_; }

  /// Replaces the workbook's defined-name list. Move-assigns to keep
  /// the I/O hand-off allocation-free for large name lists.
  void set_defined_names(std::vector<io::DefinedName> names) { defined_names_ = std::move(names); }

  /// Read-only access to the workbook's table-metadata list (in
  /// archive-discovery order, which matches the per-sheet rels walk).
  const std::vector<io::TableMetadata>& tables() const noexcept { return tables_; }

  /// Replaces the workbook's table-metadata list. Move-assigns for the
  /// same reason as `set_defined_names`.
  void set_tables(std::vector<io::TableMetadata> tables) { tables_ = std::move(tables); }

  /// Read-only access to the verbatim Override-listed parts the reader
  /// did not consume. The writer emits each entry as-is, including its
  /// `<Override>` registration in `[Content_Types].xml`. Default-typed
  /// binary parts (images, OLE objects) are NOT represented here.
  const std::vector<io::PassthroughPart>& passthrough_parts() const noexcept { return passthrough_parts_; }

  /// Replaces the workbook's passthrough-part list. Move-assigns to
  /// keep the I/O hand-off allocation-free for archives carrying many
  /// preserved parts.
  void set_passthrough_parts(std::vector<io::PassthroughPart> parts) { passthrough_parts_ = std::move(parts); }

  // ---------------------------------------------------------------------------
  // Pivot caches
  // ---------------------------------------------------------------------------
  //
  // A workbook owns the pivot caches that its pivot tables reference; one
  // cache may back multiple pivot tables (multiple `<pivotTable>` parts can
  // share a `<pivotCacheDefinition>`). Caches are heap-owned via
  // `unique_ptr` so their addresses are stable and the OOXML reader can
  // hand pointers into the workbook's pivot evaluator without worrying
  // about vector reallocations.

  /// Read-only access to the workbook's pivot caches in
  /// document-discovery order.
  const std::vector<std::unique_ptr<pivot::PivotCache>>& pivot_caches() const noexcept { return pivot_caches_; }

  /// Appends a pivot cache to the workbook. Ownership transfers; no
  /// validation is performed (the OOXML reader is responsible for
  /// assigning unique cache ids).
  void add_pivot_cache(std::unique_ptr<pivot::PivotCache> cache);

  /// Looks up a pivot cache by id. Linear scan over `pivot_caches_`;
  /// workbooks typically have fewer than ten caches so the cost is
  /// negligible. Returns `nullptr` when no cache matches.
  const pivot::PivotCache* find_pivot_cache(std::uint32_t cache_id) const noexcept;

  /// Returns the OOXML workbook variant this instance round-trips as.
  /// Defaults to `WorkbookKind::kXlsx`; the reader updates it from the
  /// workbook part's content type, and the writer consults it when
  /// emitting `[Content_Types].xml`. The engine treats all four kinds
  /// identically for cell evaluation; only the package envelope (and,
  /// for `.xlsm` / `.xltm`, the captured `xl/vbaProject.bin` carried in
  /// `passthrough_parts()`) differs.
  io::WorkbookKind kind() const noexcept { return kind_; }

  /// Sets the workbook variant. Plain data — no lifecycle implications,
  /// no recalc-engine interaction. Callers using the writer to emit a
  /// macro-enabled variant must additionally ensure `xl/vbaProject.bin`
  /// is present in `passthrough_parts()`; the writer does not synthesise
  /// the part.
  void set_kind(io::WorkbookKind kind) noexcept { kind_ = kind; }

 private:
  Workbook();

  std::vector<Sheet> sheets_;
  // Embedded recalc engine. PIMPL-via-unique_ptr so callers can include
  // `workbook.h` without dragging the `eval/recalc_engine.h` header (and
  // its own transitive deps on the dep graph / arena / parser).
  std::unique_ptr<eval::RecalcEngine> engine_;
  // Passive OOXML metadata; populated by the reader and consumed by
  // the writer for round-trip preservation. Empty by default.
  std::vector<io::DefinedName> defined_names_;
  std::vector<io::TableMetadata> tables_;
  std::vector<io::PassthroughPart> passthrough_parts_;
  // Pivot caches owned by the workbook. One cache may be referenced by
  // multiple pivot tables (per `Sheet::pivot_tables()`).
  std::vector<std::unique_ptr<pivot::PivotCache>> pivot_caches_;
  // OOXML workbook variant. Defaults to plain `.xlsx`; the reader sets
  // this from `[Content_Types].xml` and the writer consults it when
  // emitting the workbook content-type Override. Plain data; no
  // lifecycle implications.
  io::WorkbookKind kind_ = io::WorkbookKind::kXlsx;
};

}  // namespace formulon

#endif  // FORMULON_WORKBOOK_H_
