//
// Top-level workbook model. The current surface owns a vector of `Sheet`
// instances and exposes a `save()` method that serialises the workbook to
// a `.xlsx` byte stream via the OOXML writer slice. Defined-name and
// table metadata are preserved here as passive round-trip state so the
// reader/writer pipeline can carry them through unchanged; named-range
// resolution at evaluation time and structured-reference parsing arrive
// in a follow-up. Styles, shared strings, pivots and the full cell store
// will be layered on in follow-up work.

#ifndef FORMULON_WORKBOOK_H_
#define FORMULON_WORKBOOK_H_

#include <cstddef>
#include <cstdint>
#include <deque>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "calc_settings.h"
#include "default_content_type.h"
#include "defined_name.h"
#include "excel_profile.h"
#include "external_link.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "styles.h"
#include "table.h"
#include "theme.h"
#include "unknown_relationship.h"
#include "utils/date_time.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {

// `WorkbookKind` is held by value but is an enum class with an explicit
// underlying type, so we can opaque-forward-declare it here. Concrete
// callers that need to name enumerators include `workbook_kind.h`
// directly. The default value is assigned in the out-of-line
// `Workbook()` constructor in workbook.cpp so the header itself never
// references an enumerator.
enum class WorkbookKind : std::uint8_t;
enum class WorkbookFormat : std::uint8_t;

namespace eval {
class FunctionRegistry;
class RecalcEngine;
struct RecalcStats;
struct SchedulerConfig;
struct SchedulerStats;
struct SheetCellRange;
using IterativeProgressCb = bool (*)(std::uint32_t iteration, double max_residual, std::uint32_t max_iterations,
                                     void* user_data);
}  // namespace eval

namespace pivot {
class PivotCache;
}  // namespace pivot

namespace parser {
class AstNode;
struct ExternalBookIndexer;
}  // namespace parser

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
  /// The style table is seeded with Excel's reserved defaults exactly as
  /// `create()` seeds it, so a style record appended through any mutator
  /// lands after the reserved slots whichever factory built the workbook.
  /// A reader reconstructing a workbook from a file overrides that with
  /// `set_styles`, since the file's style table — including its absence —
  /// is the authoritative one.
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

  /// Hard ceiling on `sheet_count()`.
  ///
  /// `eval::CellNodeId` addresses a sheet with a 16-bit id, so a sheet at
  /// index 65536 would alias sheet 0 in the dependency graph and silently
  /// register and dirty the wrong cells. Every append path enforces this
  /// bound, which is what makes each `static_cast<std::uint16_t>` of a
  /// sheet index lossless. Excel itself caps sheets at available memory
  /// and no real workbook approaches this number.
  static constexpr std::size_t kMaxSheets = 0xFFFFU;

  /// Number of sheets in the workbook. Always at least 1 for
  /// `create()`-constructed instances, never more than `kMaxSheets`.
  std::size_t sheet_count() const noexcept { return sheets_.size(); }

  /// Estimated heap bytes this workbook holds.
  ///
  /// Exists for hosts whose garbage collector only sees the handle and
  /// not what hangs off it — a JS engine will happily keep thousands of
  /// multi-megabyte workbooks alive because each one looks like a
  /// pointer-sized object — so they can report the real weight and let
  /// collection pressure track it.
  ///
  /// Counted: the cell store (materialised slots plus each cell's formula
  /// text, owned cached text and phonetic text), the shared-string
  /// storage every text value borrows from, the passthrough part payloads
  /// (which embedded media dominates), and the workbook-level round-trip
  /// metadata strings. Not counted: allocator bookkeeping, the recalc
  /// engine's dependency graph and arenas, styles, and the pivot caches —
  /// all of which are bounded by the counted parts in practice.
  ///
  /// The result is an estimate for pressure reporting, not an accounting
  /// figure: it walks every materialised cell, so it is O(cells) and
  /// belongs at coarse boundaries (after a load, after a recalc) rather
  /// than in a loop.
  std::size_t approximate_memory_bytes() const noexcept;

  /// Immutable access to the sheet at `index`. `index` must be `<
  /// sheet_count()`.
  const Sheet& sheet(std::size_t index) const { return sheets_[index]; }

  /// Mutable access to the sheet at `index`. `index` must be `<
  /// sheet_count()`.
  Sheet& sheet(std::size_t index) { return sheets_[index]; }

  /// Appends a new sheet with display name `name` and returns its index.
  ///
  /// An index rather than a reference, because `sheets_` is a vector: a
  /// reference handed out here would be dangling as soon as the next
  /// append reallocated, and the two-appends-then-use ordering that
  /// triggers it reads as obviously correct. Callers that want to mutate
  /// the new sheet go through `sheet(index)`, whose result has the same
  /// lifetime as any other element reference and is re-fetched after a
  /// structural mutation.
  ///
  /// Returns `kMaxSheets` — never a valid index, since indices run
  /// `0 .. kMaxSheets - 1` — when the workbook is already at the ceiling,
  /// leaving it unchanged. This overload has no error channel, so the
  /// out-of-range sentinel is what keeps a refused append from reading as
  /// a successful one.
  ///
  /// Duplicate names are not rejected at this layer, deliberately: this is
  /// the entry point that can build a workbook Excel would refuse, which
  /// is what the reader-rejection tests need to author their fixtures.
  /// Anything appending from outside the process must not use it. Callers
  /// that append from external input use `add_sheet_checked`; public API
  /// callers that also want name validation use `add_sheet_validated`,
  /// which is the path every binding and both file readers take.
  std::size_t add_sheet(std::string name);

  /// Appends a new sheet with display name `name` and returns its index,
  /// rejecting the append once the workbook holds `kMaxSheets` sheets.
  ///
  /// The name itself is taken verbatim, so this is the entry point for
  /// callers that must preserve whatever name their source carried while
  /// still refusing a sheet count the dependency graph cannot address.
  /// The file readers do not use it: a name arriving from a file is
  /// untrusted input and goes through `add_sheet_validated`, because a
  /// duplicate name silently resolves every lookup to the first match.
  ///
  /// Errors:
  ///   * `kSheetCountLimitExceeded` when `sheet_count() == kMaxSheets`.
  Expected<std::size_t, Error> add_sheet_checked(std::string name);

  /// Appends a new sheet after validating `name` the same way
  /// `rename_sheet` does: strict UTF-8, non-empty, at most 31 UTF-16 code units, no
  /// forbidden character (`: \ / ? * [ ]`), and no Unicode-simple-fold
  /// collision with an existing sheet. Returns a pointer to the new sheet
  /// on success (owned by the Workbook, invalidated by later structural
  /// mutations) or `kInvalidSheetName` on any violation.
  ///
  /// This is also the append path the OOXML and XLSB readers use, so the
  /// name validation `add_sheet` defers to the I/O boundary is in fact
  /// enforced there. After a successful load `sheet_by_name` and
  /// `sheet_index_by_name` therefore resolve unambiguously: no two sheets
  /// share a case-folded name.
  ///
  /// Errors:
  ///   * `kInvalidSheetName` on any name violation.
  ///   * `kSheetCountLimitExceeded` when `sheet_count() == kMaxSheets`.
  Expected<Sheet*, Error> add_sheet_validated(std::string name);

  /// Renames the sheet at `index` to `new_name`.
  ///
  /// Updates the sheet's stored name and rewrites every affected local
  /// reference in cell formulas, defined names, conditional-format and
  /// validation formulas, hyperlinks, table formulas, and pivot sources.
  /// Unrelated literals, unresolved references, and external-workbook
  /// references remain byte-identical.
  ///
  /// Errors:
  ///   * `kSheetIndexOutOfRange` when `index >= sheet_count()`.
  ///   * `kInvalidSheetName` when `new_name` is malformed UTF-8, empty, exceeds 31
  ///     UTF-16 units, contains a forbidden character (`: \ / ? * [ ]`),
  ///     or collides under Unicode simple case folding with another sheet's
  ///     name. A no-op rename to the sheet's existing name under the same
  ///     fold is accepted: the sheet's stored name is updated to the
  ///     new casing, no other state changes.
  Expected<void, Error> rename_sheet(std::uint32_t index, std::string new_name);

  /// Removes the sheet at `index`.
  ///
  /// References to the removed sheet are rewritten to `#REF!` at the
  /// affected AST subtree, preserving partial expressions and literals.
  /// Sheet-scoped names and tables owned by the removed sheet are dropped;
  /// surviving scopes and table indices are remapped. Pivot caches whose
  /// real worksheet source was removed, and pivot tables bound to those
  /// caches, are dropped. One final dependency-graph rebuild runs against
  /// the final workbook topology.
  ///
  /// Errors:
  ///   * `kSheetIndexOutOfRange` when `index >= sheet_count()`.
  ///   * `kCannotRemoveLastSheet` when `sheet_count() == 1` (Excel's
  ///     UI rejects the same op; we mirror it so `save()` never lands
  ///     in the empty-sheet-list state Excel refuses to open).
  Expected<void, Error> remove_sheet(std::uint32_t index);

  /// Moves the sheet at `from_index` to `to_index`.
  ///
  /// `to_index` is interpreted in the post-removal sheet list (matches
  /// Excel's UI semantics). For example, with three sheets, moving
  /// sheet 0 to the end uses `to_index == 2`, not `3`. A move to the
  /// same position is a successful no-op. Cell formulas and defined
  /// names continue to reference sheets by name, so observers reading
  /// `sheet(i)` see the new positional layout while cross-sheet
  /// formulas remain correct.
  ///
  /// Errors:
  ///   * `kSheetIndexOutOfRange` when either `from_index >=
  ///     sheet_count()` or `to_index >= sheet_count()`.
  Expected<void, Error> move_sheet(std::uint32_t from_index, std::uint32_t to_index);

  /// Unicode-simple-fold lookup by display name. `"SHEET2"` and
  /// `"Sheet2"` locate the same sheet, as do Unicode simple-fold pairs such
  /// as `Ä` and `ä`. This is locale-independent simple folding only; it does
  /// not claim full Excel-equivalence, normalization, or locale tailoring.
  /// Returns `nullptr` when no sheet matches. A linear scan is used;
  /// workbooks typically carry O(1)–O(10) sheets, so a hash index is
  /// premature.
  const Sheet* sheet_by_name(std::string_view name) const noexcept;

  /// Returns the 0-based index of the sheet whose display name matches
  /// `name` under Unicode simple case folding. Returns
  /// `static_cast<size_t>(-1)` when no sheet matches. Linear scan, like
  /// `sheet_by_name`.
  std::size_t sheet_index_by_name(std::string_view name) const noexcept;

  /// Serialises the workbook to an in-memory `.xlsx` byte stream. Delegates
  /// to `io::write_ooxml`; see that function's documentation for the exact
  /// set of OOXML parts emitted by the empty-workbook writer slice.
  /// Equivalent to `save_as(WorkbookFormat::Ooxml)`.
  Expected<std::vector<std::uint8_t>, Error> save() const;

  /// Serialises the workbook using an explicit container `format`.
  /// `WorkbookFormat::Ooxml` delegates to `io::write_ooxml` (the
  /// `.xlsx` writer); `WorkbookFormat::Xlsb` delegates to
  /// `io::xlsb::write_xlsb` (the MS-XLSB writer). `WorkbookFormat::Unknown`
  /// is not a valid save target and returns `kInvalidArgument`.
  Expected<std::vector<std::uint8_t>, Error> save_as(WorkbookFormat format) const;

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

  /// Stores a text literal whose bytes are copied into the destination cell.
  /// This is appropriate for short-lived caller buffers; source readers that
  /// keep a workbook-scoped shared-string store can use `set_cell_value`
  /// with `Value::text` instead.
  Expected<void, Error> set_cell_text(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                      std::string_view text);

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

  /// Canonical formula-bar spelling of stored formula text: Excel's
  /// `_xlfn.` / `_xlws.` / `_xlpm.` storage prefixes stripped and the stored
  /// `SINGLE` / `ANCHORARRAY` calls read back as `@x` / `x#`. This is the
  /// normalization `set_cell_formula` applies, exposed so an entry check
  /// judges exactly the text that would be stored.
  static std::string normalize_formula_text(std::string formula);

  /// Model spelling of a formula read from a file: `normalize_formula_text`,
  /// then every `[N]` link index rewritten to the book name the formula bar
  /// shows (with the link's absolute directory when it has one), then
  /// `bind_external_books`. `[0]!Name` is kept as stored.
  std::string ingest_stored_formula(std::string_view stored);

  /// Spells every loaded formula without the dynamic-array mark as Excel 365
  /// shows it, with an `@` wherever the legacy formula intersects implicitly
  /// (`io::xlsb::legacy_intersections`), so it evaluates as Excel does. Run by
  /// each reader once cells, marks and defined names are in.
  void apply_legacy_implicit_intersections();

  /// Dirties every existing dependent of `(sheet_index, row, col)` via the
  /// engine's ordinary dep-graph edges, without touching the cell's own
  /// value, formula, or registration. Out-of-range coordinates are a no-op.
  ///
  /// Call after a mutation that changes something a formula can read about
  /// a cell without changing the cell's own `Value` -- currently the
  /// phonetic-run/phonetic-reading setters, which `PHONETIC()` reads
  /// through the same reference edge any other function argument creates,
  /// so no dep-graph re-registration is needed here, only the dirty mark.
  void mark_cell_dependents_dirty(std::size_t sheet_index, std::uint32_t row, std::uint32_t col);

  /// Adds a normalized merge range through the workbook mutation path.  The
  /// engine mutex is held while the sheet merge vector is updated so a spill
  /// collision check cannot observe a partially-written metadata mutation.
  Expected<void, Error> add_merge(std::size_t sheet_index, MergeRange merge);

  /// Removes every merge intersecting `merge` and wakes any blocked spill
  /// anchor whose attempted footprint intersects the removed range.
  Expected<void, Error> remove_merges_intersecting(std::size_t sheet_index, MergeRange merge);

  /// Removes one merge by insertion order and wakes blocked spill anchors
  /// intersecting the removed rectangle.
  Expected<void, Error> remove_merge_at(std::size_t sheet_index, std::size_t index);

  /// Removes all merges from a sheet and wakes all of that sheet's blocked
  /// spill anchors.
  Expected<void, Error> clear_merges(std::size_t sheet_index);

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
  ///
  /// `max_iterations` is clamped to `kMaxIterationsCap`, so a value
  /// above it reads back as the cap. The clamp lives here because this is
  /// where the budget enters the model: an iteration loop elsewhere in the
  /// engine inherits the bound instead of restating it, and a caller
  /// cannot hand the engine a count that no wall-clock limit or
  /// cancellation hook would ever stop. The cap is Excel's own dialog
  /// limit, so it costs no fidelity. The lower end is not clamped here:
  /// the solver documents `0` as meaning one pass.
  void set_iterative_options(IterativeOptions opts);

  /// Returns the active iterative-calc options. `max_iterations` is always
  /// within `kMaxIterationsCap` when the options were installed
  /// through `set_iterative_options`.
  const IterativeOptions& iterative_options() const noexcept;

  // --- Recalc ---
  // Pass-through wrappers that surface a few `RecalcEngine` knobs from
  // the workbook handle. Kept intentionally trivial so the workbook
  // header does not depend on `eval/recalc_engine.h` for the type body.

  /// Drives an incremental recalc bounded by `viewport`. Forwards to
  /// `RecalcEngine::partial_recalc`; see that method for the closure
  /// semantics.
  Expected<eval::RecalcStats, Error> partial_recalc(const eval::FunctionRegistry& registry,
                                                    const eval::SheetCellRange& viewport);

  /// Installs / clears the iterative-solver progress callback. Forwards
  /// to `RecalcEngine::set_iterative_progress`. Pass `nullptr` to clear.
  void set_iterative_progress(eval::IterativeProgressCb cb, void* user_data) noexcept;

  /// Dirties every formula cell in the workbook whose text mentions
  /// SUBTOTAL or AGGREGATE (case-insensitive), without touching dep-graph
  /// edges. Call after a row's hidden flag actually changes: both
  /// functions can read row-visibility state that no ordinary dep-graph
  /// edge models, so the edit cannot reach them through the dirty
  /// propagation an ordinary cell-value write triggers.
  void mark_row_visibility_dependents_dirty();

  /// Dirties every formula cell in the workbook, without touching dep-graph
  /// edges or spill geometry. Call after a workbook-wide setting that
  /// changes how a formula evaluates without changing any cell's own text
  /// or the dep graph's shape -- currently the Excel-compatibility profile
  /// switch, whose reach (SUMIF/COUNTIF criteria matching, INFO/CELL, and
  /// other host-dependent branches) is too broad to scope the way a
  /// defined-name or table edit can.
  void mark_all_formulas_dirty();

  // ---------------------------------------------------------------------------
  // Defined names and table metadata
  // ---------------------------------------------------------------------------
  //
  // Defined names and tables preserve their round-trip metadata and supply
  // named-range and structured-reference dependencies to the recalc engine.

  /// Read-only access to the workbook's defined-name list (in source
  /// declaration order).
  const std::vector<DefinedName>& defined_names() const noexcept { return defined_names_; }

  /// Replaces the workbook's defined-name list and refreshes dependencies
  /// of formulas using changed definitions. Metadata-only changes preserve
  /// cached values and spill geometry. Formula validation remains deferred
  /// to evaluation, matching the I/O readers' raw name-list hand-off.
  void set_defined_names(std::vector<DefinedName> names);

  /// Sets the formula text of the workbook-scoped defined name `name`,
  /// or appends it if it does not exist. An empty `formula` removes
  /// the entry instead. This is a compatibility wrapper around
  /// `set_defined_name_scoped(name, formula, -1)`.
  ///
  /// Returns `kOk` on success. Validation of the formula text is
  /// deliberately deferred to evaluation time (matching how the I/O
  /// reader carries formulas through unchanged), so this function
  /// never surfaces a parser error.
  Expected<void, Error> set_defined_name(std::string name, std::string formula);

  /// Sets the formula text of a defined name in `local_sheet_id` scope,
  /// or appends it if it does not exist. Pass `-1` for workbook scope;
  /// pass a 0-based sheet index for sheet scope. An empty formula
  /// removes the matching entry. Lookups are case-insensitive within
  /// the same scope only, matching Excel's name-resolution model.
  ///
  /// Returns `kInvalidArgument` when `name` is empty or
  /// `local_sheet_id` names a sheet outside the workbook.
  Expected<void, Error> set_defined_name_scoped(std::string name, std::string formula, std::int32_t local_sheet_id);

  /// Sets the hidden flag of the defined name `name` in `local_sheet_id`
  /// scope (`-1` for workbook scope). Returns `kInvalidArgument` when no
  /// such name exists.
  Expected<void, Error> set_defined_name_hidden(std::string_view name, std::int32_t local_sheet_id, bool hidden);

  // --- AutoFilter ---
  // Every change below dirties row-visibility dependents, since whether a
  // sheet carries filter criteria decides which hidden rows are filtered.

  /// Replaces the sheet AutoFilter with `filter` and creates or updates the
  /// sheet's hidden `_xlnm._FilterDatabase` name over its range. Returns
  /// `kAutoFilterInvalid` when `validate_auto_filter` rejects the model and
  /// `kInvalidArgument` for a bad sheet index.
  Expected<void, Error> set_sheet_auto_filter(std::size_t sheet_index, AutoFilter filter);

  /// Removes the sheet AutoFilter and its `_xlnm._FilterDatabase` name.
  Expected<void, Error> remove_sheet_auto_filter(std::size_t sheet_index);

  /// XML view of `set_sheet_auto_filter`: empty removes the AutoFilter, and
  /// an element the model cannot represent is stored opaque without touching
  /// `_FilterDatabase`.
  Expected<void, Error> set_sheet_auto_filter_xml(std::size_t sheet_index, std::string_view xml);

  /// Replaces the AutoFilter of table `table_index`. Tables carry no
  /// `_FilterDatabase` name. Returns `kAutoFilterInvalid` or
  /// `kInvalidArgument` as `set_sheet_auto_filter` does.
  Expected<void, Error> set_table_auto_filter(std::size_t table_index, AutoFilter filter);

  /// Removes the AutoFilter of table `table_index`.
  Expected<void, Error> remove_table_auto_filter(std::size_t table_index);

  /// Read-only access to the workbook's table-metadata list (in
  /// archive-discovery order, which matches the per-sheet rels walk).
  const std::vector<TableMetadata>& tables() const noexcept { return tables_; }

  /// Mutable access for structural edits that update a table's owning sheet
  /// and A1 rectangle before the writer rebuilds its relationships.
  std::vector<TableMetadata>& mutable_tables() noexcept { return tables_; }

  /// Replaces the workbook's table-metadata list. Move-assigns for the
  /// same reason as `set_defined_names`.
  void set_tables(std::vector<TableMetadata> tables) { tables_ = std::move(tables); }

  /// Re-registers dep-graph edges and dirties every formula cell whose
  /// structured reference names one of `table_names` (typically a table's
  /// `name` and `display_name`). Call after a `mutable_tables()` edit
  /// (create, ref/column change, remove) so structured references' dep-
  /// graph edges and cached values reflect the current table metadata --
  /// mirrors `set_defined_name_scoped`'s scoped reindex, since a
  /// `StructuredRef` resolves to a static rectangle at registration time
  /// the same way a `NameRef` does. An empty `table_names` is a no-op.
  void reindex_formulas_for_table_change(const std::vector<std::string>& table_names);

  /// Read-only access to the verbatim parts the reader did not model.
  /// The writer emits each entry as-is. `<Override>`-listed parts
  /// replicate their `<Override>` registration in `[Content_Types].xml`;
  /// Default-typed binary/media parts (vbaProject.bin, images, drawings,
  /// VML, and their rels) rely on the round-tripped `<Default>` entries
  /// exposed via `default_content_types()`.
  const std::vector<PassthroughPart>& passthrough_parts() const noexcept { return passthrough_parts_; }

  /// Replaces the workbook's passthrough-part list. Move-assigns to
  /// keep the I/O hand-off allocation-free for archives carrying many
  /// preserved parts.
  void set_passthrough_parts(std::vector<PassthroughPart> parts) { passthrough_parts_ = std::move(parts); }

  /// Adds one passthrough part, keeping the list ordered by path. Returns
  /// `kInvalidArgument` when a part with the same path is already present or
  /// the path is empty.
  Expected<void, Error> add_passthrough_part(PassthroughPart part);

  /// Replaces the bytes of the passthrough part at `path`, leaving its
  /// content type untouched. Returns `kInvalidArgument` when no such part
  /// exists.
  Expected<void, Error> replace_passthrough_part(std::string_view path, std::vector<std::uint8_t> bytes);

  /// Registers a workbook-level relationship of `type` to the package part
  /// `target` (full package path, e.g. `xl/theme/theme1.xml`). A relationship
  /// with the same type and target is not added twice. The writer mints the
  /// rId, and omits the relationship if `target` is not an emitted part.
  void add_workbook_relationship(std::string type, std::string target);

  /// Parses the theme part (a passthrough part) on every call, never cached.
  /// Never fails: an absent or unparseable part yields `default_theme()` with
  /// the matching `ThemeSource`.
  LoadedTheme load_theme() const;

  /// Sets the 12 scheme colours, patching only `a:clrScheme`. When the
  /// workbook has no theme part, a minimal valid one is generated first,
  /// registered with its content type and the workbook relationship. Returns
  /// the parse error when an existing part is not a parseable theme (the part
  /// is then left untouched).
  Expected<void, Error> set_theme_colors(const ThemeColors& colors);

  /// Sets the four theme fonts, patching only `a:fontScheme`; same generation
  /// and failure rules as `set_theme_colors`.
  Expected<void, Error> set_theme_fonts(const ThemeFonts& fonts);

  // ---------------------------------------------------------------------------
  // Threaded comments and persons
  // ---------------------------------------------------------------------------

  /// The workbook's person list (`xl/persons/person.xml`) in file order.
  const std::vector<Person>& persons() const noexcept { return persons_; }

  /// Mutable person list for the reader; edits go through `add_person` /
  /// `remove_person`.
  std::vector<Person>& mutable_persons() noexcept { return persons_; }

  /// Appends `person`. Fails with `kThreadedCommentInvalid` when the id is
  /// not a brace-wrapped uppercase GUID, is already listed, or the display
  /// name is empty.
  Expected<void, Error> add_person(Person person);

  /// Removes the person with `id`. Fails with `kThreadedCommentInvalid` when
  /// no such person exists or a threaded comment or mention still names them.
  Expected<void, Error> remove_person(std::string_view id);

  /// Adds `comment` to sheet `sheet`. An empty `parent_id` opens a new thread
  /// at `row`/`col`, which must hold neither a note nor a thread; a non-empty
  /// one appends a reply to that thread, taking its anchor and ignoring
  /// `row`/`col` and `done`. Ids, timestamps, persons and mention ranges are
  /// validated; any violation fails with `kThreadedCommentInvalid` and leaves
  /// the sheet unchanged.
  Expected<void, Error> add_threaded_comment(std::size_t sheet, ThreadedComment comment);

  /// Replaces the text and mentions of the comment with `id`.
  Expected<void, Error> edit_threaded_comment(std::size_t sheet, std::string_view id, std::string text,
                                              std::vector<Mention> mentions);

  /// Sets the resolved state of the thread opened by `thread_id`, which must
  /// name an opening comment, not a reply.
  Expected<void, Error> set_thread_resolved(std::size_t sheet, std::string_view thread_id, bool done);

  /// Removes the comment with `id`. Removing an opening comment removes its
  /// whole thread; removing a reply removes that reply only.
  Expected<void, Error> remove_threaded_comment(std::size_t sheet, std::string_view id);

  /// Read-only access to the verbatim workbook-level `<Relationship>`
  /// entries (from `xl/_rels/workbook.xml.rels`) whose Type URI the
  /// reader did not recognise. Captured so the writer can re-emit
  /// every entry, keeping passthrough-listed parts (theme, calcChain,
  /// vbaProject, customXml, ...) reachable through the relationship
  /// graph. The writer mints fresh rIds; the original `id` is
  /// preserved on the struct for diagnostics only.
  const std::vector<UnknownRelationship>& unknown_workbook_rels() const noexcept { return unknown_workbook_rels_; }

  /// Replaces the workbook's unknown-relationship list. Move-assigns
  /// to keep the I/O hand-off allocation-free.
  void set_unknown_workbook_rels(std::vector<UnknownRelationship> rels) { unknown_workbook_rels_ = std::move(rels); }

  /// Read-only access to unrecognised package-root relationships from
  /// `_rels/.rels` (for example thumbnails and digital-signature origins).
  /// Their targets reside in `passthrough_parts()`; preserving the edge keeps
  /// those parts reachable after an OOXML read-modify-write cycle.
  const std::vector<UnknownRelationship>& unknown_package_rels() const noexcept { return unknown_package_rels_; }

  /// Replaces the captured package-root relationship list.
  void set_unknown_package_rels(std::vector<UnknownRelationship> rels) { unknown_package_rels_ = std::move(rels); }

  /// Read-only access to the `<Default>` content-type registrations the
  /// reader captured from `[Content_Types].xml`. The writer re-emits the
  /// entries whose extension is used by a Default-typed passthrough part
  /// (vbaProject.bin, images, VML, ...) so those parts keep a resolvable
  /// content type in the round-tripped package.
  const std::vector<DefaultContentType>& default_content_types() const noexcept { return default_content_types_; }

  /// Replaces the workbook's captured `<Default>` content-type list.
  /// Move-assigns to keep the I/O hand-off allocation-free.
  void set_default_content_types(std::vector<DefaultContentType> defaults) {
    default_content_types_ = std::move(defaults);
  }

  /// Read-only access to the workbook's external-link list (in
  /// `<externalReferences>` document order). Each entry surfaces the
  /// relationship metadata for one cross-workbook reference; the body
  /// part itself round-trips through `passthrough_parts()` unchanged.
  const std::vector<ExternalLinkRecord>& external_links() const noexcept { return external_links_; }

  /// The external links ordered by `index` (document order for equal
  /// indices). `skip_ole_dde` drops OLE and DDE links, which hold no
  /// supporting workbook.
  std::vector<const ExternalLinkRecord*> external_links_by_index(bool skip_ole_dde) const;

  /// Replaces the workbook's external-link list. Move-assigns to keep
  /// the I/O hand-off allocation-free.
  void set_external_links(std::vector<ExternalLinkRecord> links) { external_links_ = std::move(links); }

  /// Ensures every external book `root` names has a link: a book no link
  /// matches gets a new record appended (index `size()+1`), and a sheet the
  /// matched link does not list is appended to its sheet names. `[0]!Name`
  /// is this workbook and binds nothing. Idempotent.
  void bind_external_books(const parser::AstNode& root);

  /// Text form of `bind_external_books`: parses `formula` (with or without
  /// a leading `=`) and binds it; an unparseable formula binds nothing.
  void bind_external_books(std::string_view formula);

  /// The link a reference spelled with directory `path` (empty for none,
  /// otherwise with its trailing separator) and file `book` names, under
  /// ASCII case folding. A path matches a link's absolute path, falling back
  /// to a pathless link of the same file; a bare book matches the file name
  /// of the lowest-index link. Null when none matches.
  const ExternalLinkRecord* find_external_link(std::string_view path, std::string_view book) const noexcept;

  /// The `[N]` mapping a storage formatter writes cross-workbook references
  /// with, backed by `find_external_link`. Valid while this workbook lives.
  parser::ExternalBookIndexer external_book_indexer() const noexcept;

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

  /// Mutable access for workbook-structural edits that must keep a pivot
  /// cache's worksheet source attached to its renamed sheet.
  std::vector<std::unique_ptr<pivot::PivotCache>>& mutable_pivot_caches() noexcept { return pivot_caches_; }

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
  WorkbookKind kind() const noexcept { return kind_; }

  /// Sets the workbook variant. Plain data — no lifecycle implications,
  /// no recalc-engine interaction. Callers using the writer to emit a
  /// macro-enabled variant must additionally ensure `xl/vbaProject.bin`
  /// is present in `passthrough_parts()`; the writer does not synthesise
  /// the part.
  void set_kind(WorkbookKind kind) noexcept { kind_ = kind; }

  // ---------------------------------------------------------------------------
  // Calculation mode (workbook-level `<calcPr>` policy)
  // ---------------------------------------------------------------------------
  //
  // Excel's "Calculation options" workbook setting. `kAuto` is the
  // default; `kManual` suppresses automatic recalc on input, leaving
  // it to the host UI to drive `recalc()` explicitly; `kAutoNoTable`
  // recalcs everything except data-table cells. The engine itself
  // does not gate evaluation on this setting (every `recalc()` call
  // honours all dirty cells); it is preserved as round-trip metadata
  // and surfaced through the bindings so the host UI can mirror
  // Excel's user-visible state.

  /// Workbook-level calc-mode enum. Type-aliased from `CalcMode`
  /// (declared in `calc_settings.h`) so the C ABI, the OOXML
  /// reader/writer and the bindings can all reference the enum
  /// without pulling in the full `workbook.h`. Existing source that
  /// names `Workbook::CalcMode::kAuto` keeps compiling unchanged.
  using CalcMode = formulon::CalcMode;

  /// Returns the workbook-level calc mode. Defaults to `kAuto`.
  CalcMode calc_mode() const noexcept { return calc_mode_; }

  /// Sets the workbook-level calc mode. Plain metadata — does not
  /// affect the recalc engine's evaluation policy.
  void set_calc_mode(CalcMode mode) noexcept { calc_mode_ = mode; }

  // ---------------------------------------------------------------------------
  // Date system (`<workbookPr date1904>`)
  // ---------------------------------------------------------------------------
  //
  // Excel supports two date-serial epochs. The 1900 system (serial 1 =
  // 1900-01-01, with Lotus's fictitious 1900-02-29) is the default; the
  // 1904 system (serial 0 = 1904-01-01, no ghost day) is common in
  // Mac-authored workbooks. A serial in the 1904 system is 1462 less than
  // the 1900 serial for the same calendar day, so losing this flag on a
  // save shifts every date four years. The flag is parsed from
  // `<workbookPr date1904>` and synchronised with the retained raw
  // `<workbookPr>` attributes when the writer emits the part.

  /// True when the workbook uses the 1904 date system. Defaults to false
  /// (1900 system). Consumed by date-serial conversions
  /// (`date_time::serial_from_ymd` / `ymd_from_serial`).
  bool date1904() const noexcept { return date1904_; }

  /// Sets the 1904-date-system flag. The writer synchronises the retained
  /// `<workbookPr>` fragment's `date1904` / legacy `1904` attributes with
  /// this model value while preserving the other raw attributes.
  void set_date1904(bool value);

  // ---------------------------------------------------------------------------
  // Clock seam
  // ---------------------------------------------------------------------------
  //
  // Some results depend on when they are computed: `NOW` / `TODAY`, and the
  // pivot relative-period filters that ask for "this month" or "year to
  // date". Left to themselves each would read the host clock independently,
  // which makes a recalc internally inconsistent across a midnight boundary
  // and makes every such result untestable.
  //
  // Pinning a reading here gives the whole workbook one instant to agree on.
  // This is model state, not file state: nothing in the OOXML package
  // records it, and a save drops it.

  /// The pinned wall-clock reading, or `std::nullopt` when the workbook
  /// follows the host clock (the default). Threaded into `EvalContext` at
  /// the evaluator boundary and into the pivot filter engine.
  const std::optional<date_time::CivilTime>& pinned_now() const noexcept { return pinned_now_; }

  /// Pins every clock-dependent result to `value`. Intended for tests and
  /// for hosts that need a reproducible recalc; production callers leave it
  /// unset so the host clock shows through.
  void set_pinned_now(date_time::CivilTime value);

  /// Releases the pin so clock-dependent results follow the host clock again.
  void clear_pinned_now();

  // ---------------------------------------------------------------------------
  // Workbook-level element round-trip (`<workbookPr>` / `<workbookProtection>`
  // / `<bookViews>`)
  // ---------------------------------------------------------------------------
  //
  // These three workbook.xml elements are captured as raw XML (the same
  // pattern `SheetPrintSettings` uses for `<pageSetup>` / `<pageMargins>`)
  // so the writer can re-emit them verbatim in the correct ECMA-376
  // element order. Without this, tab-selection state (`bookViews` →
  // `activeTab`), workbook protection, and — most importantly — the
  // `date1904` flag are silently dropped on save.

  /// Raw `<workbookPr>` element captured from workbook.xml, or empty.
  const std::string& workbook_pr_xml() const noexcept { return workbook_pr_xml_; }
  void set_workbook_pr_xml(std::string xml) { workbook_pr_xml_ = std::move(xml); }

  /// Raw `<bookViews>` element captured from workbook.xml, or empty.
  const std::string& book_views_xml() const noexcept { return book_views_xml_; }
  void set_book_views_xml(std::string xml) { book_views_xml_ = std::move(xml); }

  /// Raw `<workbookProtection>` element captured from workbook.xml, or
  /// empty.
  const std::string& workbook_protection_xml() const noexcept { return workbook_protection_xml_; }
  void set_workbook_protection_xml(std::string xml) { workbook_protection_xml_ = std::move(xml); }

  /// Raw `<fileVersion>` / `<fileSharing>` and trailing `<extLst>` elements
  /// captured from workbook.xml. They are passive round-trip metadata.
  const std::string& file_version_xml() const noexcept { return file_version_xml_; }
  void set_file_version_xml(std::string xml) { file_version_xml_ = std::move(xml); }
  const std::string& file_sharing_xml() const noexcept { return file_sharing_xml_; }
  void set_file_sharing_xml(std::string xml) { file_sharing_xml_ = std::move(xml); }
  const std::string& workbook_ext_lst_xml() const noexcept { return workbook_ext_lst_xml_; }
  void set_workbook_ext_lst_xml(std::string xml) { workbook_ext_lst_xml_ = std::move(xml); }

  /// Extra namespace declarations (and `mc:Ignorable`) captured from the
  /// source `<workbook>` root element, serialised as ` name="value"`
  /// attribute pairs (leading space included), excluding the default
  /// `xmlns` and `xmlns:r` the writer always emits. Re-emitted on the
  /// writer's `<workbook>` root so namespaced attributes carried inside the
  /// raw `<bookViews>` / `<workbookPr>` captures (e.g. `xr2:uid`) resolve
  /// to a declared prefix — without them Excel rejects the file as
  /// malformed XML.
  const std::string& workbook_root_extra_attrs() const noexcept { return workbook_root_extra_attrs_; }
  void set_workbook_root_extra_attrs(std::string attrs) { workbook_root_extra_attrs_ = std::move(attrs); }

  // ---------------------------------------------------------------------------
  // Excel host compatibility
  // ---------------------------------------------------------------------------
  //
  // Some formula behaviours are host-specific even for the same Microsoft 365
  // channel and locale. The default runtime profile tracks Windows Excel 365
  // ja-JP; oracle tests can override this per golden set.

  /// Returns the full formula-behaviour profile. Defaults to `win-365-ja_JP`.
  ExcelProfile excel_profile() const noexcept { return excel_profile_; }

  /// Sets the full formula-behaviour profile used by future recalc calls.
  void set_excel_profile(ExcelProfile profile);

  // ---------------------------------------------------------------------------
  // Styles
  // ---------------------------------------------------------------------------
  //
  // The workbook owns a single `StylesTable` populated by the OOXML
  // reader and consumed by the OOXML writer. Cells reference an entry
  // via `Cell::xf_index`. A fresh / `create()`-built workbook starts
  // with a default-constructed table; `read_styles` populates it
  // wholesale during package load.

  /// Read-only access to the workbook's styles table.
  const StylesTable& styles() const noexcept { return styles_; }

  /// Replaces the workbook's styles table. Move-assigns to keep the
  /// reader hand-off allocation-free.
  void set_styles(StylesTable styles) { styles_ = std::move(styles); }

  /// Mutable access to the workbook's styles table.
  ///
  /// The OOXML reader populates the table during package load via
  /// `set_styles`; mutators added through the C ABI (font / fill /
  /// border / num-fmt / xf inserts) reach the underlying records
  /// through this accessor instead of round-tripping the whole table.
  StylesTable& mutable_styles() noexcept { return styles_; }

  // ---------------------------------------------------------------------------
  // Text storage (workbook-lifetime backing for `Value::text` views)
  // ---------------------------------------------------------------------------
  //
  // `Value::text` is a non-owning `string_view`. Cells that carry text
  // payloads decoded by the OOXML / XLSB readers (inline strings, SST
  // entries, BrtCellSt records) need a workbook-lifetime backing store
  // so the views remain valid after the read result is destructed and
  // the workbook is moved out. `text_storage_` provides exactly that:
  // a pointer-stable `std::deque<std::string>` owned by the workbook.
  //
  // Thread safety: appends are NOT synchronised. The current readers
  // (`read_ooxml`, `read_xlsb`) drive a single-threaded decode pass per
  // workbook, so concurrent `intern_text` calls are not possible. If
  // future code introduces parallel decode it must serialise around
  // this store (or replace it with a workbook-scoped shared-string
  // pool, which is the planned long-term design).
  //
  // Duplicate de-duplication is intentionally NOT performed: that
  // optimisation belongs in the future `SharedStringPool` redesign and
  // would otherwise complicate the lifetime invariant for callers that
  // still hold views into earlier appends.

  /// Appends `text`'s bytes to the workbook's text-storage deque and
  /// returns a `string_view` aliasing the freshly-stored copy. The
  /// returned view is valid for the workbook's lifetime; it is
  /// preserved across move construction / move assignment of the
  /// workbook because `std::deque`'s element addresses are stable
  /// under appends (and because the deque's storage is itself moved,
  /// not copied, by the workbook's defaulted move operations).
  std::string_view intern_text(std::string_view text);

  /// Mutable access to the workbook's text-storage deque. Reader
  /// pipelines (sheet_reader, sst_reader, cell_parser, xlsb::reader)
  /// take a `std::deque<std::string>&` and append payloads directly so
  /// they can build `string_view`s aliasing the stored bytes without
  /// going through `intern_text` (which would copy from a `string_view`
  /// rather than emplace in place). Non-reader callers should prefer
  /// `intern_text`.
  std::deque<std::string>& mutable_text_storage() noexcept { return text_storage_; }

  /// Read-only access to the workbook's text-storage deque. Exposed
  /// for tests / tooling that need to introspect how many distinct
  /// text payloads the workbook owns.
  const std::deque<std::string>& text_storage() const noexcept { return text_storage_; }

  /// Persists the cell-level xf index without otherwise mutating cell
  /// state. The cell at `(row, col)` is created (as a default-blank
  /// literal) when it does not yet exist, mirroring the growth
  /// semantics of `set_cell_value`. Returns `kInvalidArgument` for an
  /// out-of-range `sheet_index`. Coexists with `set_cell_value` /
  /// `set_cell_formula`: those calls leave `xf_index` untouched, so the
  /// caller can layer a style update on top of a value or formula
  /// write in either order.
  Expected<void, Error> set_cell_xf_index(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                          std::uint32_t xf_index);

  // ---------------------------------------------------------------------------
  // Row / column structural edits
  // ---------------------------------------------------------------------------
  //
  // Insert and delete operations migrate cells in the affected sheet and
  // rewrite every reference in the workbook (cell formulas across all
  // sheets, defined names, integer-coordinate metadata) in lockstep with
  // the move. Cell formulas pointing into the touched range follow Excel's
  // rules: an insert pushes affected coordinates forward; a delete drops
  // references inside the deleted interval (collapsing them to `#REF!`)
  // and shifts trailing references back. Sheet-scoped metadata —
  // merges, hyperlinks, validation ranges, comment anchors — receives
  // the same shift on the affected sheet only; cross-sheet references
  // are rewritten via the AST-based reference transform.
  //
  // Ordinary 2-D range clamping follows Excel's behaviour: deleting only
  // part of a range shrinks the surviving interval, while deleting every
  // coordinate in a range produces `#REF!`. The same rule applies to cell
  // formulas, defined names, and conditional-format formulas. A Ref3D
  // sheet-span keeps its shared inner coordinates unchanged. All other Excel
  // semantics — propagation across sheets, defined-name updates,
  // out-of-bounds collapse for inserts that overflow the sheet — match
  // Excel today.

  /// Inserts `count` rows at `row` on `sheet_index`. Existing rows at
  /// `row` and beyond shift down by `count`; rows that would land past
  /// `Sheet::kMaxRows` are dropped (their cells are lost). Returns
  /// `kInvalidArgument` for an out-of-range sheet index, a row beyond
  /// the sheet bound, or `count == 0`.
  Expected<void, Error> insert_rows(std::size_t sheet_index, std::uint32_t row, std::uint32_t count);

  /// Deletes `count` rows starting at `row` on `sheet_index`. The deleted
  /// rows are dropped wholesale; rows past `row + count` shift up by
  /// `count`. Returns `kInvalidArgument` on the same error paths as
  /// `insert_rows`.
  Expected<void, Error> delete_rows(std::size_t sheet_index, std::uint32_t row, std::uint32_t count);

  /// Inserts `count` columns at `col` on `sheet_index`. Mirrors
  /// `insert_rows` along the column axis.
  Expected<void, Error> insert_cols(std::size_t sheet_index, std::uint32_t col, std::uint32_t count);

  /// Deletes `count` columns starting at `col` on `sheet_index`. Mirrors
  /// `delete_rows` along the column axis.
  Expected<void, Error> delete_cols(std::size_t sheet_index, std::uint32_t col, std::uint32_t count);

 private:
  Workbook();

  // `bind_external_books` for a caller already holding the engine mutex.
  void bind_external_books_unlocked(const parser::AstNode& root);
  void bind_external_books_unlocked(std::string_view formula);

  std::vector<Sheet> sheets_;
  // Embedded recalc engine. PIMPL-via-unique_ptr so callers can include
  // `workbook.h` without dragging the `eval/recalc_engine.h` header (and
  // its own transitive deps on the dep graph / arena / parser).
  std::unique_ptr<eval::RecalcEngine> engine_;
  // Passive OOXML metadata; populated by the reader and consumed by
  // the writer for round-trip preservation. Empty by default.
  std::vector<DefinedName> defined_names_;
  std::vector<TableMetadata> tables_;
  std::vector<PassthroughPart> passthrough_parts_;
  std::vector<ExternalLinkRecord> external_links_;
  // Workbook-rels entries with unrecognised Type URIs (theme, calcChain,
  // vbaProject, customXml, ...). Round-trip metadata only; the parts
  // themselves live in `passthrough_parts_`.
  std::vector<UnknownRelationship> unknown_workbook_rels_;
  // Person list referenced by the sheets' threaded comments.
  std::vector<Person> persons_;
  // Package-root relationships with unrecognised Type URIs (thumbnail,
  // digital-signature origin, and vendor extensions). Their target parts
  // remain in `passthrough_parts_`.
  std::vector<UnknownRelationship> unknown_package_rels_;
  // `<Default>` content-type registrations captured from
  // `[Content_Types].xml` (extension -> content type). Re-emitted by the
  // writer for Default-typed passthrough parts. Empty by default.
  std::vector<DefaultContentType> default_content_types_;
  // Pivot caches owned by the workbook. One cache may be referenced by
  // multiple pivot tables (per `Sheet::pivot_tables()`).
  std::vector<std::unique_ptr<pivot::PivotCache>> pivot_caches_;
  // OOXML workbook variant. Defaults to plain `.xlsx`; the reader sets
  // this from `[Content_Types].xml` and the writer consults it when
  // emitting the workbook content-type Override. Plain data; no
  // lifecycle implications. The default `kXlsx` value is assigned in
  // the out-of-line constructor (workbook.cpp) so `workbook.h` does
  // not need the full `workbook_kind.h` definition.
  WorkbookKind kind_;
  // Workbook-level calc mode. Round-trip metadata mirroring `<calcPr
  // calcMode=...>`. Default `kAuto` matches a freshly created
  // workbook in Excel.
  CalcMode calc_mode_ = CalcMode::kAuto;
  // 1904 date system flag parsed from `<workbookPr date1904>`. Default
  // false (1900 system).
  bool date1904_ = false;
  // Pinned wall-clock reading for the clock seam above. Empty by default so
  // an untouched workbook behaves exactly as it did before the seam existed.
  std::optional<date_time::CivilTime> pinned_now_;
  // Raw workbook.xml level elements captured for verbatim re-emission
  // (see the `workbook_pr_xml` accessor group). Empty when absent from
  // the source.
  std::string workbook_pr_xml_;
  std::string book_views_xml_;
  std::string workbook_protection_xml_;
  std::string file_version_xml_;
  std::string file_sharing_xml_;
  std::string workbook_ext_lst_xml_;
  // Extra `<workbook>` root namespace declarations (xmlns:mc / xmlns:xr* /
  // x15 + mc:Ignorable) captured for verbatim re-emission so namespaced
  // attributes in the raw captures above resolve. Empty by default.
  std::string workbook_root_extra_attrs_;
  // Formula compatibility profile. Runtime default is Windows Excel 365 ja-JP.
  ExcelProfile excel_profile_ = default_excel_profile();
  // Workbook-scoped style records (fonts, fills, borders, num fmts,
  // and the cellXfs index that ties them together). The default table
  // is empty and the writer falls back to a minimal-but-valid styles
  // document; the OOXML reader replaces this wholesale via
  // `set_styles(...)` when an `xl/styles.xml` part is present.
  StylesTable styles_;
  // Backing store for every `Value::text` view owned by cells in this
  // workbook. See the `intern_text` / `mutable_text_storage` block in
  // the public API for the contract. `std::deque` is required for
  // pointer stability across appends.
  std::deque<std::string> text_storage_;
};

}  // namespace formulon

#endif  // FORMULON_WORKBOOK_H_
