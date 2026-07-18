// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
#include <string>
#include <string_view>
#include <vector>

#include "eval/compat.h"
#include "io/calc_mode.h"
#include "io/default_content_type.h"
#include "io/defined_names.h"
#include "io/external_links.h"
#include "io/passthrough_part.h"
#include "io/styles_reader.h"
#include "io/tables_reader.h"
#include "io/unknown_relationship.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {

// `io::WorkbookKind` is held by value but is an enum class with an
// explicit underlying type, so we can opaque-forward-declare it here.
// Concrete callers that need to name enumerators include
// `io/workbook_kind.h` directly. The default value is assigned in the
// out-of-line `Workbook()` constructor in workbook.cpp so the header
// itself never references an enumerator.
namespace io {
enum class WorkbookKind : std::uint8_t;
enum class WorkbookFormat : std::uint8_t;
}  // namespace io

namespace eval {
class FunctionRegistry;
class RecalcEngine;
struct IterativeOptions;
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
  ///
  /// This unchecked overload is for trusted callers (the package reader,
  /// which trusts names Excel already wrote, and internal setup). Public
  /// API callers should use `add_sheet_validated`.
  Sheet& add_sheet(std::string name);

  /// Appends a new sheet after validating `name` the same way
  /// `rename_sheet` does: non-empty, at most 31 UTF-16 code units, no
  /// forbidden character (`: \ / ? * [ ]`), and no case-insensitive
  /// collision with an existing sheet. Returns a pointer to the new sheet
  /// on success (owned by the Workbook, invalidated by later structural
  /// mutations) or `kInvalidSheetName` on any violation.
  Expected<Sheet*, Error> add_sheet_validated(std::string name);

  /// Renames the sheet at `index` to `new_name`.
  ///
  /// Updates the sheet's stored name and any workbook-scoped defined-name
  /// targets that referenced the sheet by its old name. Cell formulas
  /// inside the renamed sheet (and other sheets) are LEFT UNTOUCHED:
  /// post-rename, cell formulas may carry stale sheet references; the
  /// AST-level reference shifter generalisation handles those in a
  /// separate follow-up bundle. Tables and pivot caches likewise keep
  /// their authored sheet identifiers — they reference sheets by index,
  /// not by name, so the rename is invisible to them.
  ///
  /// Errors:
  ///   * `kSheetIndexOutOfRange` when `index >= sheet_count()`.
  ///   * `kInvalidSheetName` when `new_name` is empty, exceeds 31
  ///     characters, contains a forbidden character (`: \ / ? * [ ]`),
  ///     or collides case-insensitively with another sheet's name. A
  ///     no-op rename to the sheet's existing name (case-insensitively
  ///     equal) is accepted: the sheet's stored name is updated to the
  ///     new casing, no other state changes.
  Expected<void, Error> rename_sheet(std::uint32_t index, std::string new_name);

  /// Removes the sheet at `index`.
  ///
  /// Defined names that reference the removed sheet are dropped (after
  /// a future ref-shifter bundle they would surface `#REF!` instead;
  /// the simpler drop semantics are an explicit limitation of this
  /// bundle). The recalc engine's dep-graph entries for every cell on
  /// the removed sheet are also dropped so subsequent recalcs do not
  /// chase dangling references; cells on remaining sheets keep their
  /// edges, but any edges that pointed into the removed sheet are
  /// gone.
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
  /// Equivalent to `save_ex(io::WorkbookFormat::Ooxml)`.
  Expected<std::vector<std::uint8_t>, Error> save() const;

  /// Serialises the workbook using an explicit container `format`.
  /// `io::WorkbookFormat::Ooxml` delegates to `io::write_ooxml` (the
  /// `.xlsx` writer); `io::WorkbookFormat::Xlsb` delegates to
  /// `io::xlsb::write_xlsb` (the MS-XLSB writer). `io::WorkbookFormat::Unknown`
  /// is not a valid save target and returns `kInvalidArgument`.
  Expected<std::vector<std::uint8_t>, Error> save_ex(io::WorkbookFormat format) const;

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

  /// Read-only access to the workbook's table-metadata list (in
  /// archive-discovery order, which matches the per-sheet rels walk).
  const std::vector<io::TableMetadata>& tables() const noexcept { return tables_; }

  /// Replaces the workbook's table-metadata list. Move-assigns for the
  /// same reason as `set_defined_names`.
  void set_tables(std::vector<io::TableMetadata> tables) { tables_ = std::move(tables); }

  /// Read-only access to the verbatim parts the reader did not model.
  /// The writer emits each entry as-is. `<Override>`-listed parts
  /// replicate their `<Override>` registration in `[Content_Types].xml`;
  /// Default-typed binary/media parts (vbaProject.bin, images, drawings,
  /// VML, and their rels) rely on the round-tripped `<Default>` entries
  /// exposed via `default_content_types()`.
  const std::vector<io::PassthroughPart>& passthrough_parts() const noexcept { return passthrough_parts_; }

  /// Replaces the workbook's passthrough-part list. Move-assigns to
  /// keep the I/O hand-off allocation-free for archives carrying many
  /// preserved parts.
  void set_passthrough_parts(std::vector<io::PassthroughPart> parts) { passthrough_parts_ = std::move(parts); }

  /// Read-only access to the verbatim workbook-level `<Relationship>`
  /// entries (from `xl/_rels/workbook.xml.rels`) whose Type URI the
  /// reader did not recognise. Captured so the writer can re-emit
  /// every entry, keeping passthrough-listed parts (theme, calcChain,
  /// vbaProject, customXml, ...) reachable through the relationship
  /// graph. The writer mints fresh rIds; the original `id` is
  /// preserved on the struct for diagnostics only.
  const std::vector<io::UnknownRelationship>& unknown_workbook_rels() const noexcept { return unknown_workbook_rels_; }

  /// Replaces the workbook's unknown-relationship list. Move-assigns
  /// to keep the I/O hand-off allocation-free.
  void set_unknown_workbook_rels(std::vector<io::UnknownRelationship> rels) {
    unknown_workbook_rels_ = std::move(rels);
  }

  /// Read-only access to the `<Default>` content-type registrations the
  /// reader captured from `[Content_Types].xml`. The writer re-emits the
  /// entries whose extension is used by a Default-typed passthrough part
  /// (vbaProject.bin, images, VML, ...) so those parts keep a resolvable
  /// content type in the round-tripped package.
  const std::vector<io::DefaultContentType>& default_content_types() const noexcept { return default_content_types_; }

  /// Replaces the workbook's captured `<Default>` content-type list.
  /// Move-assigns to keep the I/O hand-off allocation-free.
  void set_default_content_types(std::vector<io::DefaultContentType> defaults) {
    default_content_types_ = std::move(defaults);
  }

  /// Read-only access to the workbook's external-link list (in
  /// `<externalReferences>` document order). Each entry surfaces the
  /// relationship metadata for one cross-workbook reference; the body
  /// part itself round-trips through `passthrough_parts()` unchanged.
  const std::vector<io::ExternalLinkRecord>& external_links() const noexcept { return external_links_; }

  /// Replaces the workbook's external-link list. Move-assigns to keep
  /// the I/O hand-off allocation-free.
  void set_external_links(std::vector<io::ExternalLinkRecord> links) { external_links_ = std::move(links); }

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

  /// Workbook-level calc-mode enum. Type-aliased from `io::CalcMode`
  /// (declared in `io/calc_mode.h`) so the C ABI, the OOXML
  /// reader/writer and the bindings can all reference the enum
  /// without pulling in the full `workbook.h`. Existing source that
  /// names `Workbook::CalcMode::kAuto` keeps compiling unchanged.
  using CalcMode = io::CalcMode;

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
  // `<workbookPr date1904>` and re-emitted verbatim via the raw
  // `<workbookPr>` capture below.

  /// True when the workbook uses the 1904 date system. Defaults to false
  /// (1900 system). Consumed by date-serial conversions
  /// (`eval::date_time::serial_from_ymd` / `ymd_from_serial`).
  bool date1904() const noexcept { return date1904_; }

  /// Sets the 1904-date-system flag. Plain model value; the raw
  /// `<workbookPr>` capture (`workbook_pr_xml`) is the source of truth for
  /// re-emission when present, so programmatic callers that need the
  /// attribute written should clear `workbook_pr_xml` or rely on the
  /// synthesised fallback the writer emits when no raw block exists.
  void set_date1904(bool value) noexcept { date1904_ = value; }

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
  eval::ExcelProfile excel_profile() const noexcept { return excel_profile_; }

  /// Sets the full formula-behaviour profile used by future recalc calls.
  void set_excel_profile(eval::ExcelProfile profile) noexcept { excel_profile_ = profile; }

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
  const io::StylesTable& styles() const noexcept { return styles_; }

  /// Replaces the workbook's styles table. Move-assigns to keep the
  /// reader hand-off allocation-free.
  void set_styles(io::StylesTable styles) { styles_ = std::move(styles); }

  /// Mutable access to the workbook's styles table.
  ///
  /// The OOXML reader populates the table during package load via
  /// `set_styles`; mutators added through the C ABI (font / fill /
  /// border / num-fmt / xf inserts) reach the underlying records
  /// through this accessor instead of round-tripping the whole table.
  io::StylesTable& mutable_styles() noexcept { return styles_; }

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
  // Range clamping (Excel's behaviour where deleting only part of a
  // range shrinks the range rather than collapsing it) is a follow-up
  // enhancement; the current implementation collapses to `#REF!` whenever
  // a range endpoint sits inside the deleted interval. All other Excel
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
  std::vector<io::ExternalLinkRecord> external_links_;
  // Workbook-rels entries with unrecognised Type URIs (theme, calcChain,
  // vbaProject, customXml, ...). Round-trip metadata only; the parts
  // themselves live in `passthrough_parts_`.
  std::vector<io::UnknownRelationship> unknown_workbook_rels_;
  // `<Default>` content-type registrations captured from
  // `[Content_Types].xml` (extension -> content type). Re-emitted by the
  // writer for Default-typed passthrough parts. Empty by default.
  std::vector<io::DefaultContentType> default_content_types_;
  // Pivot caches owned by the workbook. One cache may be referenced by
  // multiple pivot tables (per `Sheet::pivot_tables()`).
  std::vector<std::unique_ptr<pivot::PivotCache>> pivot_caches_;
  // OOXML workbook variant. Defaults to plain `.xlsx`; the reader sets
  // this from `[Content_Types].xml` and the writer consults it when
  // emitting the workbook content-type Override. Plain data; no
  // lifecycle implications. The default `kXlsx` value is assigned in
  // the out-of-line constructor (workbook.cpp) so `workbook.h` does
  // not need the full `io/workbook_kind.h` definition.
  io::WorkbookKind kind_;
  // Workbook-level calc mode. Round-trip metadata mirroring `<calcPr
  // calcMode=...>`. Default `kAuto` matches a freshly created
  // workbook in Excel.
  CalcMode calc_mode_ = CalcMode::kAuto;
  // 1904 date system flag parsed from `<workbookPr date1904>`. Default
  // false (1900 system).
  bool date1904_ = false;
  // Raw workbook.xml level elements captured for verbatim re-emission
  // (see the `workbook_pr_xml` accessor group). Empty when absent from
  // the source.
  std::string workbook_pr_xml_;
  std::string book_views_xml_;
  std::string workbook_protection_xml_;
  // Extra `<workbook>` root namespace declarations (xmlns:mc / xmlns:xr* /
  // x15 + mc:Ignorable) captured for verbatim re-emission so namespaced
  // attributes in the raw captures above resolve. Empty by default.
  std::string workbook_root_extra_attrs_;
  // Formula compatibility profile. Runtime default is Windows Excel 365 ja-JP.
  eval::ExcelProfile excel_profile_ = eval::default_excel_profile();
  // Workbook-scoped style records (fonts, fills, borders, num fmts,
  // and the cellXfs index that ties them together). The default table
  // is empty and the writer falls back to a minimal-but-valid styles
  // document; the OOXML reader replaces this wholesale via
  // `set_styles(...)` when an `xl/styles.xml` part is present.
  io::StylesTable styles_;
  // Backing store for every `Value::text` view owned by cells in this
  // workbook. See the `intern_text` / `mutable_text_storage` block in
  // the public API for the contract. `std::deque` is required for
  // pointer stability across appends.
  std::deque<std::string> text_storage_;
};

}  // namespace formulon

#endif  // FORMULON_WORKBOOK_H_
