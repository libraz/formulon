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
#include "style_region.h"
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
class PivotTable;
struct PivotCacheField;
}  // namespace pivot

namespace parser {
class AstNode;
struct ExternalBookIndexer;
}  // namespace parser
namespace detail {
struct PivotReportPlan;
}  // namespace detail

/// The shape a table pre-check takes a table to have: its range and which
/// of its header and totals rows it shows.
struct TableShape {
  MergeRange rect{};
  bool header_row = true;
  bool totals_row = false;
};

/// An Insert / Delete Cells shift a pre-check takes as already applied, as
/// `insert_cells` / `delete_cells` would make it (`kDown` / `kRight` insert,
/// `kUp` / `kLeft` delete).
struct PendingShift {
  std::size_t sheet_index = 0;
  MergeRange range{};
  CellShift shift = CellShift::kDown;
};

/// Where a workbook is saved: the folder and the file name. Metadata only;
/// nothing reads or writes a file system through it.
///
/// `directory` is absolute in one of three canonical forms, POSIX (`/a/b`,
/// `/` alone the root), drive (`C:/a`, `C:/` the root) or UNC
/// (`//server/share/a`), with no trailing separator except at a root.
/// It is empty when the book has a name but no folder.
struct WorkbookLocation {
  std::optional<std::string> directory;
  std::string file_name;

  bool operator==(const WorkbookLocation& other) const {
    return directory == other.directory && file_name == other.file_name;
  }
  bool operator!=(const WorkbookLocation& other) const { return !(*this == other); }
};

/// How `Workbook::project_pivot` answers a report range that an occupant covers or that overflows the sheet.
enum class ReportPlacement : std::uint8_t {
  kBlock,   ///< An edit of an existing report: occupants or the sheet edge block it.
  kRefuse,  ///< Creation: occupants or the sheet edge refuse it.
};

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

  /// Factory for the workbook a new Excel book in `profile` starts as: one
  /// sheet, `excel_profile()` set to `profile`, and the locale's Normal
  /// font, phonetic font, Normal style name, sheet name, vertical alignment
  /// and page margins (`new_workbook_defaults.h`). The body font is 12 pt
  /// on Mac and 11 pt on Windows; the Mac default row height and base
  /// column width are set, the Windows ones are left at the engine
  /// defaults. Windows values other than the size are estimated from the
  /// Mac values. No theme part is created; Excel reads its absence as
  /// the Office theme. The phonetic font (font 1) keeps the locale face
  /// when font 0 is changed later. Sheets added afterwards through any
  /// `add_sheet*` receive the first sheet's row metrics, page margins and
  /// `phoneticPr`.
  static Workbook create_for_profile(const ExcelProfile& profile);

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

  /// `insert_before` value that appends a new sheet.
  static constexpr std::size_t kAppendSheet = static_cast<std::size_t>(-1);

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
  /// Counted: the cell store (row nodes, stored cells, style runs, each cell's formula
  /// text, owned cached text and phonetic text), the shared-string
  /// storage every text value borrows from, the passthrough part payloads
  /// (which embedded media dominates), and the workbook-level round-trip
  /// metadata strings. Not counted: allocator bookkeeping, the recalc
  /// engine's dependency graph and arenas, styles, and the pivot caches —
  /// all of which are bounded by the counted parts in practice.
  ///
  /// The result is an estimate for pressure reporting, not an accounting
  /// figure: it walks every stored cell, so it is O(cells) and
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
  /// Anything appending from outside the process must not use it: the
  /// file readers validate the name through `add_sheet_validated`, and
  /// every binding through `add_sheet_entered`.
  std::size_t add_sheet(std::string name);

  /// Appends a new sheet after validating `name` structurally: strict
  /// UTF-8, non-empty, at most 31 UTF-16 code units, no forbidden character
  /// (`: \ / ? * [ ]`), no leading or trailing apostrophe, and no
  /// Unicode-simple-fold collision with an existing sheet. The name only
  /// Windows Excel's naming UI refuses is left to `add_sheet_entered`. Returns a pointer to the new sheet
  /// on success (owned by the Workbook, invalidated by later structural
  /// mutations), `kAlreadyExists` on a collision, or `kInvalidSheetName` on
  /// any other violation.
  ///
  /// This is also the append path the OOXML and XLSB readers use, so the
  /// name validation `add_sheet` defers to the I/O boundary is in fact
  /// enforced there. No two loaded sheets share a case-folded name, though
  /// two may still be equal under the wide fold (`sheet_names::wide_equal`).
  ///
  /// Errors:
  ///   * `kInvalidSheetName` on any name violation.
  ///   * `kAlreadyExists` when another sheet holds the name under Unicode
  ///     simple case folding.
  ///   * `kSheetCountLimitExceeded` when `sheet_count() == kMaxSheets`.
  Expected<Sheet*, Error> add_sheet_validated(std::string name);

  /// The name a sheet added without one takes: the profile's first sheet
  /// name without its trailing number (`Sheet`, `Tabelle`, `Feuil`, ...)
  /// followed by the smallest N >= 1 that collides with no sheet name
  /// (`sheet_names::write_collides`).
  std::string next_default_sheet_name() const;

  /// Adds a new sheet whose name a user entered, placed before the sheet now
  /// at `insert_before` (`sheet_count()`, the default, appends), validating
  /// the name exactly as
  /// `rename_sheet` does: `add_sheet_validated`'s structural rules, a
  /// collision under `sheet_names::write_collides` (the wide fold, not only
  /// case), plus, under a Windows profile, not `History` in any case, which
  /// Windows Excel's naming UI reserves (Mac Excel accepts it). The add path
  /// of every binding goes through here; the file readers stay on
  /// `add_sheet_validated`.
  ///
  /// Placing a new sheet rewrites no formula: it is no endpoint of any span.
  ///
  /// Errors: as `add_sheet_validated`; `kSheetIndexOutOfRange` when
  /// `insert_before > sheet_count()`.
  Expected<Sheet*, Error> add_sheet_entered(std::string name, std::size_t insert_before = kAppendSheet);

  /// Renames the sheet at `index` to `new_name`.
  ///
  /// Updates the sheet's stored name and rewrites every affected local
  /// reference in cell formulas, defined names, conditional-format and
  /// validation formulas, table formulas, and pivot sources. Internal
  /// hyperlink locations keep their text, as Excel's do.
  /// Unrelated literals, unresolved references, and external-workbook
  /// references remain byte-identical.
  ///
  /// Errors:
  ///   * `kSheetIndexOutOfRange` when `index >= sheet_count()`.
  ///   * `kInvalidSheetName` when `new_name` is malformed UTF-8, empty, exceeds 31
  ///     UTF-16 units, contains a forbidden character (`: \ / ? * [ ]`),
  ///     begins or ends with an apostrophe, is `History` in any case under
  ///     a Windows profile.
  ///   * `kAlreadyExists` when `new_name` collides with another sheet's name
  ///     (`sheet_names::write_collides`). The sheet's own name is not
  ///     consulted, so a rename that only changes case, width or kana type
  ///     is accepted.
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
  ///   * `kPreconditionFailed` when `sheet_count() == 1` (Excel's
  ///     UI rejects the same op; we mirror it so `save()` never lands
  ///     in the empty-sheet-list state Excel refuses to open).
  Expected<void, Error> remove_sheet(std::uint32_t index);

  /// Moves the sheet at `from_index` to `to_index`.
  ///
  /// `to_index` is interpreted in the post-removal sheet list (matches
  /// Excel's UI semantics). For example, with three sheets, moving
  /// sheet 0 to the end uses `to_index == 2`, not `3`. A move to the
  /// same position is a successful no-op. References name their sheets,
  /// so they keep pointing at the same sheets; a 3-D span changes only
  /// when the move carries one endpoint past the other (the first moved
  /// after the last becomes the sheet that followed it, the last moved
  /// before the first the sheet that preceded it), collapsing to one
  /// sheet when the endpoints meet or when a single-sheet span's sheet
  /// moves. Formulas, defined names, conditional formats, validations,
  /// tables and charts are rewritten alike.
  ///
  /// Errors:
  ///   * `kSheetIndexOutOfRange` when either `from_index >=
  ///     sheet_count()` or `to_index >= sheet_count()`.
  Expected<void, Error> move_sheet(std::uint32_t from_index, std::uint32_t to_index);

  /// Copies the worksheet at `source` into this workbook, placed before the
  /// sheet now at `insert_before` (`sheet_count()` appends), and returns the
  /// copy's index.
  ///
  /// `new_name` is validated like `add_sheet_entered`; without one the copy is
  /// named as Excel names it (`Src (2)`). References qualified with the source
  /// sheet are re-pointed at the copy in its formulas, conditional formats,
  /// validations and charts; a 3-D span ending on the source becomes `#REF!`.
  /// The copy's tables are renamed and rebound, the source's sheet-scoped
  /// names, and workbook names referring to the source, gain a copy scoped to
  /// the new sheet. Pivot tables share the source's cache. Other sheets and
  /// hyperlink locations are untouched; the copy keeps a hidden source's
  /// visibility, is not tab-selected, and the active sheet does not change.
  ///
  /// Errors (the workbook is left unchanged):
  ///   * `kSheetIndexOutOfRange` when `source >= sheet_count()` or
  ///     `insert_before > sheet_count()`.
  ///   * `kSheetCountLimitExceeded` when the workbook is full.
  ///   * `kPreconditionFailed` for a very hidden source.
  ///   * `kNotImplemented` when the source is not a worksheet or carries
  ///     content the copy cannot model: unrecognised relationships, slicers,
  ///     drawing objects other than pictures and charts, pivot charts, or a
  ///     `.xlsb` drawing the reader could not lift.
  ///   * `kInvalidSheetName` when `new_name` fails validation.
  ///   * `kAlreadyExists` when `new_name` collides with another sheet's name.
  Expected<std::size_t, Error> copy_sheet(std::size_t source, std::size_t insert_before,
                                          std::optional<std::string> new_name);

  /// Lookup by display name through `sheet_names::resolve_by`: an exact
  /// match, else a case-fold match (`"SHEET2"` for `"Sheet2"`), else the one
  /// sheet equal under the wide fold (`"Ｓｈｅｅｔ２"`). Returns `nullptr`
  /// when no sheet matches or two sheets match only under the wide fold. A linear scan is used;
  /// workbooks typically carry O(1)–O(10) sheets, so a hash index is
  /// premature.
  const Sheet* sheet_by_name(std::string_view name) const noexcept;

  /// Returns the 0-based index of the sheet `name` resolves to, as
  /// `sheet_by_name` does. Returns
  /// `kNoSheet` when no sheet matches. Linear scan, like
  /// `sheet_by_name`.
  std::size_t sheet_index_by_name(std::string_view name) const noexcept;

  /// What `sheet_index_by_name` returns when no sheet matches.
  static constexpr std::size_t kNoSheet = static_cast<std::size_t>(-1);

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
  /// Returns `kInvalidArgument` when `sheet_index >= sheet_count()`, and
  /// `kPreconditionFailed`, changing nothing, when the cell is a non-anchor
  /// cell of a legacy array formula block; a write to the anchor replaces
  /// the whole block. A cell inside a data table's result range is refused
  /// the same way, anchor included.
  Expected<void, Error> set_cell_value(std::size_t sheet_index, std::uint32_t row, std::uint32_t col, Value value);

  /// Stores a text literal whose bytes are copied into the destination cell.
  /// This is appropriate for short-lived caller buffers; source readers that
  /// keep a workbook-scoped shared-string store can use `set_cell_value`
  /// with `Value::text` instead. Refuses a non-anchor cell of a legacy array
  /// formula block like `set_cell_value`.
  Expected<void, Error> set_cell_text(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                      std::string_view text);

  /// Where formula text handed to the workbook comes from.
  enum class FormulaSource : std::uint8_t {
    /// Text Excel parses on entry, typed, set through an API or read from an
    /// `.xlsx` part: the whitespace its parsed form cannot hold is dropped,
    /// its number literals are respelled from their values and a unary sign
    /// touching one is absorbed (`parser::spell_entered_formula`).
    kEntered,
    /// Text already in stored spelling: rendered from a parsed formula (an
    /// `.xlsb` Ptg stream) or already normalized as `kEntered`. Respelling
    /// it again would absorb a sign Excel kept (`=--0` is stored as `=-0`).
    kRendered,
  };

  /// Stores a formula at `(row, col)` on `sheet_index` and registers its
  /// dependencies with the recalc engine.
  ///
  /// `formula` is the raw Excel-style formula text. Whether or not it begins
  /// with `=` is not checked here — the parser owns that contract — but
  /// callers typically pass the leading `=`. The stored text is
  /// `normalize_formula_text(formula, source)`, `kEntered` text then bound by
  /// `bind_self_book_names`. The cell's `cached_value` is
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
  /// Returns `kInvalidArgument` when `sheet_index >= sheet_count()`, and
  /// `kPreconditionFailed`, changing nothing, when the cell is a non-anchor
  /// cell of a legacy array formula block; a write to the anchor replaces
  /// the whole block.
  Expected<void, Error> set_cell_formula(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                         std::string formula, FormulaSource source = FormulaSource::kEntered);

  /// The legacy array formula (CSE block) that `(row, col)` belongs to.
  struct ArrayFormulaInfo {
    /// Top-left cell of the block, which holds the formula.
    std::uint32_t anchor_row = 0;
    std::uint32_t anchor_col = 0;
    /// Extent of the block in cells.
    std::uint32_t rows = 0;
    std::uint32_t cols = 0;
    /// The anchor's stored formula text, as `Sheet::read_formula_cell` returns it.
    std::string formula;
  };

  /// The legacy array formula covering `(row, col)`, identical for every cell
  /// of the block, or `nullopt` when the cell is outside any block, the sheet
  /// does not exist or the coordinate is outside the grid. A dynamic-array
  /// spill is not a legacy array formula.
  std::optional<ArrayFormulaInfo> array_formula_at(std::size_t sheet_index, std::uint32_t row, std::uint32_t col) const;

  /// Enters `formula` as one legacy array formula over the `rows` x `cols`
  /// rectangle whose top-left cell is `(row, col)`, as Ctrl+Shift+Enter does.
  /// A 1 x 1 rectangle is allowed. Every cell of the rectangle loses its value
  /// and formula (styles are kept); the anchor takes the formula, marked as a
  /// block of that extent and not as a dynamic array, and the next recalc
  /// distributes the result, filling `#N/A` where a smaller result cannot
  /// reach. A write to the anchor replaces the whole block, which is how a
  /// block is cleared.
  ///
  /// Returns `kInvalidArgument` for a missing sheet, a zero extent, or a
  /// rectangle outside the grid or larger than a spill may be. Returns
  /// `kPreconditionFailed`, changing nothing, when the rectangle partially
  /// overlaps another array block (one wholly inside it is replaced),
  /// intersects a merged area, a data table's result range or a PivotTable,
  /// or, for more than one cell, a table. A formula that does not parse is
  /// stored as `set_cell_formula` stores it.
  Expected<void, Error> set_array_formula(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                          std::uint32_t rows, std::uint32_t cols, std::string formula,
                                          FormulaSource source = FormulaSource::kEntered);

  /// Formula-bar spelling of formula text as the model stores it, for cells,
  /// defined names, conditional formats and data validations alike: Excel's
  /// `_xlfn.` / `_xlws.` / `_xlpm.` storage prefixes stripped, the stored
  /// `SINGLE` / `ANCHORARRAY` calls read back as `@x` / `x#`, and, for
  /// `kEntered` text, `parser::spell_entered_formula` applied. Exposed so an
  /// entry check judges the text that would be stored.
  static std::string normalize_formula_text(std::string formula, FormulaSource source);

  /// Model spelling of a formula read from a file:
  /// `normalize_formula_text(stored, source)`, then every `[N]` link index
  /// rewritten to the book name the formula bar shows (with the link's
  /// absolute directory when it has one), then `bind_external_books`.
  /// `[0]!Name` is kept as stored. The result is `kRendered` text.
  std::string ingest_stored_formula(std::string_view stored, FormulaSource source);

  /// Model spelling of entered formula text, in place:
  /// `normalize_formula_text(formula, kEntered)`, then the structured-
  /// reference bind of a formula outside every table (`T[@X]` to the long
  /// form, a bare table name not hidden by a workbook-scope defined name to
  /// `T[]`), then `bind_self_book_names`, then `bind_external_books`.
  void ingest_entered_formula(std::string& formula);

  /// The stored form of the structured references in `text` entered at
  /// (`sheet`, `row`, `col`) (`parser::bind_structured_refs`): an unqualified
  /// `[...]` takes the table holding that cell, and a bare table name not
  /// hidden by a defined name visible from `sheet` gains `[]`.
  std::string bind_entered_structured_refs(std::string_view text, std::size_t sheet, std::uint32_t row,
                                           std::uint32_t col) const;

  /// Where the workbook is saved; empty when unset, as for a book that has
  /// not been saved. Loading from bytes leaves it unset. It drives
  /// `CELL("filename")`, `CELL("address")` for other sheets, the `[0]!Name`
  /// spelling, the saved `absPath` and the relative form of local hyperlink
  /// targets in a saved file.
  const std::optional<WorkbookLocation>& location() const noexcept { return location_; }

  /// Sets `location()`; an empty `new_location` unsets it. The directory is
  /// stored in its canonical form (see `WorkbookLocation`): `C:\a`, `c:/a/`
  /// become `C:/a`, `\\server\share\a` becomes `//server/share/a`. Returns
  /// `kInvalidArgument`, changing nothing, when the file name is empty, does
  /// not end in a workbook extension, holds one of `[ ] : / \ ? *` or NUL, is
  /// not valid UTF-8, is longer than 218 bytes, or equals the book name of an
  /// external link under ASCII case folding; or when the directory is empty,
  /// not absolute in one of the three forms (`C:`, `C:rel`, `//a`,
  /// `\\server` refused), holds NUL, `[` or `]`, is not valid UTF-8, has a
  /// `.` / `..` or empty inner segment, or makes `directory/file_name` longer
  /// than 1,024 bytes. When a directory first appears, every relative local
  /// hyperlink target is resolved against it (`resolve_relative_target`).
  Expected<void, Error> set_location(std::optional<WorkbookLocation> new_location);

  /// The file name of `location()`, extension included; empty when unset. A
  /// formula bar spells `[0]!Name` with it (`M5.xlsx!Name`).
  const std::string& file_name() const noexcept;

  /// `text` with every `[0]!Name` spelled with `file_name()`
  /// (`parser::spell_self_book_names`); unchanged when no name is set.
  std::string spell_self_book_names(std::string_view text) const;

  /// Whether `spell_self_book_names` may change `text`.
  bool spells_self_book_names(std::string_view text) const noexcept;

  /// Entered `text` bound to this workbook (`parser::bind_self_book_names`):
  /// `<file_name>!Name` becomes `[0]!Name` and `[<file_name>]Sheet!ref`
  /// becomes `Sheet!ref`; unchanged when no name is set.
  std::string bind_self_book_names(std::string_view text) const;

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
  /// Returns `kInvalidArgument` for a missing sheet or a range outside the
  /// grid, `kAlreadyExists` when the range intersects an existing merge (an
  /// identical one included), and `kPreconditionFailed` when it intersects a
  /// table, a PivotTable, a legacy array formula or a data table's result
  /// range; nothing changes on a refusal. File readers do not come here.
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

  // ---------------------------------------------------------------------------
  // Hyperlinks
  // ---------------------------------------------------------------------------

  /// Sets `link` the way Excel adds a hyperlink: one on exactly the same range
  /// is replaced in place, otherwise `link` is appended. Every cell of the
  /// range takes the built-in Hyperlink cell style (found by builtin id 8,
  /// created when absent): the font becomes the Normal font underlined in
  /// theme colour 10 and the rest of the cell format is kept; blank and
  /// merge-covered cells are formatted, a range spanning every row (column)
  /// styles the column (row). An empty top-left cell (no value, formula or
  /// spilled value, not hidden by a merge) receives `display`, else `target`,
  /// else `location`, as text. With a location directory set, a relative
  /// local target is stored resolved against it (`resolve_relative_target`).
  /// `kInvalidArgument` for a bad sheet or range or
  /// that text over 32,767 UTF-16 units; `kSecurityResourceLimit` when the
  /// range fails `check_styled_range_area`. Nothing changes on failure. File
  /// readers do not come here.
  Expected<void, Error> set_hyperlink(std::size_t sheet_index, Hyperlink link);

  /// Removes every hyperlink intersecting `range`, each whole, and resets every
  /// cell of each removed link's range to cell format 0, cells another link
  /// still covers included. Returns how many were removed (0 changes nothing).
  /// `kSecurityResourceLimit`, nothing removed, when a removed link's range
  /// fails `check_styled_range_area`; `Sheet::mutable_hyperlinks` removes such
  /// a link without the reset.
  Expected<std::size_t, Error> remove_hyperlinks_intersecting(std::size_t sheet_index, MergeRange range);

  /// Removes every hyperlink on the sheet with the reset of
  /// `remove_hyperlinks_intersecting`.
  Expected<void, Error> clear_hyperlinks(std::size_t sheet_index);

  // ---------------------------------------------------------------------------
  // What-If data tables
  // ---------------------------------------------------------------------------

  /// Creates a data table whose result range is `ref` (the cells of its
  /// top row and left column outside `ref` hold the formulas and the
  /// substitution values). `row_input` / `col_input` name the input cells
  /// the values of the top row / left column are substituted into: one input
  /// gives a one-variable table, both a two-variable one. The result cells
  /// are set to 0 until the next recalc fills them.
  ///
  /// Returns `kInvalidArgument` for a bad sheet index, a `ref` or input
  /// outside the grid, and `kDataTableInvalid` when no input is given, `ref`
  /// has no row above or column left of it, an input lies inside `ref`, or
  /// `ref` overlaps another table.
  Expected<void, Error> add_data_table(std::size_t sheet_index, MergeRange ref, DataTableInput row_input,
                                       DataTableInput col_input);

  /// The data table whose result range contains `(row, col)`, or null.
  const DataTable* data_table_at(std::size_t sheet_index, std::uint32_t row, std::uint32_t col) const noexcept;

  /// Removes the table containing `(row, col)` and blanks its result cells.
  /// `kDataTableInvalid` when no table covers the cell.
  Expected<void, Error> remove_data_table(std::size_t sheet_index, std::uint32_t row, std::uint32_t col);

  /// Recomputes every data table now, whatever the calc mode: the explicit
  /// calculation that `kAutoNoTable` leaves to the host. `recalc()` and
  /// `recalc_parallel()` do the same after a pass that had a dirty cell or a
  /// volatile formula, except under `kAutoNoTable`. Each result cell is found by
  /// substituting into the input cell(s), recalculating the cells that read
  /// them, and restoring the inputs; `partial_recalc` leaves tables alone.
  Expected<void, Error> recalc_data_tables(const eval::FunctionRegistry& registry);

  /// Drives a full incremental recalc using the embedded recalc engine.
  ///
  /// `registry` supplies the function dispatch table — typically
  /// `eval::default_registry()`. Returns the engine's `RecalcStats` on
  /// success, of the first pass. After the pass, blocked PivotTable reports
  /// whose range came free are rendered (`settle_blocked_pivots`) and their
  /// dependents recalculated in further passes, at most one more than the
  /// blocked reports at the start.
  Expected<eval::RecalcStats, Error> recalc(const eval::FunctionRegistry& registry);

  /// Brings cached values up to date before a value read or a save: unless
  /// the calc mode is `kManual`, runs `recalc(registry)` when the dirty set
  /// is non-empty and discards its stats. A clean workbook is left alone, so
  /// volatile formulas are not recomputed per read. Called from inside a
  /// recalc on this thread it fails with `kGraphRecalcReentrant`; any recalc
  /// error is returned as is, with the dirty set as `recalc()` left it.
  Expected<void, Error> ensure_calculated(const eval::FunctionRegistry& registry);

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
  /// Wrapper around `eval::recalc_parallel(*this, registry, cfg, stats)`,
  /// settling blocked PivotTable reports after it as `recalc` does; `stats`
  /// describes the first pass.
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
  /// to evaluation, matching the I/O readers' raw name-list hand-off. Each
  /// formula passes through `bind_self_book_names`.
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
  /// removes the matching entry. `name` resolves within the same scope only
  /// (`sheet_names::resolve_by`); the entry found takes the new spelling.
  /// The formula passes through `bind_self_book_names`.
  ///
  /// Returns `kInvalidArgument` when `name` is empty, is not a valid new
  /// name, or `local_sheet_id` names a sheet outside the workbook, and
  /// `kAlreadyExists` when a new name collides (`sheet_names::write_collides`)
  /// with a table's name or with a name of the same scope it does not
  /// resolve to.
  Expected<void, Error> set_defined_name_scoped(std::string name, std::string formula, std::int32_t local_sheet_id);

  /// Sets the hidden flag of the defined name `name` in `local_sheet_id`
  /// scope (`-1` for workbook scope). Returns `kInvalidArgument` when no
  /// such name exists.
  Expected<void, Error> set_defined_name_hidden(std::string_view name, std::int32_t local_sheet_id, bool hidden);

  /// Sets the Name Manager comment of the defined name `name` in
  /// `local_sheet_id` scope (`-1` for workbook scope); empty clears it.
  /// Returns `kInvalidArgument` when no such name exists.
  Expected<void, Error> set_defined_name_comment(std::string_view name, std::int32_t local_sheet_id,
                                                 std::string comment);

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

  /// Hands out the next session id for `TableMetadata::runtime_id`,
  /// `pivot::PivotTable::runtime_id()`, `DataValidation::runtime_id` and the
  /// sheet's sparkline-group ids: one counter for all, starting at 1, never
  /// reused, never saved.
  std::uint32_t allocate_runtime_id() const noexcept { return next_runtime_id_++; }

  // --- Data-validation and sparkline-group ids ---
  // Both collections are addressed by session id (see `utils/runtime_id.h`):
  // assigned on first request or on add, kept across structural edits and the
  // removal of other members, fresh on a copied sheet and after a load.

  /// Session id of validation `index` on sheet `sheet`, assigned on first
  /// request. Precondition: both indices in range.
  std::uint32_t validation_runtime_id(std::size_t sheet, std::size_t index) const;

  /// Position of the validation holding session id `id` on sheet `sheet`, or
  /// `std::nullopt` (also for an out-of-range sheet).
  std::optional<std::size_t> find_validation(std::size_t sheet, std::uint32_t id) const noexcept;

  /// Appends `validation` to sheet `sheet` and returns its session id, always
  /// a fresh one. `kSheetIndexOutOfRange` for an out-of-range sheet.
  Expected<std::uint32_t, Error> add_validation(std::size_t sheet, DataValidation validation);

  /// Removes the validation holding session id `id`. `kSheetIndexOutOfRange`
  /// for an out-of-range sheet, `kNotFound` when no validation holds `id`.
  Expected<void, Error> remove_validation(std::size_t sheet, std::uint32_t id);

  /// Session id of sparkline group `index` (document order in
  /// `Sheet::ext_lst_xml()`) on sheet `sheet`, assigned on first request.
  /// Precondition: both indices in range.
  std::uint32_t sparkline_group_runtime_id(std::size_t sheet, std::size_t index) const;

  /// Position of the sparkline group holding session id `id` on sheet
  /// `sheet`, or `std::nullopt` (also for an out-of-range sheet). Adding and
  /// removing a group by id live with the group model, in
  /// `io/sparkline_groups.h`.
  std::optional<std::size_t> find_sparkline_group(std::size_t sheet, std::uint32_t id) const;

  /// Replaces the workbook's table-metadata list and re-registers every
  /// formula whose structured reference names a table of the old or new
  /// list, as `set_defined_names` does for names: a reader stores cell
  /// formulas before their tables.
  void set_tables(std::vector<TableMetadata> tables);

  /// Validates and appends `table`, assigning its id, default AutoFilter and
  /// the formula reindex a new table needs; returns its index in `tables()`.
  /// `table.ref` must be a single-area A1 range as wide as `table.columns`,
  /// whose names must be non-empty and pairwise free of
  /// `sheet_names::write_collides` collisions.
  /// Returns `kInvalidArgument` for a bad sheet index, an empty or unparsable
  /// identity or range, a column mismatch, or a display name
  /// `is_valid_table_name` refuses; `kAlreadyExists` for a name or display
  /// name any existing table (a verbatim-retained XLSB table part included)
  /// or defined name of any scope already uses; `kPreconditionFailed` for a
  /// range that overlaps another table, a PivotTable, a data table or a
  /// multi-cell array formula, or cuts through merged cells. Merges wholly
  /// inside the range are removed. `table.style_info` is kept as given.
  ///
  /// A one-row range with a header row becomes the header plus one blank
  /// data row below it (`insert_row`). When that row holds something within
  /// the table's columns, those columns' cells shift down first
  /// (`insert_cells`, `insert_row_shift`); every check runs before the shift.
  /// Such a range on the grid's last row is `kPreconditionFailed`, the shift
  /// is refused as `insert_cells` refuses it, and a one-row range with both a
  /// header and a totals row is `kInvalidArgument`. A table created without a
  /// totals row is marked as never having shown one. One created with it
  /// reads each column's settings from its own totals cell
  /// (`sync_table_totals_cell`) when any totals-row cell holds content,
  /// keeps the columns' given settings when the row is blank, and takes the
  /// totals defaults `set_table_totals_row` writes only when the row is blank
  /// and no column carries a setting.
  Expected<std::size_t, Error> add_table(TableMetadata table);

  /// Checks `rect` as the new range of table `index`, as Excel's resize does:
  /// the first row stays, a table with a header row keeps a data row, and the
  /// range overlaps no other table, PivotTable, data table, multi-cell array
  /// formula or merged cell. The table's current range always passes.
  /// Returns `kPreconditionFailed` for an overlap, or for any other range
  /// while the table's creation shifted cells (`insert_row_shift`), and
  /// `kInvalidArgument` for any other refusal or an out-of-range `index`.
  Expected<void, Error> check_table_resize(std::size_t index, const MergeRange& rect) const;

  /// `check_table_resize` for table `index` taken to have shape `assumed`
  /// (its range is the "current" one) after the shifts of `pending`, in
  /// order: the structures `rect` must not overlap are read where those
  /// shifts would move them.
  Expected<void, Error> check_table_resize(std::size_t index, const MergeRange& rect, const TableShape& assumed,
                                           const std::vector<PendingShift>& pending) const;

  /// The refusals `set_table_totals_row(index, shown)` would report for
  /// table `index` taken to have shape `assumed` after the shifts of
  /// `pending`, without changing anything; `set_table_totals_row` runs the
  /// same check. Returns the cell shift the toggle makes (the row inserted
  /// below, or the totals row deleted), or nullopt when it shifts nothing.
  Expected<std::optional<PendingShift>, Error> check_table_totals_row(std::size_t index, bool shown,
                                                                      const TableShape& assumed,
                                                                      const std::vector<PendingShift>& pending) const;

  /// The refusals `insert_cells` (`kDown` / `kRight`) or `delete_cells`
  /// (`kUp` / `kLeft`) would report for `range` on `sheet_index`, without
  /// changing anything.
  Expected<void, Error> check_band_edit(std::size_t sheet_index, const MergeRange& range, CellShift shift) const;

  /// Sets table `index`'s range to `rect` after `check_table_resize`: a
  /// modelled AutoFilter follows the header and data rows, the placeholder
  /// and creation-shift marks clear, and the formulas naming the table are
  /// re-indexed. An opaque AutoFilter the caller holds is not touched yet;
  /// its text stays the caller's to rewrite. Errors as
  /// `check_table_resize`, leaving the table as it was.
  Expected<void, Error> resize_table(std::size_t index, const MergeRange& rect);

  /// Shows or hides the totals row of table `index`. Showing it inserts one
  /// row of cells below the table within its columns (`insert_cells`,
  /// shifting down), grows `ref` by that row and writes each column's totals
  /// cell there (`write_table_totals_cell`). When a merge, table, PivotTable,
  /// multi-cell array formula or data table below the table straddles its
  /// columns, nothing shifts and an empty row directly below becomes the
  /// totals row instead. Hiding it shrinks `ref` by its last row and deletes
  /// that row's cells within the table's columns (`delete_cells`, shifting
  /// up), or only clears them when such a straddling object lies below; the
  /// columns' totals settings are kept for the next show. Showing the row of
  /// a table that has never shown one (`totals_row_shown` false) first
  /// replaces the columns' settings with the defaults: the profile locale's
  /// totals label on the first column and, on the last, `sum` when every
  /// non-blank data cell holds a number not formatted as a date or time and
  /// there is one, else `count`; a one-column table takes the function
  /// alone. A call that changes
  /// nothing is a no-op. Every check runs before the first change, so a
  /// refusal leaves the table and the cells as they were. Returns
  /// `kInvalidArgument` for an out-of-range `index` or a range that is not a
  /// single area, and `kPreconditionFailed` when showing the row and the
  /// table ends on the grid's last row, the insert is refused, or a
  /// straddling object lies below and the row below is not empty (a value,
  /// formula, spilled result, merge or table there); or when hiding the row
  /// would leave a header row with no data row, its cells hold part of a
  /// multi-cell array formula or a data table, or a cell formula or defined
  /// name references the totals row (a reference meeting it, or the table's
  /// `[#Totals]` / `[#All]`).
  Expected<void, Error> set_table_totals_row(std::size_t index, bool shown);

  /// Shows or hides the header row of table `index`. Hiding it deletes the
  /// header cells outright (values and formats), moves `ref`'s first row
  /// down by one without shifting anything and removes the AutoFilter, showing
  /// the body rows when it carried criteria; the column names are kept. Showing it writes the column names into the row
  /// above the data and grows `ref` up by that row when the row holds
  /// nothing within the table's columns and the grown range passes
  /// `check_table_resize`'s overlap checks; otherwise it inserts one row of
  /// cells at the first data row within the table's columns (`insert_cells`,
  /// shifting down) and writes the names into the vacated row. Either way the
  /// table gets a plain AutoFilter over its new range. A call that changes
  /// nothing is a no-op, and every check runs before the first change.
  /// Returns `kInvalidArgument` for an out-of-range `index` or a range that
  /// is not a single area, and `kPreconditionFailed` when hiding the row
  /// would leave a table with no data row, or when the insert is refused as
  /// `insert_cells` refuses it.
  Expected<void, Error> set_table_header_row(std::size_t index, bool shown);

  /// The refusals `set_table_header_row(index, shown)` would report for
  /// table `index` taken to have shape `assumed` after the shifts of
  /// `pending`, without changing anything; `set_table_header_row` runs the
  /// same check. Returns the cell shift showing the row makes (the insert at
  /// the first data row), or nullopt when the toggle shifts nothing: hiding
  /// moves the first row down by one, showing without a shift grows the
  /// range up by one.
  Expected<std::optional<PendingShift>, Error> check_table_header_row(std::size_t index, bool shown,
                                                                      const TableShape& assumed,
                                                                      const std::vector<PendingShift>& pending) const;

  /// Writes column `column`'s totals cell on the last row of table `index`'s
  /// range from the column's totals settings: the `SUBTOTAL` formula of a
  /// built-in totals function (the 1xx code, which skips hidden rows), the
  /// custom formula, the label, or a blank. The caller decides whether the
  /// table shows a totals row. Returns `kInvalidArgument` for an out-of-range
  /// `index` or `column`, or a range that is not a single area.
  Expected<void, Error> write_table_totals_cell(std::size_t index, std::size_t column);

  /// Inserts data rows into table `table_index` of `sheet_index`, as
  /// `ListRows.Add` does: the cells of the table's columns from the first new
  /// row downwards move down by the span's length, the cells of other columns
  /// stay, and the table grows by the span. `first` and `last` are 0-based
  /// data-row positions, inclusive; `first` equal to the data-row count
  /// appends, above the totals row when there is one. Every calculated
  /// column's formula is written into the new rows (relative references moved
  /// with the row); other columns stay blank.
  ///
  /// Errors, each leaving the workbook unchanged:
  ///   * `kInvalidArgument` for an out-of-range sheet or table index, a table
  ///     not on `sheet_index`, `first` past the data-row count, `last` below
  ///     `first`, or a span reaching past the grid.
  ///   * `kPreconditionFailed` as `insert_cells` for the band below the
  ///     inserted rows' columns (a PivotTable in the shift path, a multi-cell
  ///     array formula or data table cut by the band, a value or formula, or
  ///     the table itself, pushed off the grid).
  Expected<void, Error> insert_table_rows(std::size_t sheet_index, std::size_t table_index, std::uint32_t first,
                                          std::uint32_t last);

  /// Inserts columns into table `table_index` of `sheet_index`, as
  /// `ListColumns.Add` does: the cells of the table's rows (header through
  /// totals row) from the first new column rightwards move right by the span's
  /// length, the cells of other rows stay, and the table grows by the span.
  /// `first` and `last` are 0-based table-column positions, inclusive; `first`
  /// equal to the column count appends after the last column. Each new column
  /// is named from its left neighbour (`Q2` gives `Q3`; a name without
  /// trailing digits, or no left neighbour, takes the profile's column word
  /// plus the lowest free number), the name is written to the header cell when
  /// the table has a header row, and the data and totals cells stay blank.
  ///
  /// Errors as `insert_table_rows`, with the band being the table's rows.
  Expected<void, Error> insert_table_columns(std::size_t sheet_index, std::size_t table_index, std::uint32_t first,
                                             std::uint32_t last);

  /// Makes `formula` (with or without a leading `=`) the calculated column of
  /// column `column` of table `table_index` of `sheet_index`, as entered in
  /// the column's first data row: the mark takes that spelling without `=`
  /// (a this-row reference in its stored long form), and every data cell of
  /// the column, constants included and the totals row excluded, takes the
  /// formula moved to its row. A formula with no structured reference that is
  /// not recalculated every time is registered as a shared-formula group over
  /// those cells. A table with no data row only takes the mark.
  ///
  /// Returns `kInvalidArgument` for a missing sheet, table or column, an
  /// empty formula, a dynamic-array formula, or one whose references would
  /// leave the grid in the last data row; the parse status for text that does
  /// not parse. Nothing changes on a refusal.
  Expected<void, Error> set_table_calculated_column(std::size_t sheet_index, std::size_t table_index,
                                                    std::size_t column, std::string_view formula);

  /// The refusals `set_table_calculated_column` would report for `formula`,
  /// without changing anything, for a table whose data rows are
  /// `first_data_row` to `last_data_row` inclusive (none when the last is
  /// below the first). The table itself keeps its current range, in which
  /// the formula is bound; `set_table_calculated_column` runs the same check.
  /// Returns the mark the column would take, as `TableColumn` stores it.
  Expected<std::string, Error> check_table_calculated_formula(std::size_t sheet_index, std::size_t table_index,
                                                              std::size_t column, std::string_view formula,
                                                              std::uint32_t first_data_row,
                                                              std::uint32_t last_data_row) const;

  /// Removes the calculated-column mark of column `column` of table
  /// `table_index` of `sheet_index`; the cells keep what they hold. Errors as
  /// `set_table_calculated_column` for the target.
  Expected<void, Error> clear_table_calculated_column(std::size_t sheet_index, std::size_t table_index,
                                                      std::size_t column);

  /// Removes the calculated-column mark of column `column` of table
  /// `table_index` of `sheet_index` when it has at least one data row and
  /// every data cell of the column is blank. Errors as
  /// `set_table_calculated_column` for the target.
  Expected<void, Error> drop_calculated_column_if_empty(std::size_t sheet_index, std::size_t table_index,
                                                        std::size_t column);

  /// Brings the calculated columns of table `table_index` of `sheet_index` in
  /// line with its current range: an unmarked column takes the mark
  /// `infer_calculated_column` finds, leaving its cells as they are, and a
  /// marked column keeps its mark and fills its blank data cells from sheet
  /// row `first_new_row` down (the rows a resize added) with it.
  /// Returns `kInvalidArgument` for a missing sheet or table.
  Expected<void, Error> infer_table_calculated_columns(std::size_t sheet_index, std::size_t table_index,
                                                       std::uint32_t first_new_row);

  /// Re-reads the totals cell of column `column` of table `table_index` (an
  /// index into `tables()`) into the column's totals settings
  /// (`classify_totals_cell`); a value becomes a label and its cell text. Does
  /// nothing while the table's totals row is hidden or for a missing column.
  void sync_table_totals_cell(std::size_t table_index, std::size_t column);

  /// `sync_table_totals_cell` for every column of every table showing its
  /// totals row that stores a cell there; the readers call it once cells and
  /// tables are both loaded.
  void normalize_table_totals_cells();

  /// Smallest N >= 1 such that `<word><N>` is not the name or display name
  /// of any table nor a defined name (`sheet_names::write_collides`), as
  /// `add_table` checks.
  std::string next_default_table_name(std::string_view word) const;

  /// Smallest table id above `after` that no model table and no verbatim-kept
  /// XLSB table part owns.
  std::uint32_t next_free_table_id(std::uint32_t after) const;

  /// Index in `tables()` of the first table on `sheet_index` whose `ref`
  /// (header through totals row) contains the 0-based (`row`, `col`).
  std::optional<std::size_t> table_containing(std::size_t sheet_index, std::uint32_t row,
                                              std::uint32_t col) const noexcept;

  /// True when `name` resolves to a table through its name or display name
  /// (`sheet_names::resolve_table`). Defined names are not consulted.
  bool is_table_name(std::string_view name) const noexcept;

  /// Re-registers dep-graph edges and dirties every formula cell whose
  /// structured reference names one of `table_names` (typically a table's
  /// `name` and `display_name`). Call after a `mutable_tables()` edit
  /// (create, ref/column change, remove) so structured references' dep-
  /// graph edges and cached values reflect the current table metadata --
  /// mirrors `set_defined_name_scoped`'s scoped reindex, since a
  /// `StructuredRef` resolves to a static rectangle at registration time
  /// the same way a `NameRef` does. Also hands every sheet its tables'
  /// current rectangles for spill admission and dirties the spills a
  /// rectangle that came or went blocks or releases. An empty `table_names`
  /// is a no-op.
  void reindex_formulas_for_table_change(const std::vector<std::string>& table_names);

  /// Renames table `index` and rewrites every structured reference naming its
  /// old name or display name (`sheet_names::resolve_table`) to `display_name`: cell,
  /// defined-name, metadata (CF, DV, calculated-column, totals-row) and chart
  /// formulas, plus a pivot cache sourced from the table by name. Returns
  /// `kInvalidArgument` for an out-of-range `index` and refuses a name pair
  /// as `add_table` would (`kInvalidArgument` / `kAlreadyExists`), the
  /// table's own names excepted.
  Expected<void, Error> rename_table(std::size_t index, std::string name, std::string display_name);

  /// Renames column `column` of table `table` from a typed header and
  /// returns the name it takes: an empty `name` becomes the profile's column
  /// word plus the lowest unused N >= 1, a name another column already uses
  /// (`sheet_names::write_collides`) gains the lowest unused N >= 2, and anything
  /// else, spaces and a case-only change included, is kept. Every structured
  /// reference to the old name through this table is rewritten (cell,
  /// defined-name, metadata and chart formulas), and the header cell, when
  /// the table has a header row, receives the name as text. Returns
  /// `kInvalidArgument` for an out-of-range index.
  Expected<std::string, Error> rename_table_column(std::size_t table, std::size_t column, std::string_view name);

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
  void set_passthrough_parts(std::vector<PassthroughPart> parts);

  /// Erases the passthrough entries marked non-zero in `tombstones`, keeping
  /// the surviving entries in their original order and moving their payload
  /// buffers instead of copying them. The mask must have one byte per entry.
  Expected<void, Error> erase_passthrough_parts(std::vector<std::uint8_t> tombstones);

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

  /// Sets the face of one `a:font script` entry of `a:majorFont` (`major`)
  /// or `a:minorFont`, appending the entry when the script has none. Same
  /// generation and failure rules as `set_theme_colors`; a generated part
  /// carries every default script face. Setting `Jpan` also moves `*_ea`
  /// while `a:ea` is empty. `kInvalidArgument` unless `script` is four
  /// ASCII letters and `typeface` is non-empty.
  Expected<void, Error> set_theme_script_font(bool major, std::string_view script, std::string_view typeface);

  /// Returns the workbook to the default theme by removing the theme part, its
  /// own rels part, the parts only it referenced and the workbook relationship
  /// (an unparseable part is removed the same way). A workbook without a theme
  /// part is left as is. `load_theme` then reports `ThemeSource::kDefault`.
  /// Excel adds a theme part again when the book is saved again in Excel.
  Expected<void, Error> reset_theme();

  // ---------------------------------------------------------------------------
  // Threaded comments and persons
  // ---------------------------------------------------------------------------

  /// The workbook's person list (`xl/persons/person.xml`) in file order.
  const std::vector<Person>& persons() const noexcept { return persons_; }

  /// Mutable person list for the reader; edits go through `add_person` /
  /// `remove_person`.
  std::vector<Person>& mutable_persons() noexcept { return persons_; }

  /// Appends `person`, its id stored uppercase; an empty id takes a generated
  /// GUID from the threaded-comment generator (reproducible while the clock
  /// is pinned). Fails with `kInvalidArgument`
  /// when the id is not a brace-wrapped GUID (either letter case) or the
  /// display name is empty, and with `kAlreadyExists` when the id is already
  /// listed. Person ids compare ASCII case-insensitively here and wherever
  /// a person is named.
  Expected<void, Error> add_person(Person person);

  /// Removes the person with `id`. Fails with `kNotFound` when no such person
  /// exists and with `kPreconditionFailed` when a threaded comment or mention
  /// still names them.
  Expected<void, Error> remove_person(std::string_view id);

  /// Adds `comment` to sheet `sheet` and returns its id. An empty `parent_id`
  /// opens a new thread at `row`/`col`, which must hold neither a note nor a
  /// thread; a non-empty one appends a reply to that thread, taking its anchor
  /// and ignoring `row`/`col` and `done`. An empty `id` takes a generated
  /// GUID, reproducible while the clock is pinned (`set_pinned_now`), and so
  /// does an empty mention id; an empty `created` takes the pinned clock, else the host clock. The author
  /// and mentioned persons and the parent are stored in their listed spelling,
  /// the comment and mention ids uppercase. Fails with `kInvalidArgument` for
  /// an id or mention id that is not a GUID, a malformed timestamp, a mention
  /// outside the text or a parent that is a reply, `kNotFound` for an author
  /// or mention no listed person matches, and
  /// `kAlreadyExists` for an id already used on the sheet or a cell already
  /// holding a note or thread. A failure leaves the sheet unchanged.
  Expected<std::string, Error> add_threaded_comment(std::size_t sheet, ThreadedComment comment);

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

  /// The link whose book-scope names a sheet-qualified name `qualifier!Name`
  /// denotes: the external-book link `find_external_link("", qualifier)`
  /// returns, when no sheet is named `qualifier` (compared as
  /// `sheet_by_name` does). Null otherwise, including for OLE and DDE links.
  const ExternalLinkRecord* link_for_sheet_qualifier(std::string_view qualifier) const noexcept;

  /// The `[N]` mapping a storage formatter writes cross-workbook references
  /// with, backed by `find_external_link`, and sheet-qualified names with,
  /// backed by `link_for_sheet_qualifier`. Valid while this workbook lives.
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

  /// Appends `field` to cache `cache_id` and one unplaced pivot field to every
  /// pivot bound to that cache; returns the new cache field index. Fails with
  /// `kNotFound` for an unknown cache.
  Expected<std::size_t, Error> add_pivot_cache_field(std::uint32_t cache_id, pivot::PivotCacheField field);

  /// Appends `table` to sheet `sheet` and returns its pivot index. The table
  /// is bound to its `pivot_cache_id()` cache and gets one unplaced field per
  /// cache field and a creation order after every existing pivot; its report is not projected. Fails with
  /// `kInvalidArgument` for a bad sheet or an empty name, `kNotFound` for an unknown cache, and `kAlreadyExists` for a
  /// name another pivot on the sheet holds under the name identity rule.
  Expected<std::size_t, Error> add_pivot(std::size_t sheet, std::unique_ptr<pivot::PivotTable> table);

  /// Renames pivot `pivot` on sheet `sheet`. Fails with `kInvalidArgument` for
  /// a bad sheet or an empty name, `kIndexOutOfRange` for a pivot index past the sheet's
  /// pivots, and `kAlreadyExists` for a name another pivot on the sheet holds
  /// under the name identity rule.
  Expected<void, Error> rename_pivot(std::size_t sheet, std::size_t pivot, std::string_view name);

  /// Runs Excel's Show Details on the value cell at absolute (`row`, `col`)
  /// of pivot `pivot` on sheet `sheet`, and returns the new sheet's index.
  ///
  /// The new sheet is inserted immediately before the pivot's sheet, which
  /// therefore moves to `sheet + 1`, and becomes the active sheet. It is named
  /// the locale's new-sheet word plus the smallest unused number from 1. Row 1
  /// holds the cache's database field names in cache order; below it come the
  /// records behind the cell, after the page, label, date and hidden-item
  /// filters but not the value filters. Rows are ordered by the row fields
  /// innermost first, then the column fields innermost first, each in its
  /// displayed item order, ties in source order. Each column takes its cache
  /// field's number format. The block becomes a `TableStyleMedium9` table
  /// named the locale's table word plus the smallest unused number.
  ///
  /// Errors: `kInvalidArgument` for an index out of range or a cell that is
  /// not a data, subtotal or grand-total cell; `kPreconditionFailed` when the
  /// pivot disables drill-down; `kEvalPivotMissing` when its cache is absent;
  /// the sheet-capacity error of `add_sheet_entered`.
  Expected<std::size_t, Error> show_pivot_details(std::size_t sheet, std::size_t pivot, std::uint32_t row,
                                                  std::uint32_t col);

  /// Evaluates a pivot and writes its report range into the sheet cells.
  ///
  /// The cells and merges the previous projection covered (on the first one,
  /// the location and the page area above it) are cleared, then each report
  /// cell is written as a literal with the format Excel gives its role, and
  /// the layout's merges are added. A report with no field at all is the
  /// bordered placeholder frame. A pivot over a cache loaded without records
  /// writes nothing.
  ///
  /// The new range is occupied when it covers, outside the pivot's current
  /// range, a non-blank or formula cell, a spill range cell, a merge, a legacy
  /// array block, a data-table result cell, a table or another PivotTable's
  /// current range. An occupied range, or one past the sheet edge, refuses
  /// the projection under `kRefuse`. Under `kBlock` the definition stays
  /// applied but the report is withheld: its old cells are cleared, the anchor
  /// keeps its format and holds `#SPILL!`, `useAutoFormatting` turns off and
  /// the span shrinks to the anchor (`pivot::PivotTable::blocked`). The next
  /// recalc renders it once the range is free (`settle_blocked_pivots`).
  ///
  /// Errors, each leaving the sheet unchanged:
  ///   * `kInvalidArgument` for an index out of range.
  ///   * `kEvalPivotMissing` when the pivot's cache is absent.
  ///   * `kPreconditionFailed` under `kRefuse`, naming the occupant or the edge.
  ///   * an evaluation or layout error of the pivot.
  Expected<void, Error> project_pivot(std::size_t sheet_index, std::size_t pivot_index, ReportPlacement placement);

  /// Moves a pivot's report so its anchor is (`row`, `col`), as a cut and
  /// paste does: in the destination, the rectangle the report occupies now
  /// moved by the same offset, values, formulas, merges, legacy array blocks
  /// and data tables wholly inside are cleared and tables deleted. The report
  /// then projects as `project_pivot` under `kBlock`, so only an occupant
  /// beyond that rectangle (a page-field block wider than it) can block it.
  ///
  /// Errors, each leaving the workbook unchanged:
  ///   * `kInvalidArgument` for an index or coordinate out of range.
  ///   * `kEvalPivotMissing` when the pivot's cache is absent.
  ///   * `kPreconditionFailed` when the destination rectangle leaves the
  ///     grid, covers another PivotTable, or meets a legacy array block or a
  ///     data table's result range that crosses its edge.
  ///   * an evaluation or layout error of the pivot.
  Expected<void, Error> move_pivot(std::size_t sheet_index, std::size_t pivot_index, std::uint32_t row,
                                   std::uint32_t col);

  /// Renders every blocked report whose needed range is now free, and
  /// re-plans every blocked report whose size is unknown or whose cache
  /// changed, earlier-created reports first (`pivot::PivotTable::creation_order`);
  /// returns how many reports it projected. `recalc` and
  /// `recalc_parallel` call it after each pass.
  std::size_t settle_blocked_pivots();

  /// Marks each loaded report saved in Excel's blocked form (a one-cell location with zero first-row and first-column
  /// offsets over an error anchor) as blocked with an unknown size, so the next recalc plans it and draws it once its
  /// range is free. Readers call it after the pivots and their caches are loaded.
  void adopt_saved_blocked_pivots();

  /// Restores a formula anchor saved as a blocked spill: the anchor caches `#SPILL!` and keeps the `rows` x `cols`
  /// footprint with `cause`, which the next recalc recomputes. Readers call it while decoding cells, before array
  /// spills are registered.
  Expected<void, Error> restore_blocked_spill(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                              std::uint32_t rows, std::uint32_t cols, SpillCause cause);

  /// Restores a cell saved with a rich error: its cached value becomes `code` and `error_detail` is recorded for it.
  Expected<void, Error> restore_error_detail(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                             ErrorCode code, CellErrorDetail error_detail);

  /// Restores pivot `pivot_index` of `sheet_index` saved blocked with a body of `body_rows` x `cols` and the rich
  /// error `sub_type` at its anchor (3 past the edge, else an occupant). The report is blocked with an unknown size,
  /// its page-field rows added above the body, and the anchor caches `#SPILL!`; the next settle re-plans it.
  Expected<void, Error> restore_blocked_pivot(std::size_t sheet_index, std::size_t pivot_index, std::uint32_t body_rows,
                                              std::uint32_t cols, std::uint32_t sub_type);

  /// Projects every pivot whose cache changed since its last projection. A
  /// pivot that cannot be projected stays pending.
  void project_pending_pivots();

  /// Clears the cells and merges a pivot's last projection wrote, as deleting
  /// the report does, and forgets them. Nothing happens for a pivot never
  /// projected.
  ///
  /// Errors: `kInvalidArgument` for an index out of range.
  Expected<void, Error> clear_pivot_projection(std::size_t sheet_index, std::size_t pivot_index);

  /// Every pivot table, across all sheets in sheet order, drawing from cache
  /// `cache_id`.
  std::vector<pivot::PivotTable*> pivots_of_cache(std::uint32_t cache_id);

  /// Re-derives the group fields of cache `cache_id` and forgets every result
  /// and authored span of the pivots drawing from it, leaving their reports
  /// pending; formulas are dirtied when there is such a pivot. Called after
  /// any edit of the cache. Nothing happens for an unknown cache.
  void invalidate_pivot_results(std::uint32_t cache_id);

  /// Rebuilds cache `cache_id` from its worksheet source and re-projects
  /// every pivot drawing from it, as Excel's Refresh does. Source values are
  /// read after `ensure_calculated(registry)`. A report whose new range is
  /// occupied is blocked (`project_pivot` under `kBlock`), reports being
  /// projected earliest-created first; any error leaves
  /// everything unchanged.
  ///
  /// Errors:
  ///   * `kNotFound` for an unknown cache.
  ///   * `kPreconditionFailed` when the source does not resolve
  ///     (`pivot::resolve_worksheet_source`).
  ///   * `kSecurityResourceLimit` when (rows - 1) x columns of the source
  ///     exceeds `kMaxPivotCacheRecordCells`.
  ///   * `kNotImplemented` as `pivot::build_refreshed_cache`.
  ///   * an evaluation error of `ensure_calculated` or of a pivot.
  Expected<void, Error> refresh_pivot_cache(std::uint32_t cache_id, const eval::FunctionRegistry& registry);

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
  // default: `ensure_calculated()` computes pending cells before a value
  // read or a save. `kManual` leaves them pending until `recalc()`.
  // `kAutoNoTable` behaves like `kAuto` except that `recalc()` leaves the
  // data tables to `recalc_data_tables()`. An explicit `recalc()` call
  // honours all dirty cells whatever the mode.

  /// Workbook-level calc-mode enum. Type-aliased from `CalcMode`
  /// (declared in `calc_settings.h`) so the C ABI, the OOXML
  /// reader/writer and the bindings can all reference the enum
  /// without pulling in the full `workbook.h`. Existing source that
  /// names `Workbook::CalcMode::kAuto` keeps compiling unchanged.
  using CalcMode = formulon::CalcMode;

  /// Returns the workbook-level calc mode. Defaults to `kAuto`.
  CalcMode calc_mode() const noexcept { return calc_mode_; }

  /// Sets the workbook-level calc mode. `kManual` turns
  /// `ensure_calculated()` into a no-op; `kAutoNoTable` makes `recalc()`
  /// skip the data tables.
  void set_calc_mode(CalcMode mode) noexcept { calc_mode_ = mode; }

  /// The remaining `<calcPr>` attributes (`fullPrecision`, `calcOnSave`,
  /// `refMode`, `calcId`, ...). Plain metadata.
  const CalcOptions& calc_options() const noexcept { return calc_options_; }
  CalcOptions& mutable_calc_options() noexcept { return calc_options_; }

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

  /// Index of the sheet the first `<workbookView>` marks active (`activeTab`);
  /// without `<bookViews>` (an XLSB load), the first `tabSelected` sheet. A
  /// missing or out-of-range value reads as the first sheet.
  std::uint32_t active_sheet() const;

  /// Marks sheet `index` active in the first `<workbookView>`, synthesizing a
  /// minimal `<bookViews>` when none exists and keeping `firstSheet` at or
  /// below it, and makes it the only `tabSelected` sheet. Other attributes
  /// survive verbatim. Requires `index < sheet_count()`.
  void set_active_sheet(std::uint32_t index);

  /// The sheet Excel activates when sheet `index` is hidden or deleted: the
  /// first visible one after it, else the nearest visible one before it, else
  /// `index` itself.
  std::uint32_t nearest_visible_sheet(std::uint32_t index) const;

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
  /// Every `<workbook>` child the model does not fold in (`oleSize`,
  /// `customWorkbookViews`, `fileRecoveryPr`, ...), captured verbatim with
  /// the schema slot it is written back at, ordered by slot.
  const WorksheetRawExtensions& raw_children() const noexcept { return raw_children_; }
  WorksheetRawExtensions& mutable_raw_children() noexcept { return raw_children_; }
  /// Framed `xl/workbook.bin` future-record blocks an `.xlsb` source carried
  /// that the model does not own (the slicer cache lists), naming the source
  /// workbook relationship ids; see `io/xlsb/slicer_records.h`.
  const std::vector<std::uint8_t>& xlsb_book_frt() const noexcept { return xlsb_book_frt_; }
  void set_xlsb_book_frt(std::vector<std::uint8_t> records) { xlsb_book_frt_ = std::move(records); }

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
  // en-US; oracle tests can override this per golden set.

  /// Returns the full formula-behaviour profile. Defaults to `win-365-en_US`.
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
  /// state, under the rules of `Sheet::set_cell_xf_index` (a blank
  /// coordinate takes it as a style run). Returns `kInvalidArgument` for an
  /// out-of-range `sheet_index`. Coexists with `set_cell_value` /
  /// `set_cell_formula`: those calls leave `xf_index` untouched, so the
  /// caller can layer a style update on top of a value or formula
  /// write in either order.
  Expected<void, Error> set_cell_xf_index(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                          std::uint32_t xf_index);

  /// Drops the cell-level style assignment at `(row, col)` without touching
  /// value, formula, phonetic or comment state, so a blank cell inherits
  /// its row / column style again; a blank coordinate stops existing. Returns
  /// `kInvalidArgument` for an out-of-range sheet or address.
  Expected<void, Error> clear_cell_xf_index(std::size_t sheet_index, std::uint32_t row, std::uint32_t col);

  /// Every xf some cell of `rect` outside `excluded` resolves to, ascending
  /// (see `formulon::range_style_bases`). `kInvalidArgument` for a bad sheet
  /// or a rect outside the grid or with `first > last`.
  Expected<std::vector<std::uint32_t>, Error> range_style_bases(std::size_t sheet_index, const StyleRect& rect,
                                                                std::vector<StyleCell> excluded = {}) const;

  /// Number of coordinates of `rect` with a non-zero own xf (see
  /// `formulon::count_cell_xf_indices`). Errors as `range_style_bases`.
  Expected<std::size_t, Error> count_cell_xf_indices(std::size_t sheet_index, const StyleRect& rect) const;

  /// Writes the first `capacity` of those coordinates into `out`, row-major,
  /// and returns how many it wrote. Errors as `range_style_bases`.
  Expected<std::size_t, Error> copy_cell_xf_indices(std::size_t sheet_index, const StyleRect& rect, StyleCellXf* out,
                                                    std::size_t capacity) const;

  /// Every stored cell of `rect` whose Show Phonetic flag is set, row-major.
  /// Errors as `range_style_bases`.
  Expected<std::vector<StyleCell>, Error> show_phonetic_cells(std::size_t sheet_index, const StyleRect& rect) const;

  /// Applies an xf remap to a selection the way Excel applies an
  /// attribute-only format change (see `formulon::remap_range_style`).
  /// `from` must not repeat and every `to` must address `cellXfs`, else
  /// `kInvalidArgument`; then `check_styled_range_area`. Nothing changes on
  /// failure.
  Expected<void, Error> remap_range_style(std::size_t sheet_index, const StyleRect& rect, std::vector<XfRemap> map);

  /// Gives every cell of `rect` the format `xf_index` (see
  /// `formulon::set_range_xf_index`): whole columns or rows become axis
  /// formats, blank cells style runs. `kInvalidArgument` for a bad sheet or
  /// rect, `kIndexOutOfRange` for an xf past `cellXfs`, then
  /// `check_styled_range_area`. Nothing changes on failure.
  Expected<void, Error> set_range_xf_index(std::size_t sheet_index, const StyleRect& rect, std::uint32_t xf_index);

  /// Opaque capture of `rect`'s style state for `restore_range_style`.
  Expected<std::vector<std::uint8_t>, Error> snapshot_range_style(std::size_t sheet_index, const StyleRect& rect) const;

  /// Puts a `snapshot_range_style` capture back; malformed bytes or an xf
  /// past `cellXfs` is `kInvalidArgument` with nothing changed.
  Expected<void, Error> restore_range_style(std::size_t sheet_index, const std::uint8_t* bytes, std::size_t len);

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
  /// `Sheet::kMaxRows` are dropped (their cells are lost). A table grows
  /// when the insert falls inside it.
  ///
  /// Errors, each leaving the workbook unchanged:
  ///   * `kInvalidArgument` for an out-of-range sheet index, a row beyond
  ///     the sheet bound, or `count == 0`.
  ///   * `kPreconditionFailed` when the insert falls inside a data table's
  ///     block, a legacy array block or a PivotTable, or would push one of
  ///     them or a table past the sheet bound.
  Expected<void, Error> insert_rows(std::size_t sheet_index, std::uint32_t row, std::uint32_t count);

  /// Deletes `count` rows starting at `row` on `sheet_index`. The deleted
  /// rows are dropped wholesale; rows past `row + count` shift up by
  /// `count`. A data table, legacy array block, PivotTable or table whose
  /// whole extent is deleted goes with it. Deleting a table's totals row
  /// turns its totals row off and clears every column's totals settings.
  /// Deleting every data row of a table with a header row keeps its first
  /// data row: the rest go, the kept row's table cells are cleared and the
  /// table is marked `insert_row`.
  ///
  /// Errors, each leaving the workbook unchanged:
  ///   * `kInvalidArgument` on the same argument errors as `insert_rows`,
  ///     or a `count` running past the sheet bound.
  ///   * `kPreconditionFailed` when the delete covers part, but not all, of
  ///     a data table's block, a legacy array block or a PivotTable; covers
  ///     a table's header row but not the whole table; or deletes every data
  ///     row of one table while it covers another table's row the kept row
  ///     would change.
  Expected<void, Error> delete_rows(std::size_t sheet_index, std::uint32_t row, std::uint32_t count);

  /// Inserts `count` columns at `col` on `sheet_index`. Mirrors
  /// `insert_rows` along the column axis.
  Expected<void, Error> insert_cols(std::size_t sheet_index, std::uint32_t col, std::uint32_t count);

  /// Deletes `count` columns starting at `col` on `sheet_index`. Mirrors
  /// `delete_rows` along the column axis.
  Expected<void, Error> delete_cols(std::size_t sheet_index, std::uint32_t col, std::uint32_t count);

  /// Inserts the cells of `range` on `sheet_index` (Insert Cells), shifting
  /// the cells of the band — the range's columns for `kDown`, its rows for
  /// `kRight` — from the range's edge onwards by the range's height or width.
  /// A reference, name target or hyperlink moves (or grows across the edge)
  /// only when its perpendicular extent lies wholly inside the band; one
  /// straddling the band, a whole column and a whole row are left as they
  /// are. A CF / DV sqref straddling the moving region splits into its moved
  /// and unmoved parts; a merge straddling it is unmerged first. Inserted
  /// cells take the XF of the cell directly above (`kDown`) or left
  /// (`kRight`). Row/column layout, page breaks, drawing anchors and the
  /// sheet view do not move. A range covering the whole perpendicular extent
  /// is the whole row/column insert. A `kDown` insert over the blank
  /// placeholder row of a header-only table (`insert_row`) pushes the cells
  /// below and leaves the table's range as it was, clearing the mark.
  ///
  /// Errors, each leaving the workbook unchanged:
  ///   * `kInvalidArgument` for an out-of-range sheet index, a range outside
  ///     the grid, `kUp` / `kLeft`, or a whole-column range shifted down (a
  ///     whole-row range shifted right).
  ///   * `kPreconditionFailed` when the moving region meets a table, a
  ///     multi-cell array formula or a data table its band does not wholly
  ///     cover, when the edit cuts into a PivotTable, meets one across a band
  ///     edge or lies wholly inside one, when the shift path holds a
  ///     PivotTable at all, when the band cuts into a table without spanning
  ///     exactly the table's columns (`kDown`) or rows (`kRight`), totals row
  ///     included, or when a value, formula, note, fill, border or merge would
  ///     be pushed off the grid.
  Expected<void, Error> insert_cells(std::size_t sheet_index, const MergeRange& range, CellShift shift);

  /// Deletes the cells of `range` on `sheet_index` (Delete Cells), shifting
  /// the band's cells past the range back by its height (`kUp`) or width
  /// (`kLeft`). References follow `insert_cells`' band rule; a wholly deleted
  /// one becomes `#REF!`, a partly deleted one shrinks, and a range whose
  /// axial extent lies in the deleted interval and that the band touches at
  /// one perpendicular end shrinks to its part outside the band. Notes and
  /// threaded comments go with their deleted cells; a merge partly inside
  /// the deleted rectangle is unmerged.
  ///
  /// Errors, each leaving the workbook unchanged:
  ///   * `kInvalidArgument` as `insert_cells`, with `kDown` / `kRight` the
  ///     refused shifts.
  ///   * `kPreconditionFailed` on the same objects as `insert_cells`; a
  ///     delete inside a PivotTable is refused rather than hiding an item.
  ///     A `kUp` delete spanning a table's columns follows `delete_rows` for
  ///     its header, totals and data rows, keeping only the table's columns
  ///     of the kept data row.
  Expected<void, Error> delete_cells(std::size_t sheet_index, const MergeRange& range, CellShift shift);

 private:
  Workbook();

  // `set_cell_value` without the array / data-table write refusals.
  void put_cell_value(std::size_t sheet_index, std::uint32_t row, std::uint32_t col, Value value);
  void put_cell_text(std::size_t sheet_index, std::uint32_t row, std::uint32_t col, std::string_view text);
  // The `put_cell_*` writes for a caller already holding the engine mutex.
  void put_cell_value_locked(std::size_t sheet_index, std::uint32_t row, std::uint32_t col, Value value);
  void put_cell_text_locked(std::size_t sheet_index, std::uint32_t row, std::uint32_t col, std::string_view text);
  // `set_table_header_row` after its argument checks, without the row visibility the filter left behind.
  Expected<void, Error> set_table_header_row_cells(std::size_t index, bool shown, const MergeRange& rect);
  void put_cell_formula_locked(std::size_t sheet_index, std::uint32_t row, std::uint32_t col, std::string formula,
                               FormulaSource source, std::uint32_t block_rows, std::uint32_t block_cols);
  // `write_table_totals_cell` for a caller already holding the engine mutex.
  Expected<void, Error> write_table_totals_cell_locked(std::size_t index, std::size_t column);
  // `insert_cells` (`is_delete` false) or `delete_cells` for a caller already
  // holding the engine mutex.
  Expected<void, Error> apply_band_edit_locked(std::size_t sheet_index, const MergeRange& range, CellShift shift,
                                               bool is_delete);
  // Runs a row-axis delete (whole rows or a band shifting up) under the
  // engine mutex, keeping the first data row of each table it would empty.
  Expected<void, Error> apply_row_delete_locked(std::size_t sheet_index, const StructuralEdit& edit);
  // Replaces the totals settings of table `index`'s columns with the defaults of a never-shown totals row.
  void apply_totals_defaults(std::size_t index);
  // Resets every cell of each rect to cell format 0 (explicit), the format a
  // removed hyperlink leaves behind.
  Expected<void, Error> reset_hyperlink_ranges(std::size_t sheet_index, const std::vector<StyleRect>& rects);

  // `add_merge` without the overlap refusals, for a PivotTable writing its
  // own report; `merge` must be normalized and inside the grid.
  void put_merge(std::size_t sheet_index, const MergeRange& merge);
  // Drops every stored cell of `rect`, marking what read them dirty.
  void erase_cells_in(std::size_t sheet_index, const MergeRange& rect);
  // `set_cell_formula` without the write refusals; a non-zero `block_rows` x
  // `block_cols` stores the cell as the anchor of a legacy array block.
  void put_cell_formula(std::size_t sheet_index, std::uint32_t row, std::uint32_t col, std::string formula,
                        FormulaSource source, std::uint32_t block_rows, std::uint32_t block_cols);
  // Whether the dirty set is non-empty, read under the engine mutex.
  bool has_dirty_cells() const;
  // Whether any sheet holds a data table, read under the engine mutex.
  bool has_data_tables() const;
  // The data-table step of `recalc()` / `recalc_parallel()`.
  Expected<void, Error> recalc_data_tables_after_pass(bool evaluated_any, const eval::FunctionRegistry& registry);
  Expected<void, Error> compute_data_tables(const eval::FunctionRegistry& registry);

  // `project_pivot`'s blocked state under `kBlock`: clears the report's cells and merges, leaves `#SPILL!` in the
  // anchor's existing format and records `plan`'s size relative to the anchor.
  Expected<void, Error> withhold_pivot_report(std::size_t sheet_index, std::size_t pivot_index,
                                              const detail::PivotReportPlan& plan);
  // Deletes every table on `sheet_index` meeting `rect`, as a paste over it does.
  void remove_tables_meeting(std::size_t sheet_index, const MergeRange& rect);

  // `move_sheet` and `remove_sheet` past their argument checks, for a caller
  // already holding the engine mutex.
  void move_sheet_locked(std::uint32_t from_index, std::uint32_t to_index);
  Expected<void, Error> remove_sheet_locked(std::uint32_t index);
  // The name and capacity checks `add_sheet_entered` applies to a new sheet.
  Expected<void, Error> check_entered_new_sheet_name(std::string_view name) const;
  Expected<void, Error> check_sheet_capacity() const;

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
  // See `allocate_runtime_id()`.
  mutable std::uint32_t next_runtime_id_ = 1;
  // Generator state behind `add_threaded_comment`'s GUIDs; seeded on first use.
  std::optional<std::uint64_t> guid_seed_;
  std::uint64_t guid_sequence_ = 0;
  std::string next_threaded_comment_guid();
  // Gives every mention with an empty id a generated GUID.
  void fill_mention_ids(std::vector<Mention>& mentions);
  std::vector<PassthroughPart> passthrough_parts_;
  std::vector<ExternalLinkRecord> external_links_;
  // See `location()`.
  std::optional<WorkbookLocation> location_;
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
  CalcOptions calc_options_;
  // A data table was created or skipped since the last table pass.
  bool data_tables_pending_ = false;
  // 1904 date system flag parsed from `<workbookPr date1904>`. Default
  // false (1900 system).
  bool date1904_ = false;
  // Most callers receive parts from a reader that already sorted them by
  // package path. Keep that fact across the move-only workbook operations so
  // the common insertion path can use one lower-bound search instead of
  // rescanning and sorting every retained part after each append.
  bool passthrough_parts_sorted_ = true;
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
  WorksheetRawExtensions raw_children_;
  std::vector<std::uint8_t> xlsb_book_frt_;
  // Extra `<workbook>` root namespace declarations (xmlns:mc / xmlns:xr* /
  // x15 + mc:Ignorable) captured for verbatim re-emission so namespaced
  // attributes in the raw captures above resolve. Empty by default.
  std::string workbook_root_extra_attrs_;
  // Formula compatibility profile. Runtime default is `win-365-en_US`.
  ExcelProfile excel_profile_ = default_excel_profile();
  // Set by `create_for_profile` and kept across `set_excel_profile`: sheets
  // added later receive the same sheet defaults as the first one. Loaded
  // and plain workbooks keep bare new sheets, since their Normal font need
  // not be the profile's.
  std::optional<ExcelProfile> new_sheet_profile_;
  const ExcelProfile* new_sheet_profile() const noexcept {
    return new_sheet_profile_.has_value() ? &*new_sheet_profile_ : nullptr;
  }
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
