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
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include "cell.h"
#include "io/unknown_relationship.h"
#include "value.h"

namespace formulon {

namespace cf {
struct ConditionalFormat;
}  // namespace cf

namespace pivot {
class PivotTable;
}  // namespace pivot

/// One `<mergeCell ref="A1:B2"/>` block: a rectangular merged range
/// stored as 0-based, inclusive `(row, col)` corners. The two corners
/// are normalised so that `first_row <= last_row` and
/// `first_col <= last_col`; degenerate `first == last` rectangles are
/// permitted (single-cell merges, which Excel emits).
struct MergeRange {
  std::uint32_t first_row = 0;
  std::uint32_t first_col = 0;
  std::uint32_t last_row = 0;
  std::uint32_t last_col = 0;
};

/// One `<hyperlink>` entry attached to a sheet. The numeric rectangle is the
/// sole source of truth; the OOXML writer regenerates its `ref` attribute from
/// these four 0-based inclusive endpoints. The OOXML wire form is
///
///     <hyperlink ref="A1" r:id="rId3" tooltip="..." display="..."/>
///
/// where `r:id` resolves to a `Target` in the sheet's
/// `_rels/<sheet>.xml.rels`. Hyperlinks may be external (URLs / mailto)
/// or internal (e.g. `#Sheet2!A1`); for internal links Excel emits the
/// `location` attribute on the `<hyperlink>` itself rather than going
/// through rels. We capture both shapes in a single struct so the
/// writer can reproduce either.
///
/// Lifetime: `target` / `display` / `tooltip` / `location` are owned
/// `std::string`s. `rid` is the relationship id observed at read time;
/// the writer reuses it verbatim for round-trip stability and assigns a
/// fresh `rId<N>` when the field is empty (newly added entries).
struct Hyperlink {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::uint32_t last_row = 0;
  std::uint32_t last_col = 0;
  std::string target;    ///< External URL / mailto / file path. Empty for purely internal links.
  std::string location;  ///< `Sheet2!A1` style internal target, or empty.
  std::string display;   ///< Optional `display="..."` attribute.
  std::string tooltip;   ///< Optional `tooltip="..."` attribute.
  std::string rid;       ///< Reader populates from sheet rels; empty for fresh entries.
};

/// One per-cell text comment (`<comment ref="A1" authorId="N"><text>...</text></comment>`).
///
/// The struct is intentionally narrow: a single plain-text payload, no
/// rich-text runs. Reads of multi-run rich text concatenate the runs
/// into `text`; writes always emit the plain form. Authors are
/// preserved verbatim so the workbook-wide author list round-trips.
struct CellComment {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::string author;
  std::string text;
};

/// One `<dataValidation>` block: a list of cell ranges that share a
/// validation rule, plus the rule itself. Storage is plain POD; the
/// engine does not yet evaluate the rules — the data is captured for
/// round-trip and surfacing through the binding APIs.
///
/// Field semantics (matches OOXML `dataValidations.xsd`):
///   * `type`         — 0 none, 1 whole, 2 decimal, 3 list, 4 date,
///                      5 time, 6 textLength, 7 custom.
///   * `op`           — 0 between, 1 notBetween, 2 equal, 3 notEqual,
///                      4 greaterThan, 5 lessThan, 6 greaterThanOrEqual,
///                      7 lessThanOrEqual.
///   * `error_style`  — 0 stop, 1 warning, 2 information.
///   * `formula1`/`formula2` — raw OOXML formula text without leading
///                      `=`; the evaluator does not consume these.
struct DataValidation {
  std::vector<MergeRange> ranges;
  std::uint8_t type = 0;
  std::uint8_t op = 0;
  std::uint8_t error_style = 0;
  bool allow_blank = true;
  bool show_input_message = false;
  bool show_error_message = false;
  // Whether the in-cell dropdown arrow is shown for a `list` validation.
  // Default matches Excel's default (arrow shown). Note the OOXML
  // `showDropDown` attribute has inverted semantics per ECMA-376: its
  // presence with value "1" SUPPRESSES the arrow. This field holds the
  // user-facing meaning (arrow shown/hidden); the reader/writer invert it
  // when translating to/from the raw XML attribute.
  bool show_dropdown = true;
  std::string formula1;
  std::string formula2;
  std::string error_title;
  std::string error_message;
  std::string prompt_title;
  std::string prompt_message;
};

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

/// The geometry of a dynamic-array spill footprint. It stores no values and
/// is safe to copy while the sheet spill mutex is held. The same value type
/// is used for committed and blocked spill snapshots.
struct SpillFootprint {
  std::uint32_t anchor_row = 0;
  std::uint32_t anchor_col = 0;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
};

/// Backwards-compatible name for the blocked-spill reverse-index record.
using BlockedSpillFootprint = SpillFootprint;

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
  // `<sheetView>` display attributes. Names mirror the OOXML attributes;
  // the values shown are the ECMA-376 §18.3.1.87 schema defaults, so a
  // freshly created sheet round-trips without emitting them. Losing these
  // on save silently re-shows hidden gridlines/headers, drops the active
  // tab, and resets the view mode.
  bool show_grid_lines = true;       ///< `showGridLines` (default true).
  bool show_row_col_headers = true;  ///< `showRowColHeaders` (default true).
  bool show_zeros = true;            ///< `showZeros` (default true).
  bool right_to_left = false;        ///< `rightToLeft` (default false).
  bool tab_selected = false;         ///< `tabSelected` (default false).
  /// `view` mode: empty (== "normal"), "pageBreakPreview", or
  /// "pageLayout". Stored verbatim so unknown future values round-trip.
  std::string view_mode;
};

/// Mirror of OOXML `<sheetProtection>` (ECMA-376 §18.3.1.85). Stored as
/// passive round-trip metadata: the engine does not enforce locks at
/// evaluation time. The host UI inspects these flags to mirror Excel's
/// "Protect Sheet" dialog state.
///
/// Field semantics follow the OOXML attribute names verbatim. Each
/// boolean represents the attribute's value as it appears in the
/// document; `false` means "attribute absent or `0`". Each attribute's
/// effect (whether `true` blocks or allows the operation) is dictated
/// by the spec, not inverted here.
///
/// `enabled` is the only synthetic field: it tracks whether the
/// `<sheetProtection>` element is present at all. A workbook with
/// `enabled == false` round-trips with no element emitted, regardless
/// of the other field values.
struct SheetProtection {
  /// True when the `<sheetProtection>` element is present.
  bool enabled = false;
  /// Hash algorithm name (e.g. "SHA-512"). Empty when no modern
  /// password is set.
  std::string algorithm_name;
  /// Base64-encoded password hash. Empty when no modern password is
  /// set.
  std::string hash_value;
  /// Base64-encoded salt for the password hash.
  std::string salt_value;
  /// Iteration count for the password hash.
  std::uint32_t spin_count = 0;
  /// Legacy 16-bit hashed password, written as a hex string (e.g.
  /// `"CC53"`). Empty when no legacy password is set. Excel writes
  /// either the legacy or the modern password attributes, not both.
  std::string legacy_password;

  // OOXML attribute flags. Names mirror the attributes verbatim; see
  // ECMA-376 §18.3.1.85 for per-attribute defaults and effects. A flag of
  // `true` means the corresponding action is LOCKED. Eleven of these
  // (format*, insert*, delete*, sort, autoFilter, pivotTables) default to
  // `true` in the schema, so the member initialisers match those defaults:
  // a default-constructed protection that is then `enabled` mirrors Excel's
  // "Protect Sheet" defaults. The five action flags whose schema default is
  // `false` (sheet, objects, scenarios, select*) stay false.
  bool sheet = false;
  bool objects = false;
  bool scenarios = false;
  bool format_cells = true;
  bool format_columns = true;
  bool format_rows = true;
  bool insert_columns = true;
  bool insert_rows = true;
  bool insert_hyperlinks = true;
  bool delete_columns = true;
  bool delete_rows = true;
  bool select_locked_cells = false;
  bool select_unlocked_cells = false;
  bool sort = true;
  bool auto_filter = true;
  bool pivot_tables = true;
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
  // Presence is distinct from the value: `<col width="0">` is an
  // explicit zero-width override, while an absent width uses the sheet
  // default. The style selector follows the same rule, including an
  // explicit `style="0"`.
  bool has_width = false;
  bool has_style = false;
  std::uint32_t style_xf = 0;
};

/// Returns whether a column has a logically explicit width.
///
/// `has_width` is the lossless source-presence bit, but callers may construct
/// the aggregate model directly using the legacy convention that a non-zero
/// width is explicit. Keep that convention observable at every output/API
/// boundary while retaining an explicit zero through `has_width`.
inline bool HasExplicitColumnWidth(const ColumnLayout& column) {
  return column.has_width || column.width != 0.0;
}

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
  bool has_height = false;
  // A row style is effective only when OOXML `customFormat="1"` is
  // present. `style_xf == 0` remains meaningful when `s` is absent or
  // explicitly set to zero.
  bool has_style = false;
  std::uint32_t style_xf = 0;
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

/// OOXML spec default values for worksheet print/format attributes.
///
/// These mirror the `<defaultValue>` entries in the ECMA-376 schema for the
/// `<sheetFormatPr>`, `<pageSetup>` and `<pageMargins>` elements. They are
/// named so the struct member initializers below carry no bare literals.
namespace ooxml_defaults {

/// `<sheetFormatPr baseColWidth>` default (characters).
inline constexpr double kBaseColWidthChars = 8.0;

/// `<pageSetup paperSize>` default (9 = A4).
inline constexpr std::uint32_t kPaperSize = 9;

/// `<pageSetup scale>` default (percentage).
inline constexpr std::uint32_t kPageScalePercent = 100;

/// `<pageMargins left>` / `<pageMargins right>` default (inches).
inline constexpr double kPageMarginSideInches = 0.7;

/// `<pageMargins top>` / `<pageMargins bottom>` default (inches).
inline constexpr double kPageMarginTopBottomInches = 0.75;

/// `<pageMargins header>` / `<pageMargins footer>` default (inches).
inline constexpr double kPageMarginHeaderFooterInches = 0.3;

}  // namespace ooxml_defaults

/// Worksheet default column/row metrics from the `<sheetFormatPr>` element.
///
/// `<sheetFormatPr>` carries the metrics Excel applies to columns and rows
/// that have no explicit `<col>` / `<row>` override. The pagination engine
/// needs these defaults to size un-overridden tracks. The `has_*` flags
/// record whether the source attribute was present, so a consumer can tell
/// an explicit `0` from an absent attribute. The field names declared here
/// are stable.
struct SheetFormatDefaults {
  double default_col_width = 0.0;   ///< `defaultColWidth`, in OOXML character-width units.
  double default_row_height = 0.0;  ///< `defaultRowHeight`, in points.
  /// `baseColWidth`, in characters (OOXML spec default 8).
  double base_col_width = ooxml_defaults::kBaseColWidthChars;
  bool has_default_col_width = false;   ///< True when `defaultColWidth` was present.
  bool has_default_row_height = false;  ///< True when `defaultRowHeight` was present.
};

/// A single manual page break.
///
/// Mirrors the OOXML `<brk id="..." min="..." max="..." man="..."/>` element
/// inside `<rowBreaks>` / `<colBreaks>`. `id` is the 0-based row or column
/// index the break sits *before*, which is what OOXML stores too -- Excel
/// writes `id="20"` for a break placed before row 21. `min` / `max` bound
/// the span the break applies to.
struct ManualBreak {
  std::uint32_t id = 0;   ///< 0-based row/column index the break precedes.
  std::uint32_t min = 0;  ///< Span start (0-based).
  std::uint32_t max = 0;  ///< Span end (0-based).
  /// True for a user-placed break (`man="1"`). ECMA-376 §18.3.1.1 defaults
  /// `man` to false: a `<brk>` without the attribute is an automatic break
  /// that the pagination engine must not treat as a forced boundary.
  bool manual = false;
};

/// Page orientation as stored in `<pageSetup orientation="...">`.
enum class Orientation { kDefault, kPortrait, kLandscape };

/// Structured page setup parsed from `<pageSetup>` and `<sheetPr>`.
///
/// Parsed alongside `SheetPrintSettings::page_setup_xml`; the raw string
/// remains the source of truth for the writer. These fields exist so the
/// pagination engine can reason about orientation, paper size and scaling
/// without re-parsing XML. Missing attributes keep the defaults shown.
struct PageSetup {
  Orientation orientation = Orientation::kDefault;  ///< `orientation` attribute.
  /// `paperSize`; OOXML default 9 (A4).
  std::uint32_t paper_size = ooxml_defaults::kPaperSize;
  /// `scale`, as a percentage.
  std::uint32_t scale = ooxml_defaults::kPageScalePercent;
  std::uint32_t fit_to_width = 1;   ///< `fitToWidth`, in pages.
  std::uint32_t fit_to_height = 1;  ///< `fitToHeight`, in pages.
  bool fit_to_page = false;         ///< True when `<sheetPr><pageSetUpPr fitToPage="1"/>`.
};

/// Structured page margins parsed from `<pageMargins>`.
///
/// Parsed alongside `SheetPrintSettings::page_margins_xml`; the raw string
/// remains the source of truth for the writer. All values are in inches.
/// The defaults match the OOXML spec defaults.
struct PageMargins {
  double left = ooxml_defaults::kPageMarginSideInches;            ///< Left margin, inches.
  double right = ooxml_defaults::kPageMarginSideInches;           ///< Right margin, inches.
  double top = ooxml_defaults::kPageMarginTopBottomInches;        ///< Top margin, inches.
  double bottom = ooxml_defaults::kPageMarginTopBottomInches;     ///< Bottom margin, inches.
  double header = ooxml_defaults::kPageMarginHeaderFooterInches;  ///< Header margin, inches.
  double footer = ooxml_defaults::kPageMarginHeaderFooterInches;  ///< Footer margin, inches.
};

/// Passive round-trip storage for worksheet print/page configuration.
///
/// The page setup surface has many Excel-specific attributes and may point
/// at a binary `printerSettings*.bin` part through `r:id`. The XML fragments
/// stay raw so a read/save cycle preserves user-authored print settings
/// verbatim; the structured `page_setup` / `page_margins` / break vectors
/// are parsed *alongside* the raw strings for consumers (such as the
/// pagination engine) that need typed access.
struct SheetPrintSettings {
  std::string sheet_pr_xml;       ///< Raw `<sheetPr>` when it carries page setup metadata.
  std::string page_margins_xml;   ///< Raw `<pageMargins .../>`.
  std::string page_setup_xml;     ///< Raw `<pageSetup .../>`.
  std::string print_options_xml;  ///< Raw `<printOptions .../>` (gridlines/headings printing, centring).
  std::string header_footer_xml;  ///< Raw `<headerFooter>...</headerFooter>` (odd/even/first page strings).
  std::string printer_settings_rid;
  std::string printer_settings_path;  ///< Package path, e.g. `xl/printerSettings/printerSettings1.bin`.

  PageSetup page_setup;      ///< Structured view of `<pageSetup>` + `<pageSetUpPr>`.
  PageMargins page_margins;  ///< Structured view of `<pageMargins>`.

  std::vector<ManualBreak> manual_row_breaks;  ///< `<rowBreaks>` entries, 0-based ids.
  std::vector<ManualBreak> manual_col_breaks;  ///< `<colBreaks>` entries, 0-based ids.
};

/// Worksheet children which the calculation model does not interpret but
/// which must retain their schema position to avoid silently damaging an
/// Excel-authored sheet. Each member is a complete XML element captured
/// verbatim from the source worksheet.
struct WorksheetRawExtensions {
  std::string protected_ranges_xml;
  std::string scenarios_xml;
  std::string custom_sheet_views_xml;
  std::string phonetic_pr_xml;
  std::string ignored_errors_xml;
  std::string legacy_drawing_hf_xml;
  std::string picture_xml;
  std::string ole_objects_xml;
  std::string controls_xml;
};

/// Worksheet records captured verbatim from an `.xlsb` sheet part that the
/// calculation model does not express: conditional formatting, data
/// validation, hyperlinks, auto-filter, print setup, manual breaks, and the
/// drawing / table part references.
///
/// The XML path keeps such content as unknown *parts* or as raw element
/// strings (`WorksheetRawExtensions`), but a binary sheet part is a single
/// record stream the reader consumes whole, so neither mechanism reaches it.
/// Retaining the framed record bytes is the binary equivalent.
///
/// The three buffers represent the grammar slots before merges, after merges
/// and before hyperlinks, and after hyperlinks. The merge and hyperlink
/// blocks are the only tail constructs the model owns and re-emits itself;
/// keeping these slots preserves source order even when a source omits one of
/// those blocks.
///
/// Retention is byte-verbatim, with the same consequences the raw-XML
/// retention already carries: a row/column edit does not remap coordinates
/// inside these records, and a retained formula blob's sheet-qualified
/// references resolve through the regenerated `BrtExternSheet` table rather
/// than the source one.
struct XlsbSheetTail {
  /// Records between the end of the cell table and the merged-cell block.
  std::vector<std::uint8_t> before_merges;
  /// Records after the merged-cell block and before the model-owned
  /// `BrtHLink` block. Raw `BrtHLink` records are decoded into
  /// `Sheet::hyperlinks()` and are never retained here.
  std::vector<std::uint8_t> after_merges_before_hyperlinks;
  /// Records after the model-owned `BrtHLink` block, up to `BrtEndSheet`.
  std::vector<std::uint8_t> after_hyperlinks;

  bool empty() const noexcept {
    return before_merges.empty() && after_merges_before_hyperlinks.empty() && after_hyperlinks.empty();
  }
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

/// One populated row's cells, stored as a contiguous run that begins at
/// `first_col()` instead of at column 0.
///
/// A row's populated columns are usually clustered, but the cluster is not
/// anchored at column A: a sheet with data in `A:H` and one cell in `CZ` is
/// ordinary, and so is a row whose only content sits far to the right. A
/// vector indexed by absolute column materialises every slot from 0, which
/// makes a sheet's memory scale with its used *width* rather than with its
/// content — at 88 bytes per `Cell` a single far-right cell costs kilobytes
/// per row. Carrying the run's own origin removes the leading gap.
///
/// Reads stay in absolute coordinates. `size()` is one past the highest
/// stored column and `operator[]` accepts any column below it, answering the
/// leading gap with a shared blank cell, so a consumer that walks
/// `0 .. size()` and skips blanks sees exactly what a dense vector gave it.
/// `find()` (and through it `Sheet::cell_at`) instead reports the gap as
/// absent, because those slots were never written.
class RowCells {
 public:
  /// Absolute column of the first materialised slot. Meaningless when the
  /// run is empty.
  std::uint32_t first_col() const noexcept { return first_col_; }

  /// One past the highest materialised column, in absolute coordinates.
  std::size_t size() const noexcept { return run_.empty() ? 0U : static_cast<std::size_t>(first_col_) + run_.size(); }

  bool empty() const noexcept { return run_.empty(); }

  /// Number of slots actually held in memory, i.e. `size()` minus the
  /// unmaterialised leading gap.
  std::size_t stored_count() const noexcept { return run_.size(); }

  /// The cell at absolute column `col`. Columns in the leading gap, and
  /// columns past the run, read as a shared blank cell.
  const Cell& operator[](std::size_t col) const noexcept {
    if (col < first_col_) {
      return blank();
    }
    const std::size_t index = col - first_col_;
    return index < run_.size() ? run_[index] : blank();
  }

  /// The materialised slot at `col`, or `nullptr` when that column was never
  /// written.
  const Cell* find(std::uint32_t col) const noexcept {
    if (run_.empty() || col < first_col_) {
      return nullptr;
    }
    const std::size_t index = static_cast<std::size_t>(col) - first_col_;
    return index < run_.size() ? &run_[index] : nullptr;
  }
  Cell* find(std::uint32_t col) noexcept { return const_cast<Cell*>(static_cast<const RowCells*>(this)->find(col)); }

  /// Materialises `col` — extending the run in either direction — and returns
  /// its slot. Existing slots keep their heap-stable `cached_text_owned`
  /// payload across the move.
  Cell& ensure(std::uint32_t col);

  /// Iteration over materialised slots only, in ascending column order.
  std::vector<Cell>::const_iterator begin() const noexcept { return run_.begin(); }
  std::vector<Cell>::const_iterator end() const noexcept { return run_.end(); }

  /// Direct access to the run for the structural-edit paths, which rewrite
  /// both the storage and its origin together.
  const std::vector<Cell>& run() const noexcept { return run_; }
  std::vector<Cell>& mutable_run() noexcept { return run_; }
  void set_first_col(std::uint32_t col) noexcept { first_col_ = col; }

 private:
  static const Cell& blank() noexcept;

  std::uint32_t first_col_ = 0;
  std::vector<Cell> run_;
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

  /// Inclusive bounding box of populated coordinates returned by
  /// `populated_extent`. The fields are copies; no sheet storage is exposed
  /// and the sheet lock may be released before the caller consumes them.
  struct PopulatedExtent {
    std::uint32_t first_row = 0;
    std::uint32_t first_col = 0;
    std::uint32_t last_row = 0;
    std::uint32_t last_col = 0;
  };

  /// True when `(row, col)` addresses a cell inside the Excel grid.
  /// The single predicate every public coordinate entry point (C ABI
  /// setters, pivot/CF mutators) checks before handing an
  /// attacker-controlled `std::uint32_t` to the storage layer, whose own
  /// bounds checks are debug-only asserts.
  static constexpr bool coord_in_grid(std::uint32_t row, std::uint32_t col) noexcept {
    return row < kMaxRows && col < kMaxCols;
  }

  /// True when `[first_row..last_row] x [first_col..last_col]` is a
  /// well-ordered rectangle fully inside the Excel grid (both axes
  /// non-inverted and both corners in-grid).
  static constexpr bool rect_in_grid(std::uint32_t first_row, std::uint32_t first_col, std::uint32_t last_row,
                                     std::uint32_t last_col) noexcept {
    return first_row <= last_row && first_col <= last_col && coord_in_grid(last_row, last_col);
  }

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

  /// Whether this is an OOXML sheet type that Formulon intentionally treats
  /// as opaque (chartsheet, dialogsheet, macrosheet, or a future sheet
  /// type). Its XML part is carried through `Workbook::passthrough_parts()`
  /// while this lightweight Sheet preserves its place, name and visibility
  /// in the workbook's ordered `<sheets>` list.
  bool is_opaque_ooxml_sheet() const noexcept { return !opaque_ooxml_part_path_.empty(); }

  /// Marks this sheet as opaque OOXML metadata. `part_path` is the package
  /// path (for example `xl/chartsheets/sheet1.xml`) and `relationship_type`
  /// is the original workbook-relationship type URI.
  void set_opaque_ooxml_sheet(std::string part_path, std::string relationship_type) {
    opaque_ooxml_part_path_ = std::move(part_path);
    opaque_ooxml_relationship_type_ = std::move(relationship_type);
  }

  const std::string& opaque_ooxml_part_path() const noexcept { return opaque_ooxml_part_path_; }
  const std::string& opaque_ooxml_relationship_type() const noexcept { return opaque_ooxml_relationship_type_; }

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

  /// Stores a text literal while copying its bytes into the cell's own
  /// heap-stable storage. This is for callers whose input buffer is only
  /// valid for the duration of the call (such as the C ABI); readers with a
  /// workbook-scoped shared-string store should use `set_cell_value`.
  /// Growth and spill-invalidation semantics match `set_cell_value`.
  void set_cell_text(std::uint32_t row, std::uint32_t col, std::string_view text);

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

  /// Updates a cached value without copying Text bytes. The caller must
  /// provide Text storage that outlives this Sheet and must not pass a view
  /// backed by this cell's previous `cached_text_owned` allocation.
  ///
  /// OOXML/XLSB readers use this for workbook-owned shared-string storage;
  /// evaluator results must use `set_cell_cached_value` instead.
  void set_cell_cached_value_borrowed(std::uint32_t row, std::uint32_t col, Value v);

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

  /// Total number of materialised `Cell` slots across all populated rows.
  ///
  /// Counts every slot a row actually holds, including the implicitly
  /// default-constructed ones a write to a later column creates between the
  /// row's first and last populated column. Columns before a row's first
  /// populated one are not materialised and are not counted. Useful for tests
  /// and as a coarse memory-footprint indicator.
  std::size_t cell_count() const noexcept;

  /// Returns the inclusive bounding box of every populated coordinate that
  /// intersects `[first_row..last_row] x [first_col..last_col]`, or
  /// `std::nullopt` when the query is invalid, empty, or contains no content.
  ///
  /// A stored coordinate is populated when its cell has formula text or a
  /// non-blank cached value. A committed dynamic-array spill contributes its
  /// whole rectangle (clipped to the query), without enumerating phantom
  /// cells. Blocked spill footprints are intentionally not included. The
  /// complete read takes `spill_mutex_` once, so the returned copy is an
  /// atomic snapshot relative to cell and spill mutations; no row, cell, or
  /// spill-table references escape the lock.
  std::optional<PopulatedExtent> populated_extent(std::uint32_t first_row, std::uint32_t first_col,
                                                  std::uint32_t last_row, std::uint32_t last_col) const noexcept;

  /// Monotonically changes whenever the flat stored-cell / spill-phantom
  /// address set changes. C-ABI iteration uses it to invalidate its
  /// per-handle sorted-address cache after any sheet mutation.
  std::uint64_t cell_enumeration_revision() const noexcept { return cell_enumeration_revision_; }

  /// Read-only access to the underlying row map.
  ///
  /// Exposed so consumers (e.g. the OOXML writer) can iterate populated
  /// rows in their own order without paying for an intermediate copy. The
  /// reference is invalidated by mutating Sheet operations.
  const std::unordered_map<std::uint32_t, RowCells>& rows() const noexcept { return rows_; }

  // ---------------------------------------------------------------------------
  // Dynamic-array spill API
  // ---------------------------------------------------------------------------

  /// Returns whether a spill footprint would collide with this sheet's
  /// current contents. The predicate is conservative: zero-sized,
  /// out-of-grid, or overflowing footprints collide. Only the stored Cell at
  /// the requested anchor and any spill region currently anchored there are
  /// ignored so callers can re-evaluate an existing producer without
  /// self-collision. Other stored cells, spill rectangles (including their
  /// anchors and phantoms), and *any* merged range intersecting the footprint
  /// collide; a merge at the requested anchor is not exempt. The sheet lock
  /// is held for the complete read.
  bool spill_would_collide(std::uint32_t anchor_row, std::uint32_t anchor_col, std::uint32_t rows,
                           std::uint32_t cols) const noexcept;

  /// Why a spill footprint would be refused, or that it would be accepted.
  /// Both refusals surface as `#SPILL!` in a cell; they are reported apart
  /// so a caller can tell "the rectangle leaves the grid" from "something
  /// occupies it" without re-deriving either.
  enum class SpillAdmission : std::uint8_t {
    kAdmissible,
    kOutsideGrid,
    kBlocked,
  };

  /// Decides whether a `rows` x `cols` footprint anchored at `(anchor_row,
  /// anchor_col)` could commit, **without building or committing any
  /// values**.
  ///
  /// This is the admission half of `commit_spill` on its own, so a producer
  /// whose result would be refused can learn that before materialising it.
  /// That matters for a footprint the size of a grid axis: a whole-column
  /// rectangle is 1,048,576 cells, and a caller that had to materialise one
  /// to discover it is blocked would pay ~25 MB and a million writes for an
  /// answer that is `#SPILL!`.
  ///
  /// The scan is proportional to what the sheet actually stores inside the
  /// rectangle, not to the rectangle's area: it walks the stored rows or
  /// probes the rectangle's rows, whichever set is smaller, and within a row
  /// only the materialised run intersected with the column span. A blocker
  /// far outside the populated data — the case that shows the footprint is
  /// the full declared rectangle rather than the used range — is found at
  /// the cost of the cells that exist, not the cells the rectangle spans.
  ///
  /// `scan_steps` (when non-null) receives the work the admission scan did:
  /// one step per row visited or probed, plus one per cell inspected. It
  /// exists so that proportionality is testable — an area-proportional
  /// implementation returns the same verdict, only slower, so nothing but a
  /// count distinguishes them. Counting inspected cells alone would not: a
  /// scan that probes a million empty rows inspects almost nothing while
  /// doing a million lookups, so the row visits have to be counted too.
  ///
  /// Pure: no cached state is built or updated, and the sheet is unchanged.
  /// The only shared state touched is `spill_mutex_`, held for the read.
  SpillAdmission probe_spill_footprint(std::uint32_t anchor_row, std::uint32_t anchor_col, std::uint32_t rows,
                                       std::uint32_t cols, std::uint64_t* scan_steps = nullptr) const noexcept;

  /// Records a refused footprint: writes `#SPILL!` to the anchor cell and
  /// remembers the rectangle so the anchor is retried once whatever occupies
  /// it goes away.
  ///
  /// Callers that refuse a spill on their own — because `probe_spill_footprint`
  /// told them to, and they skipped materialising the values — must route the
  /// refusal through here rather than returning `#SPILL!` directly. Skipping
  /// the registration leaves the anchor with no remembered rectangle, and the
  /// blocked-spill release machinery then has nothing to retry when the
  /// blocker is cleared. `commit_spill` refuses through this same body, so
  /// the two paths cannot drift.
  void reject_spill_footprint(std::uint32_t anchor_row, std::uint32_t anchor_col, std::uint32_t rows,
                              std::uint32_t cols);

  /// Returns the spill region anchored at `(row, col)`, or `nullptr` when
  /// no region is anchored there. The returned pointer is valid until the
  /// next mutating call to the spill API (`commit_spill`, `clear_spill`)
  /// or to a cell-mutating call that triggers eager invalidation.
  const SpillRegion* spill_region_at_anchor(std::uint32_t row, std::uint32_t col) const noexcept;

  /// Returns the spill region whose phantom area covers `(row, col)`, or
  /// `nullptr` when no region covers it. Returns `nullptr` for the
  /// region's anchor cell itself: only phantom coordinates match. Lookup
  /// scans the registered spill rectangles, avoiding a per-phantom index.
  /// Use `spill_region_at_anchor` to look up by anchor.
  const SpillRegion* spill_region_covering(std::uint32_t row, std::uint32_t col) const noexcept;

  /// Returns the coordinates of every phantom cell across all registered
  /// spill regions on this sheet. Each region's anchor is excluded because
  /// it already occupies a real slot in the cell store; only the phantoms
  /// (which live solely in the spill table and are absent from `rows()`)
  /// are returned. The result is unordered — callers that need a
  /// deterministic order must sort it — and empty when the sheet hosts no
  /// spill regions.
  ///
  /// Exposed so flat-enumeration consumers (cell iteration, used-range
  /// computation) can surface spill phantoms alongside the stored cells;
  /// a phantom's effective value is read back through `resolve_cell_value`.
  std::vector<CellAddress> spill_phantom_addresses() const;

  /// Returns every anchor whose last attempted spill footprint intersects the
  /// supplied rectangle.  The returned coordinates are copies, so callers
  /// may release the sheet lock before marking the anchors dirty in the
  /// workbook dependency graph.  This is a read-only query: unlike
  /// `commit_spill`, it never records a new blocked footprint.
  std::vector<CellAddress> blocked_spill_anchors_intersecting(std::uint32_t first_row, std::uint32_t first_col,
                                                              std::uint32_t rows, std::uint32_t cols) const;

  /// Returns every currently committed spill anchor whose materialised
  /// rectangle intersects the supplied rectangle. Coordinates are copies so
  /// callers may release the sheet lock before marking the anchors dirty.
  std::vector<CellAddress> committed_spill_anchors_intersecting(std::uint32_t first_row, std::uint32_t first_col,
                                                                std::uint32_t rows, std::uint32_t cols) const;

  /// Returns all currently blocked spill anchors.  Structural workbook edits
  /// use this after remapping the sheet-local table to mark surviving formula
  /// anchors dirty.  The returned coordinates are copies.
  std::vector<CellAddress> blocked_spill_anchors() const;

  /// Returns copies of the full blocked-footprint records.  Workbook
  /// structural edits snapshot these before formula text rewrites (which may
  /// temporarily clear an anchor entry), then restore the records at their
  /// mapped coordinates after the cell move.
  std::vector<BlockedSpillFootprint> blocked_spill_footprints() const;

  /// Returns a copy of the committed spill rectangle covering `(row, col)`.
  /// The anchor itself is included, unlike `spill_region_covering`, because
  /// mutation callers need the complete rectangle before they clear it.
  std::optional<BlockedSpillFootprint> committed_spill_footprint_covering(std::uint32_t row, std::uint32_t col) const;

  /// Returns copies of every currently committed spill rectangle. The sheet
  /// spill mutex is held exactly once for the snapshot and no references to
  /// the spill table escape. Blocked footprints are intentionally excluded.
  std::vector<SpillFootprint> committed_spill_footprints() const;

  /// Restores blocked-footprint records at their already-mapped coordinates.
  /// Invalid/degenerate records are ignored.  Intended for the workbook's
  /// lock-aware structural-edit path.
  void restore_blocked_spill_footprints(std::vector<BlockedSpillFootprint> footprints);

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

  /// Appends the effective value of every coordinate in the rectangle
  /// `[first_row, last_row] x [first_col, last_col]` to `out`, in row-major
  /// order, taking the sheet lock once for the whole rectangle.
  ///
  /// Reading a rectangle a cell at a time costs two lock acquisitions per
  /// cell — one for `cell_at`, one for `resolve_cell_value` — which dominates
  /// a wide aggregate such as `SUM(A1:A100000)`. Each appended value is what
  /// `resolve_cell_value` would return, except that a coordinate holding a
  /// formula is appended as its *cached* value and its position in `out` is
  /// collected in `formula_indices`. Evaluating a formula re-enters the sheet
  /// and would deadlock on the non-recursive lock, so the caller re-resolves
  /// those positions itself after this returns.
  ///
  /// A reversed or out-of-range rectangle appends nothing.
  void read_range(std::uint32_t first_row, std::uint32_t last_row, std::uint32_t first_col, std::uint32_t last_col,
                  std::vector<Value>& out, std::vector<std::size_t>& formula_indices) const;

  /// Registers a spill region anchored at `(anchor_row, anchor_col)` with
  /// the given dimensions and row-major cell payload.
  ///
  /// Behaviour:
  ///
  ///   * Any existing region anchored at `(anchor_row, anchor_col)` is
  ///     cleared first (including its reverse-map entries).
  ///   * The would-be footprint is checked through `spill_would_collide`.
  ///     Only the stored Cell at the anchor is ignored. Any other occupied
  ///     stored cell, intersection with another spill rectangle, or
  ///     intersection with a merged range (including a merge at the anchor)
  ///     sets the anchor's `cached_value` to `#SPILL!`, registers no region,
  ///     and returns `false`. Pre-existing literals and sheet metadata are
  ///     preserved.
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

  /// Invalidates every dynamic-array spill region. Structural row/column
  /// edits use this before moving cells; the next recalculation recreates
  /// regions at their new coordinates.
  void clear_all_spills() noexcept;

  /// Adds a merge while holding the same sheet lock used by spill collision
  /// checks.  Reader/setup code may continue to use `mutable_merges()` when
  /// no concurrent evaluation is possible; workbook/UI mutations should use
  /// this method.
  void add_merge(MergeRange merge);

  /// Removes every merge intersecting `range`, returning copies of the
  /// rectangles actually erased. The operation is lock-aware so a concurrent
  /// spill check cannot observe a partially erased merge vector.
  std::vector<MergeRange> remove_merges_intersecting(MergeRange range);

  /// Removes one merge by insertion order. Returns false when `index` is out
  /// of range. When `removed` is non-null, it receives a copy of the erased
  /// rectangle. The operation is lock-aware.
  bool remove_merge_at(std::size_t index, MergeRange* removed = nullptr);

  /// Removes every merge and returns the number removed. The operation is
  /// lock-aware.
  std::size_t clear_merges();

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
  // Sheet protection (passive round-trip)
  // ---------------------------------------------------------------------------

  /// Read-only access to the sheet's `<sheetProtection>` flags.
  /// Populated by the OOXML reader; an absent element leaves
  /// `enabled = false` and every other field at its default. The
  /// engine does not enforce locks at evaluation time — these flags
  /// are surfaced for the host UI's "Protect Sheet" dialog.
  const SheetProtection& protection() const noexcept { return protection_; }

  /// Mutable access to the sheet's protection flags. Exposed so the
  /// OOXML reader and the C ABI mutator can populate fields without an
  /// extra accessor pair per attribute.
  SheetProtection& mutable_protection() noexcept { return protection_; }

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

  /// Read-only access to the sheet's `<sheetFormatPr>` default metrics
  /// (default column width / row height). Populated by the OOXML
  /// reader; carries struct defaults when the element is absent.
  const SheetFormatDefaults& format_defaults() const noexcept { return format_defaults_; }

  /// Mutable access to the sheet's `<sheetFormatPr>` default metrics.
  /// Exposed so the OOXML reader can populate fields without an extra
  /// accessor pair per attribute.
  SheetFormatDefaults& mutable_format_defaults() noexcept { return format_defaults_; }

  /// Read-only access to raw print/page setup metadata captured by the
  /// OOXML reader.
  const SheetPrintSettings& print_settings() const noexcept { return print_settings_; }

  /// Mutable access for the OOXML reader.
  SheetPrintSettings& mutable_print_settings() noexcept { return print_settings_; }

  const WorksheetRawExtensions& raw_extensions() const noexcept { return raw_extensions_; }
  WorksheetRawExtensions& mutable_raw_extensions() noexcept { return raw_extensions_; }

  /// Package-relative path of the DrawingML part (`xl/drawings/drawingN.xml`)
  /// this sheet references via `<drawing r:id="...">`, or empty when the
  /// sheet anchors no drawing. Populated by the OOXML reader; the part
  /// body, its rels, and any anchored media round-trip through the
  /// workbook's passthrough parts. The writer re-emits both the
  /// `<drawing>` element and the sheet-rels relationship, minting a fresh
  /// rId, so the drawing stays reachable in the package graph.
  const std::string& drawing_rel_target() const noexcept { return drawing_rel_target_; }

  /// Sets the sheet's DrawingML part path. Plain metadata; no lifecycle
  /// or recalc interaction.
  void set_drawing_rel_target(std::string target) { drawing_rel_target_ = std::move(target); }

  /// Package-relative path of this sheet's legacy comment VML drawing, or
  /// empty for a newly-created comment collection. It is distinct from the
  /// DrawingML target above and lets comment shapes survive sheet reordering.
  const std::string& comment_vml_path() const noexcept { return comment_vml_path_; }
  void set_comment_vml_path(std::string path) { comment_vml_path_ = std::move(path); }

  /// Relationships from this worksheet's `.rels` part that Formulon does
  /// not otherwise model (OLE objects, controls, slicers, and so on).
  /// Their target parts travel through Workbook passthrough storage; keeping
  /// the original rIds lets raw worksheet XML which refers to them remain
  /// connected after a round trip.
  const std::vector<io::UnknownRelationship>& unknown_relationships() const noexcept { return unknown_relationships_; }
  void set_unknown_relationships(std::vector<io::UnknownRelationship> relationships) {
    unknown_relationships_ = std::move(relationships);
  }

  /// Raw `<autoFilter>` element captured from the worksheet, or empty when
  /// the sheet has no auto-filter. The engine does not model filter
  /// criteria yet; the element round-trips verbatim so the filter range
  /// and any column criteria survive a save cycle. The writer re-emits it
  /// in ECMA-376 order (after `<sheetProtection>`, before `<mergeCells>`).
  const std::string& auto_filter_xml() const noexcept { return auto_filter_xml_; }

  /// Sets the raw `<autoFilter>` element. Plain metadata.
  void set_auto_filter_xml(std::string xml) { auto_filter_xml_ = std::move(xml); }

  /// Raw worksheet-level `<extLst>` element captured from the worksheet,
  /// or empty. Excel stores the *data* for several 2010+ extensions here —
  /// most importantly the `x14:conditionalFormattings` block holding
  /// DataBar negative-fill / axis / gradient / direction, linked back to
  /// the legacy `cfRule` by an `x14:id` GUID. The engine does not model
  /// these extensions; the block round-trips verbatim so the extended
  /// formatting survives a save cycle. Re-emitted at the worksheet tail
  /// (after `<tableParts>`) per ECMA-376 element order.
  const std::string& ext_lst_xml() const noexcept { return ext_lst_xml_; }

  /// Sets the raw worksheet-level `<extLst>` element. Plain metadata.
  void set_ext_lst_xml(std::string xml) { ext_lst_xml_ = std::move(xml); }

  /// Extra namespace declarations (and `mc:Ignorable`) captured from the
  /// source `<worksheet>` root, serialised as ` name="value"` attribute
  /// pairs, excluding the default `xmlns` / `xmlns:r`. Re-emitted on the
  /// writer's `<worksheet>` root so namespaced attributes carried inside a
  /// raw worksheet capture (e.g. `x14ac:*` in a `<sheetPr>` fragment)
  /// resolve to a declared prefix — mirrors the workbook-root handling.
  const std::string& root_extra_ns_attrs() const noexcept { return root_extra_ns_attrs_; }
  void set_root_extra_ns_attrs(std::string attrs) { root_extra_ns_attrs_ = std::move(attrs); }

  /// Verbatim `.xlsb` worksheet-tail records; empty for a sheet that did not
  /// come from a binary package. See `XlsbSheetTail` for what they carry and
  /// what verbatim retention does not do.
  const XlsbSheetTail& xlsb_tail() const noexcept { return xlsb_tail_; }
  void set_xlsb_tail(XlsbSheetTail tail) { xlsb_tail_ = std::move(tail); }

  // ---------------------------------------------------------------------------
  // Row / column structural edits
  // ---------------------------------------------------------------------------
  //
  // Sheet-local migration of cell storage and integer-coordinate
  // metadata (merges, hyperlinks, comments, data-validation ranges) for
  // a row or column insert / delete. The workbook orchestrates the wider
  // operation: it walks every sheet and rewrites cross-sheet formula
  // text via the AST shifter, then calls these methods on the affected
  // sheet to migrate the local data.
  //
  // `row` / `col` is the 0-based index where the edit begins. `count`
  // must be `>= 1`. Cells that would land past `kMaxRows` / `kMaxCols`
  // on an insert are dropped wholesale; cells inside the deleted
  // interval on a delete are dropped wholesale. Metadata anchored on
  // dropped cells is also removed; metadata anchored on shifted cells
  // is rewritten in place.

  /// Inserts `count` rows at `row`. Cells at `row` and beyond shift down.
  void insert_rows(std::uint32_t row, std::uint32_t count);

  /// Deletes `count` rows starting at `row`. Cells in `[row, row+count)`
  /// are dropped; cells past the deletion shift up.
  void delete_rows(std::uint32_t row, std::uint32_t count);

  /// Inserts `count` columns at `col`. Cells at `col` and beyond
  /// (within every populated row) shift right.
  void insert_cols(std::uint32_t col, std::uint32_t count);

  /// Deletes `count` columns starting at `col`. Cells in
  /// `[col, col+count)` are dropped; cells past the deletion shift left.
  void delete_cols(std::uint32_t col, std::uint32_t count);

 private:
  /// One row/column structural edit, as the coordinate mapping every
  /// sheet-attached structure derives from.
  ///
  /// `index` is the 0-based band the edit starts at and `count` its width. On
  /// an insert, a coordinate `>= index` moves forward by `count`. On a delete,
  /// `[index, index + count)` is the deleted band — a coordinate inside it is
  /// dropped or clamped, and a coordinate past it moves back by `count`.
  ///
  /// The four public edit methods each build one of these and hand it to
  /// `shift_sheet_metadata`, which is the single place that enumerates the
  /// structures. A structure that is not listed there does not move, and that
  /// failure is silent in a user's file — `tests/integration/
  /// structural_edit_matrix_test.cpp` asserts each one's post-edit coordinate.
  struct StructuralEdit {
    std::uint32_t index = 0;
    std::uint32_t count = 0;
    bool is_delete = false;
    bool row_axis = false;
  };

  /// Applies `edit` to every sheet-attached structure other than the cells
  /// themselves (which each public edit method moves first, since the cell
  /// storage differs per axis).
  void shift_sheet_metadata(const StructuralEdit& edit);

  // ---------------------------------------------------------------------------
  // Non-locking spill/cell helpers (caller must already hold `spill_mutex_`).
  // ---------------------------------------------------------------------------
  //
  // The public `cell_at` / `clear_spill` / `spill_region_covering` methods each
  // acquire `spill_mutex_` and then delegate to the matching `_locked` body.
  // Because `spill_mutex_` is a non-recursive `std::mutex`, any caller that
  // already holds it (e.g. `commit_spill` scanning its footprint) must reach
  // the `_locked` variant directly to avoid self-deadlock.

  /// `cell_at` body; assumes `spill_mutex_` is held.
  const Cell* cell_at_locked(std::uint32_t row, std::uint32_t col) const noexcept;

  /// `clear_spill` body; assumes `spill_mutex_` is held.
  void clear_spill_locked(std::uint32_t anchor_row, std::uint32_t anchor_col) noexcept;

  /// `spill_region_covering` body; assumes `spill_mutex_` is held.
  const SpillRegion* spill_region_covering_locked(std::uint32_t row, std::uint32_t col) const noexcept;

  /// `spill_would_collide` body; assumes `spill_mutex_` is held.
  bool spill_would_collide_locked(std::uint32_t anchor_row, std::uint32_t anchor_col, std::uint32_t rows,
                                  std::uint32_t cols) const noexcept;

  /// `probe_spill_footprint` body; assumes `spill_mutex_` is held.
  SpillAdmission probe_spill_footprint_locked(std::uint32_t anchor_row, std::uint32_t anchor_col, std::uint32_t rows,
                                              std::uint32_t cols, std::uint64_t* scan_steps) const noexcept;

  /// `reject_spill_footprint` body; assumes `spill_mutex_` is held.
  void reject_spill_footprint_locked(std::uint32_t anchor_row, std::uint32_t anchor_col, std::uint32_t rows,
                                     std::uint32_t cols);

  /// True when a stored cell other than the anchor is occupied anywhere in
  /// the half-open rectangle `[anchor_row, row_end) x [anchor_col, col_end)`.
  /// Walks the stored rows or probes the rectangle's rows, whichever set is
  /// smaller, so the cost follows the sheet's contents rather than the
  /// rectangle's area. Adds one unit of work per row visited or probed and
  /// per cell inspected to `*scan_steps` when it is non-null. Caller must
  /// hold `spill_mutex_`.
  bool footprint_holds_occupied_cell_locked(std::uint32_t anchor_row, std::uint32_t anchor_col, std::uint64_t row_end,
                                            std::uint64_t col_end, std::uint64_t* scan_steps) const noexcept;

  /// Clears committed spill regions but preserves blocked-footprint records.
  /// Structural row/column edits use this before remapping pending anchors;
  /// public `clear_all_spills()` also drops the blocked records.
  void clear_committed_spills_locked() noexcept;

  /// Remaps blocked-footprint anchors through one structural row/column edit.
  /// Caller must hold `spill_mutex_`.
  void shift_blocked_spills_locked(const StructuralEdit& edit);

  std::string name_;
  // Non-empty only for chartsheets / dialog sheets / macro sheets and other
  // non-worksheet OOXML sheet types. Their XML and descendant parts are raw
  // passthrough payloads; these two fields preserve the workbook-level link.
  std::string opaque_ooxml_part_path_;
  std::string opaque_ooxml_relationship_type_;
  std::unordered_map<std::uint32_t, RowCells> rows_;
  std::uint64_t cell_enumeration_revision_ = 0;
  // Lazily allocated: most sheets do not host any spill regions, so the
  // table is only materialised on the first `commit_spill` call.
  std::unique_ptr<SpillTable> spill_table_;
  // Guards every access to `spill_table_` and `rows_` so that parallel
  // recalc workers committing dynamic-array spills on the same sheet (each
  // running outside the scheduler's write lock) cannot tear the underlying
  // maps. This is the innermost link of the one workbook-wide lock order,
  // stated in full in `eval/recalc_engine.h`: workbook compound mutation ->
  // `RecalcEngine::mutex_` -> the scheduler's per-pass `write_mutex` ->
  // `spill_mutex_`. It is never acquired ahead of any of them. The
  // spill-commit path runs without `write_mutex`; the only path that
  // effectively nests the two is `set_cell_cached_value`, which the
  // scheduler calls under its `write_mutex` and which then takes
  // `spill_mutex_` internally. `mutable` because const observers
  // (`resolve_cell_value`, `cell_at`, `spill_region_*`) must lock it too.
  // Heap ownership makes the lock movable with its sheet, so Sheet's
  // out-of-line defaulted move members cannot omit future metadata fields.
  mutable std::unique_ptr<std::mutex> spill_mutex_;
  // Pivot tables anchored on this sheet. Empty by default; populated by
  // the OOXML reader at workbook-load time. Heap-owned so addresses stay
  // stable across vector reallocations.
  std::vector<std::unique_ptr<pivot::PivotTable>> pivot_tables_;
  // Conditional-format blocks attached to this sheet. Empty by default;
  // populated by the OOXML reader from `<conditionalFormatting>` blocks.
  std::vector<cf::ConditionalFormat> conditional_formats_;
  // Merge ranges from `<mergeCells>`. Empty by default; populated by the
  // OOXML reader and round-tripped through the writer.
  std::vector<MergeRange> merges_;
  // Hyperlinks from `<hyperlinks>` plus the sheet's rels file. Empty by default.
  std::vector<Hyperlink> hyperlinks_;
  // Per-cell text comments aggregated from `xl/comments<N>.xml`. Empty by default.
  std::vector<CellComment> comments_;
  // `<dataValidations>` blocks. Empty by default.
  std::vector<DataValidation> validations_;
  // `<sheetProtection>` flags. `protection_.enabled = false` by
  // default; the writer emits no element when the field stays at its
  // default-constructed state.
  SheetProtection protection_;
  // Per-sheet view state (zoom, frozen panes, tab visibility). Defaults
  // are populated inline; the OOXML reader overwrites them when present.
  SheetView view_;
  // Per-sheet layout overrides (column spans + row overrides). Empty by
  // default; populated by the OOXML reader from `<cols>` / `<row>` entries.
  SheetLayout layout_;
  // Default column width / row height from `<sheetFormatPr>`. Carries
  // struct defaults when the element is absent.
  SheetFormatDefaults format_defaults_;
  // Raw print/page setup metadata and its optional printerSettings rel.
  SheetPrintSettings print_settings_;
  // Raw worksheet children not represented by the editing/evaluation model.
  WorksheetRawExtensions raw_extensions_;
  // Package-relative path of the DrawingML part referenced by
  // `<drawing r:id>`, or empty. The part itself round-trips via the
  // workbook's passthrough parts.
  std::string drawing_rel_target_;
  // Package-relative VML drawing paired with this sheet's comments. The
  // writer may assign a new output filename, but uses this path to find the
  // original bytes in passthrough storage.
  std::string comment_vml_path_;
  // Unmodelled entries from `xl/worksheets/_rels/sheetN.xml.rels`.
  std::vector<io::UnknownRelationship> unknown_relationships_;
  // Raw `<autoFilter>` element, or empty. Round-trips verbatim; filter
  // criteria are not modelled.
  std::string auto_filter_xml_;
  // Raw worksheet-level `<extLst>` element (x14 conditional-formatting
  // data etc.), or empty. Round-trips verbatim.
  std::string ext_lst_xml_;
  // Extra `<worksheet>` root namespace declarations captured for verbatim
  // re-emission so namespaced attributes in raw captures resolve. Empty by
  // default.
  std::string root_extra_ns_attrs_;
  // Verbatim `.xlsb` worksheet-tail records, or empty for a sheet that did
  // not come from a binary package.
  XlsbSheetTail xlsb_tail_;
};

}  // namespace formulon

#endif  // FORMULON_SHEET_H_
