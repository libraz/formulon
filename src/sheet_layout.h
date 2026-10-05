//
// Worksheet geometry and print setup: column spans, row overrides, the
// `<sheetFormatPr>` defaults, and the page setup, margins and manual breaks
// the pagination engine reads. Plain value types with no dependency on the
// cell store.

#ifndef FORMULON_SHEET_LAYOUT_H_
#define FORMULON_SHEET_LAYOUT_H_

#include <cstdint>
#include <string>
#include <vector>

namespace formulon {

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
  // True when the height is an explicit override (`customHeight="1"` /
  // BrtRowHdr fUnsynced); false for an auto height that only caches `ht`.
  bool custom_height = false;
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

/// Order in which a sheet's pages are numbered and printed
/// (`<pageSetup pageOrder>`).
enum class PageOrder : std::int32_t {
  kDownThenOver = 0,  ///< Each column of pages top to bottom, then the next column (default).
  kOverThenDown = 1,  ///< Each row of pages left to right, then the next row.
};

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
  std::uint32_t fit_to_width = 1;                   ///< `fitToWidth`, in pages.
  std::uint32_t fit_to_height = 1;                  ///< `fitToHeight`, in pages.
  bool fit_to_page = false;                         ///< True when `<sheetPr><pageSetUpPr fitToPage="1"/>`.
  PageOrder page_order = PageOrder::kDownThenOver;  ///< `pageOrder` attribute.
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

}  // namespace formulon

#endif  // FORMULON_SHEET_LAYOUT_H_
