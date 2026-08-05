
#include "print/pagination.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <limits>
#include <optional>
#include <string>
#include <vector>

#include "cell.h"
#include "print/page_setup.h"
#include "print/print_area.h"
#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace print {
namespace {

// --- Excel column-width geometry constants. ---
//
// Excel stores a column's width in "character" units relative to the
// default font's maximum digit width (MDW). The character-to-pixel
// conversion below is Excel's documented formula; the constants are
// named so the arithmetic carries no bare literals.

/// Maximum digit width of the default font (Calibri 11), in pixels.
constexpr double kMaxDigitWidthPx = 7.0;

/// Fixed-point scale Excel uses inside the width formula (1 character
/// expands to 256 internal units).
constexpr double kWidthFixedPointScale = 256.0;

/// Padding term in the width formula (half the fixed-point scale).
constexpr double kWidthPaddingHalf = 128.0;

/// Screen pixel density Excel's width model assumes, in pixels per inch.
constexpr double kScreenPixelsPerInch = 96.0;

/// Excel's standard default column width, in character units. Used when
/// neither a `<col>` override nor `<sheetFormatPr defaultColWidth>`
/// applies.
constexpr double kStandardColWidthChars = 8.43;

/// Excel's standard default row height, in points. Used when neither a
/// `<row ht>` override nor `<sheetFormatPr defaultRowHeight>` applies.
constexpr double kStandardRowHeightPt = 15.0;

/// Lower clamp for the effective scale factor (1%). Mirrors Excel's
/// `<pageSetup scale>` minimum so a degenerate `scale="0"` cannot divide
/// the printable area by zero.
constexpr double kMinScaleFactor = 0.01;

/// Converts an Excel column width in character units to a width in
/// points, following Excel's character -> pixel -> point pipeline.
double ColumnCharsToPoints(double chars) {
  // pixels = floor(((256 * chars + floor(128 / MDW)) / 256) * MDW)
  const double padded = std::floor(kWidthPaddingHalf / kMaxDigitWidthPx);
  const double pixels =
      std::floor(((kWidthFixedPointScale * chars + padded) / kWidthFixedPointScale) * kMaxDigitWidthPx);
  return pixels * kPointsPerInch / kScreenPixelsPerInch;
}

/// Returns the width, in character units, of column `col` on `sheet`.
///
/// Precedence: an explicit `<col>` span covering `col`, then
/// `<sheetFormatPr defaultColWidth>`, then Excel's 8.43-character
/// standard default.
double ColumnWidthChars(const Sheet& sheet, std::uint32_t col) {
  for (const ColumnLayout& span : sheet.layout().columns) {
    if (col >= span.first && col <= span.last) {
      // Hidden columns occupy no printed width, so they never advance the
      // page grid (Excel excludes them from pagination extent).
      return span.hidden ? 0.0 : span.width;
    }
  }
  const SheetFormatDefaults& defaults = sheet.format_defaults();
  if (defaults.has_default_col_width) {
    return defaults.default_col_width;
  }
  return kStandardColWidthChars;
}

/// Returns the height, in points, of row `row` on `sheet`.
///
/// Precedence: an explicit `<row ht>` override, then
/// `<sheetFormatPr defaultRowHeight>`, then Excel's 15-point standard
/// default.
double RowHeightPoints(const Sheet& sheet, std::uint32_t row) {
  for (const RowLayout& override_row : sheet.layout().row_overrides) {
    if (override_row.row == row) {
      // Hidden rows occupy no printed height, so they never advance the
      // page grid (Excel excludes them from pagination extent).
      return override_row.hidden ? 0.0 : override_row.height;
    }
  }
  const SheetFormatDefaults& defaults = sheet.format_defaults();
  if (defaults.has_default_row_height) {
    return defaults.default_row_height;
  }
  return kStandardRowHeightPt;
}

/// Computes the sheet's used range as a single rectangle, walking the
/// populated cells. Returns false when the sheet has no non-blank cell.
bool ComputeUsedRange(const Sheet& sheet, CellRange* out_range) {
  bool any = false;
  std::uint32_t min_row = 0;
  std::uint32_t min_col = 0;
  std::uint32_t max_row = 0;
  std::uint32_t max_col = 0;
  for (const auto& [row_index, cells] : sheet.rows()) {
    for (std::size_t c = 0; c < cells.size(); ++c) {
      const Cell& cell = cells[c];
      if (cell.formula_text.empty() && cell.cached_value.is_blank()) {
        continue;
      }
      const auto col_index = static_cast<std::uint32_t>(c);
      if (!any) {
        min_row = max_row = row_index;
        min_col = max_col = col_index;
        any = true;
        continue;
      }
      min_row = std::min(min_row, row_index);
      max_row = std::max(max_row, row_index);
      min_col = std::min(min_col, col_index);
      max_col = std::max(max_col, col_index);
    }
  }
  // Dynamic-array spill phantoms occupy their coordinates in Excel's used
  // range even though they hold no stored `Cell`; fold them into the bounding
  // box so a spilled region paginates against its full extent, not just the
  // anchor.
  for (const CellAddress& addr : sheet.spill_phantom_addresses()) {
    if (!any) {
      min_row = max_row = addr.row;
      min_col = max_col = addr.col;
      any = true;
      continue;
    }
    min_row = std::min(min_row, addr.row);
    max_row = std::max(max_row, addr.row);
    min_col = std::min(min_col, addr.col);
    max_col = std::max(max_col, addr.col);
  }
  if (!any) {
    return false;
  }
  *out_range = CellRange{min_row, min_col, max_row, max_col};
  return true;
}

/// Returns the bounding box that encloses every rectangle in `ranges`.
/// `ranges` must be non-empty.
CellRange BoundingBox(const std::vector<CellRange>& ranges) {
  CellRange box = ranges.front();
  for (const CellRange& r : ranges) {
    box.first_row = std::min(box.first_row, r.first_row);
    box.first_col = std::min(box.first_col, r.first_col);
    box.last_row = std::max(box.last_row, r.last_row);
    box.last_col = std::max(box.last_col, r.last_col);
  }
  return box;
}

/// Returns the rectangular intersection of `a` and `b`, or `std::nullopt`
/// when they are disjoint.
std::optional<CellRange> Intersect(const CellRange& a, const CellRange& b) {
  CellRange out;
  out.first_row = std::max(a.first_row, b.first_row);
  out.first_col = std::max(a.first_col, b.first_col);
  out.last_row = std::min(a.last_row, b.last_row);
  out.last_col = std::min(a.last_col, b.last_col);
  if (out.first_row > out.last_row || out.first_col > out.last_col) {
    return std::nullopt;
  }
  return out;
}

/// True when `index` is the target of a *manual* break in `breaks`.
///
/// Excel persists automatic page breaks (`man="0"`) alongside user-placed
/// ones once a sheet has been previewed or printed. Only breaks with
/// `manual == true` force a page boundary; automatic breaks are
/// recomputed by the pagination walk and must not be treated as forced.
bool HasManualBreakAt(const std::vector<ManualBreak>& breaks, std::uint32_t index) {
  return std::any_of(breaks.begin(), breaks.end(),
                     [index](const ManualBreak& brk) { return brk.manual && brk.id == index; });
}

/// One axis of the break walk.
///
/// `track_sizes[i]` is the size in points of the i-th track of the print
/// area (column or row). `manual` lists track indices (absolute, not
/// print-area-relative) that force a break before themselves.
struct AxisInput {
  std::uint32_t first = 0;                           ///< Absolute index of the first print-area track.
  std::vector<double> track_sizes;                   ///< Per-track size in points (model-scaled).
  const std::vector<ManualBreak>* manual = nullptr;  ///< Manual breaks for this axis.
  double limit_pt = 0.0;                             ///< Printable body extent for this axis.
};

/// Walks one axis, accumulating track sizes until the printable limit is
/// reached. Appends the absolute index each break precedes to `out_breaks`
/// and returns the number of pages produced (always >= 1 when the axis has
/// at least one track).
std::uint32_t WalkAxis(const AxisInput& axis, std::vector<std::uint32_t>* out_breaks) {
  const std::size_t track_count = axis.track_sizes.size();
  if (track_count == 0) {
    return 0;
  }

  std::uint32_t pages = 1;
  double accumulated = 0.0;
  for (std::size_t i = 0; i < track_count; ++i) {
    const auto absolute = static_cast<std::uint32_t>(axis.first + i);
    const double size = axis.track_sizes[i];
    const bool manual_break = i != 0 && HasManualBreakAt(*axis.manual, absolute);
    // An automatic break fires when adding this track overflows the body
    // and the page already holds at least one track (so a single track
    // wider than the page still occupies one page rather than zero).
    const bool overflow = i != 0 && accumulated > 0.0 && accumulated + size > axis.limit_pt;
    if (manual_break || overflow) {
      out_breaks->push_back(absolute);
      ++pages;
      accumulated = 0.0;
    }
    accumulated += size;
  }
  return pages;
}

/// Computes the uniform scale factor applied to cell sizes.
///
/// When `fit_to_page` is set, derives a factor that shrinks the print
/// area to fit within `fit_to_width` pages horizontally and
/// `fit_to_height` pages vertically (a `fit_to_*` of 0 leaves that axis
/// unconstrained). Otherwise the `<pageSetup scale>` percentage is used.
/// The factor multiplies every cell's point size: a factor below 1.0
/// shrinks cells so more fit per page.
double ComputeScaleFactor(const PageSetup& setup, const PrintableArea& area, double total_width_pt,
                          double total_height_pt) {
  if (!setup.fit_to_page) {
    constexpr double kPercentDivisor = 100.0;
    return std::max(kMinScaleFactor, static_cast<double>(setup.scale) / kPercentDivisor);
  }

  double factor = 1.0;
  if (setup.fit_to_width > 0 && total_width_pt > 0.0 && area.width_pt > 0.0) {
    const double allowed = area.width_pt * static_cast<double>(setup.fit_to_width);
    factor = std::min(factor, allowed / total_width_pt);
  }
  if (setup.fit_to_height > 0 && total_height_pt > 0.0 && area.height_pt > 0.0) {
    const double allowed = area.height_pt * static_cast<double>(setup.fit_to_height);
    factor = std::min(factor, allowed / total_height_pt);
  }
  // A fit factor only ever shrinks; Excel never enlarges to fill pages.
  return std::max(kMinScaleFactor, std::min(1.0, factor));
}

}  // namespace

Expected<PaginationResult, Error> paginate(const Workbook& wb, std::uint32_t sheet_index) {
  if (sheet_index >= wb.sheet_count()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "Sheet index out of range for pagination",
                      "sheet_index=" + std::to_string(sheet_index));
  }
  const Sheet& sheet = wb.sheet(sheet_index);

  // 1. Resolve the print area. The reported `result.print_area` mirrors
  // Excel's `PageSetup.PrintArea` exactly: empty when the workbook
  // defines no print area, even if the sheet has populated cells. The
  // used-range fallback is a *pagination* concern only, not a reporting
  // one, so it lives in a separate `effective_area` used solely to size
  // and walk the page grid below.
  auto area_or = resolve_print_area(wb, sheet_index);
  if (!area_or) {
    return area_or.error();
  }
  PaginationResult result;
  result.print_area = area_or.value();

  // Excel's `HPageBreaks` / `VPageBreaks` are populated against the
  // *populated* region of the print area, not its full geometric span:
  // a print area whose four corners are blank effectively paginates
  // against the (possibly much smaller) populated bounding box. Compute
  // the populated box once so each print-area rectangle can be
  // intersected with it below.
  CellRange used_box;
  const bool has_used_range = ComputeUsedRange(sheet, &used_box);

  std::vector<CellRange> effective_areas;
  if (result.print_area.empty()) {
    // No explicit print area: fall back to the used range. An empty
    // sheet with no print area produces no pages.
    if (!has_used_range) {
      return result;
    }
    effective_areas.push_back(used_box);
  } else if (!has_used_range) {
    // Explicit print area on an otherwise-empty sheet: Excel still
    // paginates a *bounded* declared geometry (the area defines the page
    // grid even when no cell carries content). A whole-column (`$A:$A`) or
    // whole-row (`$1:$1`) print area, however, has no content to paginate
    // and must not walk the full 1,048,576-row / 16,384-column grid (which
    // would also reserve a multi-megabyte track vector below). Drop the
    // unbounded rectangles; if none remain the sheet produces no pages.
    constexpr std::uint32_t kMaxRowIndex = Sheet::kMaxRows - 1U;
    constexpr std::uint32_t kMaxColIndex = Sheet::kMaxCols - 1U;
    for (const CellRange& r : result.print_area) {
      if (r.last_row < kMaxRowIndex && r.last_col < kMaxColIndex) {
        effective_areas.push_back(r);
      }
    }
    if (effective_areas.empty()) {
      return result;
    }
  } else {
    // Intersect each declared rectangle with the populated bounding box.
    // A rectangle whose populated intersection is empty contributes no
    // pages (Excel reports zero breaks across such an area).
    for (const CellRange& r : result.print_area) {
      if (auto isect = Intersect(r, used_box); isect.has_value()) {
        effective_areas.push_back(*isect);
      }
    }
    if (effective_areas.empty()) {
      return result;
    }
  }

  // 2. Printable body area is determined once from the page setup.
  const SheetPrintSettings& settings = sheet.print_settings();
  PrintableArea body = compute_printable_area(settings.page_setup, settings.page_margins);

  // Print titles (repeat-rows / repeat-columns) are reprinted on every
  // page, so they steal body extent that is otherwise available for
  // data rows / columns. Subtract their summed sizes once: pagination
  // walks the (full) print area, and the smaller body limit naturally
  // forces a break sooner.
  auto titles_or = resolve_print_titles(wb, sheet_index);
  if (!titles_or) {
    return titles_or.error();
  }
  const PrintTitles& titles = titles_or.value();
  if (titles.repeat_rows.has_value()) {
    const auto [first, last] = *titles.repeat_rows;
    double title_height = 0.0;
    for (std::uint32_t row = first; row <= last; ++row) {
      title_height += RowHeightPoints(sheet, row);
    }
    // Empirical: when print-title rows are enabled, Excel reserves a
    // minimum body band of ~5 default rows even if the actual title
    // rows sum to less. Round-3 capture: print_titles_repeat_{1,3,5}_rows
    // on A1:D40 all break at h=[39], so the subtraction cannot be the
    // raw title_height (which would yield 15 / 45 / 75 pt). Using a
    // 5-default-row floor reproduces the observed h=[39] for all three.
    constexpr double kMinTitleReserveRows = 5.0;
    const SheetFormatDefaults& defaults = sheet.format_defaults();
    const double default_row_h = defaults.has_default_row_height ? defaults.default_row_height : kStandardRowHeightPt;
    title_height = std::max(title_height, kMinTitleReserveRows * default_row_h);
    body.height_pt = std::max(0.0, body.height_pt - title_height);
  }
  if (titles.repeat_cols.has_value()) {
    const auto [first, last] = *titles.repeat_cols;
    double title_width = 0.0;
    for (std::uint32_t col = first; col <= last; ++col) {
      title_width += ColumnCharsToPoints(ColumnWidthChars(sheet, col));
    }
    body.width_pt = std::max(0.0, body.width_pt - title_width);
  }

  // 3. The model scale factor is shared across every effective area:
  // Excel applies a single sheet-wide `<pageSetup scale>` / `fitToPage`
  // factor, so we size the factor against the union bounding box.
  const CellRange union_box = BoundingBox(effective_areas);
  double union_total_width = 0.0;
  for (std::uint32_t col = union_box.first_col; col <= union_box.last_col; ++col) {
    union_total_width += ColumnCharsToPoints(ColumnWidthChars(sheet, col));
  }
  double union_total_height = 0.0;
  for (std::uint32_t row = union_box.first_row; row <= union_box.last_row; ++row) {
    union_total_height += RowHeightPoints(sheet, row);
  }
  const double scale = ComputeScaleFactor(settings.page_setup, body, union_total_width, union_total_height);

  // 4. Paginate each rectangle independently and aggregate the breaks.
  // Excel's `HPageBreaks` / `VPageBreaks` collections report the union
  // of breaks across all areas; the page count sums across areas.
  std::vector<std::uint32_t> all_h;
  std::vector<std::uint32_t> all_v;
  std::uint32_t total_pages = 0;
  for (const CellRange& rect : effective_areas) {
    // Row axis: automatic overflow breaks plus manual row breaks within the
    // area's rows, counted per area.
    std::vector<double> row_points;
    row_points.reserve(rect.last_row - rect.first_row + 1);
    for (std::uint32_t row = rect.first_row; row <= rect.last_row; ++row) {
      row_points.push_back(RowHeightPoints(sheet, row) * scale);
    }
    AxisInput row_axis;
    row_axis.first = rect.first_row;
    row_axis.track_sizes = std::move(row_points);
    row_axis.manual = &settings.manual_row_breaks;
    row_axis.limit_pt = body.height_pt;
    const std::uint32_t row_pages = WalkAxis(row_axis, &all_h);

    // Column axis: symmetric per-area walk through the same axis walker.
    // Excel never auto-breaks columns at explicit print scale (a wide
    // print area renders on a single page-column and clips at the right
    // margin; both VPageBreaks and Pages.Count ignore automatic column
    // overflow — see tests/oracle/cases_wb/print_matrix.yaml Block C and
    // print_pagination.yaml wide_table_vertical_breaks), so the width limit
    // is unbounded and only manual column breaks contribute. Using the same
    // per-area walker as the row axis makes multi-area column pagination
    // symmetric with multi-area row pagination: each area's manual column
    // breaks are counted within that area rather than de-duplicated across
    // areas.
    std::vector<double> col_points;
    col_points.reserve(rect.last_col - rect.first_col + 1);
    for (std::uint32_t col = rect.first_col; col <= rect.last_col; ++col) {
      col_points.push_back(ColumnCharsToPoints(ColumnWidthChars(sheet, col)) * scale);
    }
    AxisInput col_axis;
    col_axis.first = rect.first_col;
    col_axis.track_sizes = std::move(col_points);
    col_axis.manual = &settings.manual_col_breaks;
    col_axis.limit_pt = std::numeric_limits<double>::infinity();
    const std::uint32_t col_pages = WalkAxis(col_axis, &all_v);

    total_pages += col_pages * row_pages;
  }

  // 5. Sort and de-duplicate aggregated break positions: Excel's COM
  // collections are ascending and never repeat.
  auto sort_unique = [](std::vector<std::uint32_t>* v) {
    std::sort(v->begin(), v->end());
    v->erase(std::unique(v->begin(), v->end()), v->end());
  };
  sort_unique(&all_h);
  sort_unique(&all_v);

  result.h_breaks = std::move(all_h);
  result.v_breaks = std::move(all_v);
  result.page_count = total_pages;
  return result;
}

}  // namespace print
}  // namespace formulon
