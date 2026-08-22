
#include "print/pagination.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <string>
#include <vector>

#include "cell.h"
#include "print/page_setup.h"
#include "print/print_area.h"
#include "sheet.h"
#include "utils/resource_budget.h"
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

/// Points per character unit and the per-column padding, as Excel 365
/// resolves a character-unit column width under the Normal style pinned to
/// Calibri 11 (`Range.Width`, captured in every workbook golden's
/// `applied_geometry.column_widths_pt`).
///
/// These replace the textbook 96-DPI screen model (MDW 7px, so 5.25 pt per
/// character). That model predicts 157.5 pt for a 30-character column;
/// Excel resolves 171.0. Every measured size is an exact multiple of
/// 1/7 pt, and the law is linear across the captured widths:
///
///     20 chars -> 115.2857 pt    28 -> 159.8571 pt    30 -> 171.0 pt
///
/// Pagination compares these against the printable body, so using the
/// screen model made a wide print area fit roughly one column too many per
/// page.
///
/// The calibration is to one Normal font, and MDW is a property of that
/// font, so a workbook whose Normal style names a different one resolves a
/// different number of points per character. The scope of that was measured
/// by opening workbooks that differ only in font 0 and reading
/// `Range.Width` back:
///
///   * At 11 pt the family does not move it. `Calibri`, `游ゴシック` and
///     `ＭＳ Ｐゴシック` all resolve a 30-unit column to the same width, so
///     a ja-JP host declaring a Japanese body font paginates identically.
///   * The point size does move it, roughly in proportion, and the family
///     starts to matter away from 11 pt: `Calibri 18` resolves half again
///     as wide as `Calibri 11`, and `游ゴシック 14` a seventh wider than
///     `Calibri 14`.
///
/// Those observations are Mac Excel's, whose column geometry is a different
/// regime from the Windows primary oracle these constants come from, so
/// they establish that the dependency exists without supplying the numbers
/// to model it. Sizing the constants off the Normal font needs a Windows
/// capture over the same sweep; until then a workbook whose Normal font is
/// not 11 pt paginates against the 11 pt geometry.
constexpr double kPointsPerColumnChar = 39.0 / 7.0;
constexpr double kColumnPaddingPt = 27.0 / 7.0;

/// Excel's standard default column width, in character units. Used when
/// neither a `<col>` override nor `<sheetFormatPr defaultColWidth>`
/// applies.
constexpr double kStandardColWidthChars = 8.43;

/// Excel's default row height, in points, measured the same way as the
/// column constants above (`applied_geometry.row_heights_pt`). Used when
/// neither a `<row ht>` override nor `<sheetFormatPr defaultRowHeight>`
/// applies. The nominal 15.0 is the 96-DPI screen figure; Excel resolves
/// 102/7.
constexpr double kStandardRowHeightPt = 102.0 / 7.0;

/// Lower clamp for the effective scale factor (1%). Mirrors Excel's
/// `<pageSetup scale>` minimum so a degenerate `scale="0"` cannot divide
/// the printable area by zero.
constexpr double kMinScaleFactor = 0.01;

/// Converts an Excel column width in character units to a width in points.
double ColumnCharsToPoints(double chars) {
  return chars * kPointsPerColumnChar + kColumnPaddingPt;
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
      if (span.hidden) {
        return 0.0;
      }
      return HasExplicitColumnWidth(span)
                 ? span.width
                 : (sheet.format_defaults().has_default_col_width ? sheet.format_defaults().default_col_width
                                                                  : kStandardColWidthChars);
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
      if (override_row.hidden) {
        return 0.0;
      }
      return (override_row.has_height || override_row.height != 0.0)
                 ? override_row.height
                 : (sheet.format_defaults().has_default_row_height ? sheet.format_defaults().default_row_height
                                                                   : kStandardRowHeightPt);
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
  const auto extent = sheet.populated_extent(0U, 0U, Sheet::kMaxRows - 1U, Sheet::kMaxCols - 1U);
  if (!extent.has_value()) {
    return false;
  }
  *out_range = CellRange{extent->first_row, extent->first_col, extent->last_row, extent->last_col};
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

/// Relative tolerance for the "does this track still fit" comparison.
/// Sized to absorb the accumulated rounding of a few thousand additions
/// while staying far below one track's width.
constexpr double kAxisFitEpsilon = 1e-9;

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
std::uint32_t WalkAxis(const AxisInput& axis, std::vector<std::uint32_t>* out_breaks,
                       std::vector<std::uint32_t>* out_manual_breaks) {
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
    // Relative epsilon: `fit_to_page` sizes the scale so the content total
    // lands exactly on the limit, and a strict comparison then breaks a page
    // on nothing but accumulated rounding -- fitToWidth=1 could report two
    // page-columns. The tolerance is proportional so it stays meaningful at
    // any page size.
    const double slack = axis.limit_pt * kAxisFitEpsilon;
    const bool overflow = i != 0 && accumulated > 0.0 && accumulated + size > axis.limit_pt + slack;
    if (manual_break || overflow) {
      // A manual break belongs to the sheet, an automatic one to the page
      // region that overflowed. Excel reports them accordingly: a manual
      // break shared by two print areas appears once, while each area
      // contributes its own automatic break (`print_pagination.
      // multi_area_row_stacked_col_break` -> v=[3,7,7], where 3 is manual).
      (manual_break ? out_manual_breaks : out_breaks)->push_back(absolute);
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
///
/// `area` is the full printable body, before any print-title
/// reservation: repeated titles scale with the data, so they are
/// accounted for in the demand rather than deducted from the supply.
/// `title_*_pt` are their unscaled model sizes, and each page of a
/// multi-page fit reprints them -- so `N` pages carry `N` copies, which
/// is what puts the page count on the title term. With that denominator
/// the resulting factor satisfies `content * factor <= N * (area -
/// title * factor)` exactly, which is the same comparison the axis walk
/// then performs.
double ComputeScaleFactor(const PageSetup& setup, const PrintableArea& area, double total_width_pt,
                          double total_height_pt, double title_width_pt, double title_height_pt) {
  if (!setup.fit_to_page) {
    constexpr double kPercentDivisor = 100.0;
    return std::max(kMinScaleFactor, static_cast<double>(setup.scale) / kPercentDivisor);
  }

  double factor = 1.0;
  if (setup.fit_to_width > 0 && total_width_pt > 0.0 && area.width_pt > 0.0) {
    const double pages = static_cast<double>(setup.fit_to_width);
    const double allowed = area.width_pt * pages;
    factor = std::min(factor, allowed / (total_width_pt + title_width_pt * pages));
  }
  if (setup.fit_to_height > 0 && total_height_pt > 0.0 && area.height_pt > 0.0) {
    const double pages = static_cast<double>(setup.fit_to_height);
    const double allowed = area.height_pt * pages;
    factor = std::min(factor, allowed / (total_height_pt + title_height_pt * pages));
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
  } else {
    // An explicit print area paginates as declared. Intersecting it with
    // the populated box used to look right on cases whose content reaches
    // the area's corners, but `print_matrix.density_col_sparse_only_a`
    // (content in column A only, print area A1:H30) shows Excel breaking
    // at the same columns as the dense case: the declared geometry drives
    // the page grid, not where the cells happen to be.
    //
    // An edge declared at the grid limit is the one exception. A
    // whole-column (`$A:$A`) or whole-row (`$1:$1`) area names no
    // geometry on that axis, and honouring it literally would size the
    // track vectors below -- and the per-area walk that follows them --
    // to all 1,048,576 rows / 16,384 columns. That is multiple megabytes
    // of transient allocation on a routine print area, and
    // `kMaxPaginationPages` cannot hold it back because that check runs
    // after the walk it would have to prevent. Excel trims such an edge
    // to the content, so clip it to the populated box; on a sheet with
    // no populated cell there is nothing to trim to and the rectangle
    // drops out entirely. A rectangle left empty by the clip (its
    // declared start lies past the content) contributes no pages, and a
    // sheet whose rectangles all drop out produces none.
    constexpr std::uint32_t kMaxRowIndex = Sheet::kMaxRows - 1U;
    constexpr std::uint32_t kMaxColIndex = Sheet::kMaxCols - 1U;
    for (const CellRange& r : result.print_area) {
      CellRange clipped = r;
      if (r.last_row >= kMaxRowIndex) {
        if (!has_used_range) {
          continue;
        }
        clipped.last_row = used_box.last_row;
      }
      if (r.last_col >= kMaxColIndex) {
        if (!has_used_range) {
          continue;
        }
        clipped.last_col = used_box.last_col;
      }
      if (clipped.first_row > clipped.last_row || clipped.first_col > clipped.last_col) {
        continue;
      }
      effective_areas.push_back(clipped);
    }
    if (effective_areas.empty()) {
      return result;
    }
  }

  // Build the resolved track geometry once for the union extent. The old
  // per-track RowHeightPoints / ColumnWidthChars calls each re-scanned every
  // layout override, turning a 50k-row pagination into O(rows * overrides).
  // An override that carries only outline / hidden metadata has no `ht`, so
  // it must retain the sheet default height rather than becoming a 0pt row.
  const CellRange union_box = BoundingBox(effective_areas);
  const SheetFormatDefaults& defaults = sheet.format_defaults();
  const double default_row_h = defaults.has_default_row_height ? defaults.default_row_height : kStandardRowHeightPt;
  const double default_col_w = defaults.has_default_col_width ? defaults.default_col_width : kStandardColWidthChars;
  std::vector<double> row_heights(static_cast<std::size_t>(union_box.last_row - union_box.first_row) + 1U,
                                  default_row_h);
  std::vector<bool> row_overridden(row_heights.size(), false);
  for (const RowLayout& layout : sheet.layout().row_overrides) {
    if (layout.row < union_box.first_row || layout.row > union_box.last_row) {
      continue;
    }
    const std::size_t index = static_cast<std::size_t>(layout.row - union_box.first_row);
    if (row_overridden[index]) {
      continue;
    }
    row_overridden[index] = true;
    row_heights[index] =
        layout.hidden ? 0.0 : (layout.has_height || layout.height != 0.0 ? layout.height : default_row_h);
  }
  std::vector<double> col_widths(static_cast<std::size_t>(union_box.last_col - union_box.first_col) + 1U,
                                 default_col_w);
  std::vector<bool> col_overridden(col_widths.size(), false);
  for (const ColumnLayout& layout : sheet.layout().columns) {
    const std::uint32_t first = std::max(layout.first, union_box.first_col);
    const std::uint32_t last = std::min(layout.last, union_box.last_col);
    if (first > last) {
      continue;
    }
    for (std::uint32_t col = first; col <= last; ++col) {
      const std::size_t index = static_cast<std::size_t>(col - union_box.first_col);
      if (col_overridden[index]) {
        continue;
      }
      col_overridden[index] = true;
      col_widths[index] = layout.hidden ? 0.0 : (HasExplicitColumnWidth(layout) ? layout.width : default_col_w);
    }
  }
  const auto row_height = [&](std::uint32_t row) {
    if (row >= union_box.first_row && row <= union_box.last_row) {
      return row_heights[static_cast<std::size_t>(row - union_box.first_row)];
    }
    return RowHeightPoints(sheet, row);
  };
  const auto col_width = [&](std::uint32_t col) {
    if (col >= union_box.first_col && col <= union_box.last_col) {
      return col_widths[static_cast<std::size_t>(col - union_box.first_col)];
    }
    return ColumnWidthChars(sheet, col);
  };

  // 2. Printable body area is determined once from the page setup.
  const SheetPrintSettings& settings = sheet.print_settings();
  PrintableArea body = compute_printable_area(settings.page_setup, settings.page_margins);

  // Print titles (repeat-rows / repeat-columns) are reprinted on every
  // page, so they steal body extent that is otherwise available for
  // data rows / columns. Their sizes are summed here in model points and
  // converted below, once the scale is known.
  auto titles_or = resolve_print_titles(wb, sheet_index);
  if (!titles_or) {
    return titles_or.error();
  }
  const PrintTitles& titles = titles_or.value();
  double title_height = 0.0;
  if (titles.repeat_rows.has_value()) {
    const auto [first, last] = *titles.repeat_rows;
    for (std::uint32_t row = first; row <= last; ++row) {
      title_height += row_height(row);
    }
    // Empirical: when print-title rows are enabled, Excel reserves a
    // minimum body band of ~5 default rows even if the actual title
    // rows sum to less. Round-3 capture: print_titles_repeat_{1,3,5}_rows
    // on A1:D40 all break at h=[39], so the subtraction cannot be the
    // raw title_height (which would yield 15 / 45 / 75 pt). Using a
    // 5-default-row floor reproduces the observed h=[39] for all three.
    // The floor is a model-space quantity like the row heights it
    // stands in for, so it is applied before the scale conversion.
    constexpr double kMinTitleReserveRows = 5.0;
    title_height = std::max(title_height, kMinTitleReserveRows * default_row_h);
  }
  double title_width = 0.0;
  if (titles.repeat_cols.has_value()) {
    const auto [first, last] = *titles.repeat_cols;
    for (std::uint32_t col = first; col <= last; ++col) {
      title_width += ColumnCharsToPoints(col_width(col));
    }
  }

  // 3. The model scale factor is shared across every effective area:
  // Excel applies a single sheet-wide `<pageSetup scale>` / `fitToPage`
  // factor, so we size the factor against the union bounding box.
  // Repeated titles are part of what gets scaled onto the page, so a
  // fit factor has to accommodate them alongside the data.
  double union_total_width = 0.0;
  for (double width : col_widths) {
    union_total_width += ColumnCharsToPoints(width);
  }
  double union_total_height = 0.0;
  for (double height : row_heights) {
    union_total_height += height;
  }
  const double scale =
      ComputeScaleFactor(settings.page_setup, body, union_total_width, union_total_height, title_width, title_height);

  // The body is physical page space (paper minus margins); a track is a
  // model size that reaches the page multiplied by `scale`. Repeated
  // titles are printed content and shrink with everything else, so the
  // space they claim on the page is their scaled size. Subtracting the
  // raw model size instead would reserve `1 / scale` times too much
  // band and push data onto later pages -- at scale=50 a five-row title
  // block would claim the page space of ten.
  body.height_pt = std::max(0.0, body.height_pt - title_height * scale);
  body.width_pt = std::max(0.0, body.width_pt - title_width * scale);

  // 4. Paginate each rectangle independently and aggregate the breaks.
  // Excel's `HPageBreaks` / `VPageBreaks` collections report the union
  // of breaks across all areas; the page count sums across areas.
  std::vector<std::uint32_t> all_h;
  std::vector<std::uint32_t> all_v;
  std::vector<std::uint32_t> manual_h;
  std::vector<std::uint32_t> manual_v;
  // Accumulated in 64 bits: one axis can produce a page per grid track, so
  // the per-area product reaches 2^34 and the sum across areas grows past
  // that again. A 32-bit accumulator would wrap and report a count that is
  // simply wrong, with no diagnostic.
  std::uint64_t total_pages = 0;
  for (const CellRange& rect : effective_areas) {
    // Row axis: automatic overflow breaks plus manual row breaks within the
    // area's rows, counted per area.
    std::vector<double> row_points;
    row_points.reserve(rect.last_row - rect.first_row + 1);
    for (std::uint32_t row = rect.first_row; row <= rect.last_row; ++row) {
      row_points.push_back(row_height(row) * scale);
    }
    AxisInput row_axis;
    row_axis.first = rect.first_row;
    row_axis.track_sizes = std::move(row_points);
    row_axis.manual = &settings.manual_row_breaks;
    row_axis.limit_pt = body.height_pt;
    const std::uint32_t row_pages = WalkAxis(row_axis, &all_h, &manual_h);

    // Column axis: symmetric per-area walk through the same axis walker,
    // against the body width exactly as the row axis walks the body height.
    // A wide print area wraps onto further page-columns; it is not clipped
    // at the right margin. Using the same per-area walker as the row axis
    // also makes multi-area column pagination symmetric with multi-area row
    // pagination: each area's breaks are counted within that area rather
    // than de-duplicated across areas.
    std::vector<double> col_points;
    col_points.reserve(rect.last_col - rect.first_col + 1);
    for (std::uint32_t col = rect.first_col; col <= rect.last_col; ++col) {
      col_points.push_back(ColumnCharsToPoints(col_width(col)) * scale);
    }
    AxisInput col_axis;
    col_axis.first = rect.first_col;
    col_axis.track_sizes = std::move(col_points);
    col_axis.manual = &settings.manual_col_breaks;
    col_axis.limit_pt = body.width_pt;
    const std::uint32_t col_pages = WalkAxis(col_axis, &all_v, &manual_v);

    total_pages += static_cast<std::uint64_t>(col_pages) * static_cast<std::uint64_t>(row_pages);
    if (total_pages > kMaxPaginationPages) {
      return make_error(FormulonErrorCode::kPrintPageCountOverflow, "Pagination page count exceeds the supported limit",
                        "pages=" + std::to_string(total_pages) + " limit=" + std::to_string(kMaxPaginationPages));
    }
  }

  // 5. Sort the aggregated break positions. Excel's COM collections are
  // ascending but do repeat across areas: `print_pagination.
  // multi_area_row_stacked_col_break` (two stacked areas that each break
  // before column H) reports v=[3,7,7], one entry per area, so
  // de-duplicating here dropped a break Excel reports.
  auto merge_manual = [](std::vector<std::uint32_t>* automatic, std::vector<std::uint32_t>* manual) {
    std::sort(manual->begin(), manual->end());
    manual->erase(std::unique(manual->begin(), manual->end()), manual->end());
    automatic->insert(automatic->end(), manual->begin(), manual->end());
    std::sort(automatic->begin(), automatic->end());
  };
  merge_manual(&all_h, &manual_h);
  merge_manual(&all_v, &manual_v);

  result.h_breaks = std::move(all_h);
  result.v_breaks = std::move(all_v);
  // Bounded by `kMaxPaginationPages` above, so the narrowing is exact.
  result.page_count = static_cast<std::uint32_t>(total_pages);
  return result;
}

}  // namespace print
}  // namespace formulon
