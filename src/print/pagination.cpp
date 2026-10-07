
#include "print/pagination.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <string>
#include <vector>

#include "cell.h"
#include "print/page_setup.h"
#include "print/print_area.h"
#include "print/sheet_geometry.h"
#include "sheet.h"
#include "styles.h"
#include "utils/index_sort.h"
#include "utils/resource_budget.h"
#include "workbook.h"

namespace formulon {
namespace print {
namespace {

/// Lower clamp for the effective scale factor (1%). Mirrors Excel's
/// `<pageSetup scale>` minimum so a degenerate `scale="0"` cannot divide
/// the printable area by zero.
constexpr double kMinScaleFactor = 0.01;

/// Divisor turning the `<pageSetup scale>` percentage into a factor.
constexpr double kPercentDivisor = 100.0;

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
/// at least one track). `out_page_starts` receives the absolute index of
/// each page's first track, in ascending order.
std::uint32_t WalkAxis(const AxisInput& axis, std::vector<std::uint32_t>* out_breaks,
                       std::vector<std::uint32_t>* out_manual_breaks, std::vector<std::uint32_t>* out_page_starts) {
  const std::size_t track_count = axis.track_sizes.size();
  if (track_count == 0) {
    return 0;
  }

  std::uint32_t pages = 1;
  out_page_starts->push_back(axis.first);
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
      out_page_starts->push_back(absolute);
      accumulated = 0.0;
    }
    accumulated += size;
  }
  return pages;
}

/// The track range one page covers on one axis, with its scaled extent.
struct PageSpan {
  std::uint32_t first = 0;
  std::uint32_t last = 0;
  double extent_pt = 0.0;
};

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

  // Break math needs the print-layout geometry, not the display one.
  const ColumnWidthModel column_geometry = resolve_column_width_model(wb.styles(), GeometryMode::kPrint);

  // 1. Resolve the print area. The reported `result.print_area` mirrors
  // Excel's `PageSetup.PrintArea` exactly: empty when the workbook
  // defines no print area, even if the sheet has populated cells. The
  // used-range fallback is a *pagination* concern only, not a reporting
  // one, so it lives in a separate `effective_area` used solely to size
  // and walk the page grid below.
  auto area_or = resolve_print_area(wb, sheet_index);
  if (!area_or) {
    return std::move(area_or.error());
  }
  PaginationResult result;
  result.print_area = area_or.value();

  // Page description that depends only on the page setup, so it is reported
  // even for a sheet with nothing to print.
  const SheetPrintSettings& settings = sheet.print_settings();
  const PageSetup& page_setup = settings.page_setup;
  const PageMargins& page_margins = settings.page_margins;
  {
    const PaperDimensions paper = resolve_paper_dimensions(page_setup.paper_size);
    result.paper.landscape = page_setup.orientation == Orientation::kLandscape;
    result.paper.width_pt = result.paper.landscape ? paper.height_pt : paper.width_pt;
    result.paper.height_pt = result.paper.landscape ? paper.width_pt : paper.height_pt;
    result.paper.known = is_known_paper_size(page_setup.paper_size);
    result.margins = MarginsPt{page_margins.left * kPointsPerInch,   page_margins.right * kPointsPerInch,
                               page_margins.top * kPointsPerInch,    page_margins.bottom * kPointsPerInch,
                               page_margins.header * kPointsPerInch, page_margins.footer * kPointsPerInch};
    const PrintableArea printable = compute_printable_area(page_setup, page_margins);
    result.printable = RectPt{result.margins.left, result.margins.top, printable.width_pt, printable.height_pt};
    result.page_order = settings.page_setup.page_order;
    result.scale = page_setup.fit_to_page
                       ? 1.0
                       : std::max(kMinScaleFactor, static_cast<double>(page_setup.scale) / kPercentDivisor);
  }

  // Excel's `HPageBreaks` / `VPageBreaks` are populated against the
  // *populated* region of the print area, not its full geometric span:
  // a print area whose four corners are blank effectively paginates
  // against the (possibly much smaller) populated bounding box. Compute
  // the populated box once so each print-area rectangle can be
  // intersected with it below.
  CellRange used_box;
  const bool has_used_range = ComputeUsedRange(sheet, &used_box);

  std::vector<CellRange> effective_areas;
  // Index into `result.print_area` each effective area came from.
  std::vector<std::uint32_t> area_indices;
  if (result.print_area.empty()) {
    // No explicit print area: fall back to the used range. An empty
    // sheet with no print area produces no pages.
    if (!has_used_range) {
      return result;
    }
    effective_areas.push_back(used_box);
    area_indices.push_back(0U);
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
    for (std::size_t area_index = 0; area_index < result.print_area.size(); ++area_index) {
      const CellRange& r = result.print_area[area_index];
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
      area_indices.push_back(static_cast<std::uint32_t>(area_index));
    }
    if (effective_areas.empty()) {
      return result;
    }
  }

  // Build the resolved track geometry once for the union extent. The old
  // per-track effective_row_height_pt / effective_column_width_chars calls each re-scanned every
  // layout override, turning a 50k-row pagination into O(rows * overrides).
  // An override that carries only outline / hidden metadata has no `ht`, so
  // it must retain the sheet default height rather than becoming a 0pt row.
  const CellRange union_box = BoundingBox(effective_areas);
  const double default_row_h = default_row_height_pt(sheet);
  const std::vector<double> row_heights = row_heights_pt(sheet, union_box.first_row, union_box.last_row);
  const std::vector<double> col_widths = column_widths_chars(sheet, union_box.first_col, union_box.last_col);
  const auto row_height = [&](std::uint32_t row) {
    if (row >= union_box.first_row && row <= union_box.last_row) {
      return row_heights[static_cast<std::size_t>(row - union_box.first_row)];
    }
    return effective_row_height_pt(sheet, row);
  };
  const auto col_width = [&](std::uint32_t col) {
    if (col >= union_box.first_col && col <= union_box.last_col) {
      return col_widths[static_cast<std::size_t>(col - union_box.first_col)];
    }
    return effective_column_width_chars(sheet, col);
  };

  // 2. Printable body area is determined once from the page setup.
  PrintableArea body = compute_printable_area(settings.page_setup, settings.page_margins);

  // Print titles (repeat-rows / repeat-columns) are reprinted on every
  // page, so they steal body extent that is otherwise available for
  // data rows / columns. Their sizes are summed here in model points and
  // converted below, once the scale is known.
  auto titles_or = resolve_print_titles(wb, sheet_index);
  if (!titles_or) {
    return std::move(titles_or.error());
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
      title_width += column_chars_to_points(col_width(col), column_geometry);
    }
  }

  // 3. The model scale factor is shared across every effective area:
  // Excel applies a single sheet-wide `<pageSetup scale>` / `fitToPage`
  // factor, so we size the factor against the union bounding box.
  // Repeated titles are part of what gets scaled onto the page, so a
  // fit factor has to accommodate them alongside the data.
  double union_total_width = 0.0;
  for (double width : col_widths) {
    union_total_width += column_chars_to_points(width, column_geometry);
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
  result.scale = scale;
  if (titles.repeat_rows.has_value()) {
    result.print_titles.has_rows = true;
    result.print_titles.first_row = titles.repeat_rows->first;
    result.print_titles.last_row = titles.repeat_rows->second;
  }
  if (titles.repeat_cols.has_value()) {
    result.print_titles.has_cols = true;
    result.print_titles.first_col = titles.repeat_cols->first;
    result.print_titles.last_col = titles.repeat_cols->second;
  }
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
    std::vector<std::uint32_t> row_starts;
    const std::uint32_t row_pages = WalkAxis(row_axis, &all_h, &manual_h, &row_starts);

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
      col_points.push_back(column_chars_to_points(col_width(col), column_geometry) * scale);
    }
    AxisInput col_axis;
    col_axis.first = rect.first_col;
    col_axis.track_sizes = std::move(col_points);
    col_axis.manual = &settings.manual_col_breaks;
    col_axis.limit_pt = body.width_pt;
    std::vector<std::uint32_t> col_starts;
    const std::uint32_t col_pages = WalkAxis(col_axis, &all_v, &manual_v, &col_starts);

    total_pages += static_cast<std::uint64_t>(col_pages) * static_cast<std::uint64_t>(row_pages);
    if (total_pages > kMaxPaginationPages) {
      return make_error(FormulonErrorCode::kPrintPageCountOverflow, "Pagination page count exceeds the supported limit",
                        "pages=" + std::to_string(total_pages) + " limit=" + std::to_string(kMaxPaginationPages));
    }

    // Pages of this area: the walk's page starts bound each page's track
    // range. Print areas follow one another; within one, `page_order`
    // decides which axis runs fastest.
    const auto pages_of_axis = [](const AxisInput& axis, const std::vector<std::uint32_t>& starts,
                                  std::uint32_t area_last) {
      std::vector<PageSpan> spans;
      spans.reserve(starts.size());
      for (std::size_t i = 0; i < starts.size(); ++i) {
        PageSpan span;
        span.first = starts[i];
        span.last = i + 1 < starts.size() ? starts[i + 1] - 1U : area_last;
        for (std::uint32_t track = span.first; track <= span.last; ++track) {
          span.extent_pt += axis.track_sizes[track - axis.first];
        }
        spans.push_back(span);
      }
      return spans;
    };
    const std::vector<PageSpan> row_spans = pages_of_axis(row_axis, row_starts, rect.last_row);
    const std::vector<PageSpan> col_spans = pages_of_axis(col_axis, col_starts, rect.last_col);
    const auto emit_page = [&](const PageSpan& rows, const PageSpan& cols) {
      PageLayout page;
      page.area_index = area_indices[static_cast<std::size_t>(&rect - effective_areas.data())];
      page.first_row = rows.first;
      page.last_row = rows.last;
      page.first_col = cols.first;
      page.last_col = cols.last;
      // Titles are reprinted ahead of every page that does not already
      // contain them.
      const bool rows_offset = titles.repeat_rows.has_value() && rows.first > titles.repeat_rows->second;
      const bool cols_offset = titles.repeat_cols.has_value() && cols.first > titles.repeat_cols->second;
      page.origin_x_pt = result.printable.x + (cols_offset ? title_width * scale : 0.0);
      page.origin_y_pt = result.printable.y + (rows_offset ? title_height * scale : 0.0);
      page.width_pt = cols.extent_pt;
      page.height_pt = rows.extent_pt;
      result.pages.push_back(page);
    };
    if (result.page_order == PageOrder::kOverThenDown) {
      for (const PageSpan& rows : row_spans) {
        for (const PageSpan& cols : col_spans) {
          emit_page(rows, cols);
        }
      }
    } else {
      for (const PageSpan& cols : col_spans) {
        for (const PageSpan& rows : row_spans) {
          emit_page(rows, cols);
        }
      }
    }
  }

  // 5. Sort the aggregated break positions. Excel's COM collections are
  // ascending but do repeat across areas: `print_pagination.
  // multi_area_row_stacked_col_break` (two stacked areas that each break
  // before column H) reports v=[3,7,7], one entry per area, so
  // de-duplicating here dropped a break Excel reports.
  auto merge_manual = [](std::vector<std::uint32_t>* automatic, std::vector<std::uint32_t>* manual) {
    sort_ascending(*manual);
    manual->erase(std::unique(manual->begin(), manual->end()), manual->end());
    automatic->insert(automatic->end(), manual->begin(), manual->end());
    sort_ascending(*automatic);
  };
  merge_manual(&all_h, &manual_h);
  merge_manual(&all_v, &manual_v);

  // Flag the manual entries. A value that is also an automatic break of
  // another area appears more than once; the first occurrence is the manual one.
  const auto flag_manual = [](const std::vector<std::uint32_t>& all, const std::vector<std::uint32_t>& manual) {
    std::vector<bool> flags(all.size(), false);
    for (const std::uint32_t value : manual) {
      flags[static_cast<std::size_t>(std::lower_bound(all.begin(), all.end(), value) - all.begin())] = true;
    }
    return flags;
  };
  result.h_break_manual = flag_manual(all_h, manual_h);
  result.v_break_manual = flag_manual(all_v, manual_v);

  result.h_breaks = std::move(all_h);
  result.v_breaks = std::move(all_v);
  // Bounded by `kMaxPaginationPages` above, so the narrowing is exact.
  result.page_count = static_cast<std::uint32_t>(total_pages);
  return result;
}

}  // namespace print
}  // namespace formulon
