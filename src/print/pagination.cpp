// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.

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
      return span.width;
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
      return override_row.height;
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

/// True when `index` is the target of a manual break in `breaks`.
bool HasManualBreakAt(const std::vector<ManualBreak>& breaks, std::uint32_t index) {
  return std::any_of(breaks.begin(), breaks.end(), [index](const ManualBreak& brk) { return brk.id == index; });
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

  // 1. Resolve the print area; fall back to the used range.
  auto area_or = resolve_print_area(wb, sheet_index);
  if (!area_or) {
    return area_or.error();
  }
  PaginationResult result;
  result.print_area = area_or.value();
  if (result.print_area.empty()) {
    CellRange used;
    if (!ComputeUsedRange(sheet, &used)) {
      // An empty sheet with no print area produces no pages.
      return result;
    }
    result.print_area.push_back(used);
  }

  const CellRange box = BoundingBox(result.print_area);

  // 2 & 3. Size every column and row of the bounding box, in points.
  std::vector<double> col_points;
  col_points.reserve(box.last_col - box.first_col + 1);
  double total_width = 0.0;
  for (std::uint32_t col = box.first_col; col <= box.last_col; ++col) {
    const double pts = ColumnCharsToPoints(ColumnWidthChars(sheet, col));
    col_points.push_back(pts);
    total_width += pts;
  }
  std::vector<double> row_points;
  row_points.reserve(box.last_row - box.first_row + 1);
  double total_height = 0.0;
  for (std::uint32_t row = box.first_row; row <= box.last_row; ++row) {
    const double pts = RowHeightPoints(sheet, row);
    row_points.push_back(pts);
    total_height += pts;
  }

  // 4. Printable body area and the model scale factor.
  const SheetPrintSettings& settings = sheet.print_settings();
  const PrintableArea body = compute_printable_area(settings.page_setup, settings.page_margins);
  const double scale = ComputeScaleFactor(settings.page_setup, body, total_width, total_height);

  // Apply the scale to cell sizes (a factor below 1.0 shrinks cells so
  // more fit per page).
  for (double& v : col_points) {
    v *= scale;
  }
  for (double& v : row_points) {
    v *= scale;
  }

  // 5. Walk both axes, honouring manual breaks.
  AxisInput col_axis;
  col_axis.first = box.first_col;
  col_axis.track_sizes = std::move(col_points);
  col_axis.manual = &settings.manual_col_breaks;
  col_axis.limit_pt = body.width_pt;
  const std::uint32_t col_pages = WalkAxis(col_axis, &result.v_breaks);

  AxisInput row_axis;
  row_axis.first = box.first_row;
  row_axis.track_sizes = std::move(row_points);
  row_axis.manual = &settings.manual_row_breaks;
  row_axis.limit_pt = body.height_pt;
  const std::uint32_t row_pages = WalkAxis(row_axis, &result.h_breaks);

  // 6. Total page count.
  result.page_count = col_pages * row_pages;
  return result;
}

}  // namespace print
}  // namespace formulon
