#include "print/sheet_geometry.h"

#include <algorithm>
#include <cmath>
#include <unordered_set>
#include <utility>

#include "sheet.h"
#include "styles.h"

namespace formulon {
namespace print {
namespace {

/// Height of one row override: hidden is zero, an explicit height wins,
/// anything else keeps the sheet default.
double OverrideRowHeight(const RowLayout& layout, double default_height) {
  if (layout.hidden) {
    return 0.0;
  }
  return (layout.has_height || layout.height != 0.0) ? layout.height : default_height;
}

/// Width of one column span in character units: hidden is zero, an explicit
/// width wins, anything else keeps the sheet default.
double SpanColumnWidth(const ColumnLayout& span, double default_width) {
  if (span.hidden) {
    return 0.0;
  }
  return HasExplicitColumnWidth(span) ? span.width : default_width;
}

}  // namespace

double default_row_height_pt(const Sheet& sheet) {
  const SheetFormatDefaults& defaults = sheet.format_defaults();
  return defaults.has_default_row_height ? defaults.default_row_height : ooxml_defaults::kStandardRowHeightPt;
}

/// The nominal default remains available to callers that need the stored
/// height (for example, an explicitly visible row without its own `ht`). The
/// implicit extent is zero when `<sheetFormatPr zeroHeight="1">` hides rows by
/// default.
double implicit_row_height_pt(const Sheet& sheet) {
  return sheet.format_defaults().zero_height ? 0.0 : default_row_height_pt(sheet);
}

double effective_column_width_chars(const Sheet& sheet, std::uint32_t col, const ColumnWidthModel& model) {
  const double default_width = default_column_width_chars(sheet, model);
  for (const ColumnLayout& span : sheet.layout().columns) {
    if (col >= span.first && col <= span.last) {
      // Hidden columns occupy no printed width, so they never advance the
      // page grid.
      return SpanColumnWidth(span, default_width);
    }
  }
  return default_width;
}

double effective_row_height_pt(const Sheet& sheet, std::uint32_t row) {
  const double default_height = default_row_height_pt(sheet);
  for (const RowLayout& override_row : sheet.layout().row_overrides) {
    if (override_row.row == row) {
      return OverrideRowHeight(override_row, default_height);
    }
  }
  return implicit_row_height_pt(sheet);
}

std::vector<double> column_widths_chars(const Sheet& sheet, std::uint32_t first, std::uint32_t last,
                                        const ColumnWidthModel& model) {
  const double default_width = default_column_width_chars(sheet, model);
  std::vector<double> widths(static_cast<std::size_t>(last - first) + 1U, default_width);
  // The first span covering a column wins. `next[i]` leads to the first
  // unassigned index at or after `i`, so overlapping spans skip assigned
  // runs instead of walking them: O(columns + spans) for any span set.
  std::vector<std::size_t> next(widths.size() + 1U);
  for (std::size_t i = 0; i < next.size(); ++i) {
    next[i] = i;
  }
  const auto unassigned_from = [&next](std::size_t i) {
    while (next[i] != i) {
      next[i] = next[next[i]];
      i = next[i];
    }
    return i;
  };
  for (const ColumnLayout& span : sheet.layout().columns) {
    if (span.last < first || span.first > last || span.first > span.last) {
      continue;
    }
    const auto span_end = static_cast<std::size_t>(std::min(span.last, last) - first);
    for (std::size_t index = unassigned_from(static_cast<std::size_t>(std::max(span.first, first) - first));
         index <= span_end; index = unassigned_from(index + 1U)) {
      widths[index] = SpanColumnWidth(span, default_width);
      next[index] = index + 1U;
    }
    if (unassigned_from(0U) == widths.size()) {
      break;
    }
  }
  return widths;
}

std::vector<double> row_heights_pt(const Sheet& sheet, std::uint32_t first, std::uint32_t last) {
  const double default_height = default_row_height_pt(sheet);
  std::vector<double> heights(static_cast<std::size_t>(last - first) + 1U, implicit_row_height_pt(sheet));
  std::vector<bool> assigned(heights.size(), false);
  for (const RowLayout& layout : sheet.layout().row_overrides) {
    if (layout.row < first || layout.row > last) {
      continue;
    }
    const auto index = static_cast<std::size_t>(layout.row - first);
    if (assigned[index]) {
      continue;
    }
    assigned[index] = true;
    heights[index] = OverrideRowHeight(layout, default_height);
  }
  return heights;
}

double effective_column_width_pt(const Sheet& sheet, std::uint32_t col, const ColumnWidthModel& model) {
  return column_chars_to_points(effective_column_width_chars(sheet, col, model), model);
}

double column_offset_pt(const Sheet& sheet, std::uint32_t col, const ColumnWidthModel& model) {
  if (col == 0U) {
    return 0.0;
  }
  double offset = 0.0;
  for (const double chars : column_widths_chars(sheet, 0U, col - 1U, model)) {
    offset += column_chars_to_points(chars, model);
  }
  return offset;
}

double row_offset_pt(const Sheet& sheet, std::uint32_t row) {
  // Rows are numerous (up to 2^20), so adjust the default-height run by the
  // overrides instead of materialising every row.
  const double default_height = default_row_height_pt(sheet);
  const double implicit_height = implicit_row_height_pt(sheet);
  double offset = implicit_height * static_cast<double>(row);
  if (row == 0U) {
    return offset;
  }
  const auto& overrides = sheet.layout().row_overrides;
  const bool ordered =
      std::adjacent_find(overrides.begin(), overrides.end(), [](const RowLayout& lhs, const RowLayout& rhs) {
        return lhs.row >= rhs.row;
      }) == overrides.end();
  if (ordered) {
    for (const RowLayout& layout : overrides) {
      if (layout.row >= row) {
        break;
      }
      offset += OverrideRowHeight(layout, default_height) - implicit_height;
    }
    return offset;
  }
  // Preserve first-override precedence for directly constructed layouts
  // that bypass the sorted, unique row-edit helpers.
  std::unordered_set<std::uint32_t> seen;
  seen.reserve(overrides.size());
  for (const RowLayout& layout : overrides) {
    if (layout.row >= row) {
      continue;
    }
    if (!seen.insert(layout.row).second) {
      continue;
    }
    offset += OverrideRowHeight(layout, default_height) - implicit_height;
  }
  return offset;
}

RectPt cell_rect_pt(const Sheet& sheet, std::uint32_t row, std::uint32_t col, const ColumnWidthModel& model) {
  RectPt rect;
  rect.x = column_offset_pt(sheet, col, model);
  rect.y = row_offset_pt(sheet, row);
  rect.width = effective_column_width_pt(sheet, col, model);
  rect.height = effective_row_height_pt(sheet, row);
  return rect;
}

}  // namespace print
}  // namespace formulon
