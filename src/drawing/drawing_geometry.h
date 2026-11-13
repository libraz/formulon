#ifndef FORMULON_DRAWING_DRAWING_GEOMETRY_H_
#define FORMULON_DRAWING_DRAWING_GEOMETRY_H_

#include <cmath>
#include <cstdint>
#include <vector>

#include "column_width.h"
#include "drawing/drawing_xml.h"
#include "print/sheet_geometry.h"
#include "sheet.h"
#include "workbook.h"

namespace formulon {

inline constexpr std::int64_t kEmuPerPoint = 12700;

// Cell extents in EMU under the Windows display geometry. Column widths are
// resolved once for the whole grid, so walking columns costs O(columns +
// spans) however many `<col>` spans the sheet carries.
class Geometry {
 public:
  Geometry(const Workbook& wb, const Sheet& sheet)
      : sheet_(sheet),
        model_(resolve_column_width_model(wb, wb.excel_profile(), GeometryMode::kDisplay)),
        column_chars_(print::column_widths_chars(sheet, 0U, Sheet::kMaxCols - 1U, model_)) {}

  std::int64_t size(bool row_axis, std::uint32_t index) const {
    const double pt = row_axis ? print::effective_row_height_pt(sheet_, index) : column_pt(index);
    return std::llround(pt * kEmuPerPoint);
  }

  std::int64_t offset(bool row_axis, std::uint32_t index) const {
    if (row_axis) {
      return std::llround(print::row_offset_pt(sheet_, index) * kEmuPerPoint);
    }
    // Summed in the order `print::column_offset_pt` sums them.
    double pt = 0.0;
    for (std::uint32_t col = 0; col < index && col < column_chars_.size(); ++col) {
      pt += column_pt(col);
    }
    return std::llround(pt * kEmuPerPoint);
  }

  // EMU from marker `from` to marker `to` along the axis, summed line by line
  // as `advance` walks them.
  std::int64_t span(bool row_axis, const AnchorPoint& from, const AnchorPoint& to) const {
    const std::uint32_t first = row_axis ? from.row : from.col;
    const std::uint32_t last = row_axis ? to.row : to.col;
    std::int64_t extent = (row_axis ? to.row_off : to.col_off) - (row_axis ? from.row_off : from.col_off);
    for (std::uint32_t line = first; line < last; ++line) {
      extent += size(row_axis, line);
    }
    return extent;
  }

  // Moves the marker (`index`, `off`) `extent` EMU further along the axis.
  // Returns false when the result lies past the sheet's last line.
  bool advance(bool row_axis, std::uint32_t& index, std::int64_t& off, std::int64_t extent) const {
    const std::uint32_t limit = a1::axis_extent(row_axis);
    std::int64_t rem = off + extent;
    for (std::int64_t s = size(row_axis, index); rem >= s; s = size(row_axis, index)) {
      if (index + 1 >= limit) {
        off = rem;
        return rem <= s;
      }
      rem -= s;
      ++index;
    }
    off = rem;
    return true;
  }

 private:
  double column_pt(std::uint32_t col) const {
    // An anchor read from a file may name a column past the grid; it has the default width.
    return column_chars_to_points(
        col < column_chars_.size() ? column_chars_[col] : default_column_width_chars(sheet_, model_), model_);
  }

  const Sheet& sheet_;
  ColumnWidthModel model_;
  std::vector<double> column_chars_;
};

}  // namespace formulon

#endif  // FORMULON_DRAWING_DRAWING_GEOMETRY_H_
