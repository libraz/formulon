//
// C ABI <-> model enum conversion.
//
// The C enums follow the binding families (one value per member name across
// a family, `-1` for an absent element), while the model enums keep the order
// the readers and writers use. Every crossing goes through a switch here, so
// a reordered C enum never leaks into the model or back. Each `*_from_c`
// expects a value its caller has already checked with `check_enum_domain`.

#ifndef FORMULON_C_API_PARTS_ENUM_MAP_H_
#define FORMULON_C_API_PARTS_ENUM_MAP_H_

#include <cstdint>

#include "c_api/formulon_c.h"
#include "cf/cf_types.h"
#include "chart/chart_model.h"
#include "drawing/drawing_xml.h"
#include "drawing/shape_edit.h"
#include "drawing/slicer_edit.h"
#include "io/sparkline_groups.h"

namespace formulon {
namespace c_api {
namespace parts {

// --- anchor family -----------------------------------------------------------

inline AnchorKind anchor_kind_from_c(std::int32_t value) {
  switch (value) {
    case FM_ANCHOR_KIND_TWO_CELL:
      return AnchorKind::kTwoCell;
    case FM_ANCHOR_KIND_ABSOLUTE:
      return AnchorKind::kAbsolute;
    default:
      return AnchorKind::kOneCell;
  }
}

inline std::int32_t to_c(AnchorKind value) {
  switch (value) {
    case AnchorKind::kTwoCell:
      return FM_ANCHOR_KIND_TWO_CELL;
    case AnchorKind::kAbsolute:
      return FM_ANCHOR_KIND_ABSOLUTE;
    case AnchorKind::kOneCell:
      break;
  }
  return FM_ANCHOR_KIND_ONE_CELL;
}

inline EditAs edit_as_from_c(std::int32_t value) {
  switch (value) {
    case FM_ANCHOR_EDIT_AS_ONE_CELL:
      return EditAs::kOneCell;
    case FM_ANCHOR_EDIT_AS_ABSOLUTE:
      return EditAs::kAbsolute;
    default:
      return EditAs::kTwoCell;
  }
}

inline std::int32_t to_c(EditAs value) {
  switch (value) {
    case EditAs::kOneCell:
      return FM_ANCHOR_EDIT_AS_ONE_CELL;
    case EditAs::kAbsolute:
      return FM_ANCHOR_EDIT_AS_ABSOLUTE;
    case EditAs::kTwoCell:
      break;
  }
  return FM_ANCHOR_EDIT_AS_TWO_CELL;
}

// Copies the anchor kind and the two anchor points of drawing object `obj`
// into the record `out`; both share the `from_*` / `to_*` field names.
template <typename Object, typename Record>
void fill_anchor(const Object& obj, Record* out) {
  out->anchor_kind = to_c(obj.anchor_kind);
  out->from_row = obj.from.row;
  out->from_col = obj.from.col;
  out->from_row_offset_emu = obj.from.row_off;
  out->from_col_offset_emu = obj.from.col_off;
  out->to_row = obj.to.row;
  out->to_col = obj.to.col;
  out->to_row_offset_emu = obj.to.row_off;
  out->to_col_offset_emu = obj.to.col_off;
}

// --- shape paint (`-1` = no element) ----------------------------------------

inline ShapePaint shape_paint_from_c(std::int32_t value) {
  switch (value) {
    case FM_SHAPE_PAINT_SOLID:
      return ShapePaint::kSolid;
    case FM_SHAPE_PAINT_NONE:
      return ShapePaint::kNone;
    case FM_SHAPE_PAINT_OTHER:
      return ShapePaint::kOther;
    default:
      return ShapePaint::kUnset;
  }
}

inline std::int32_t to_c(ShapePaint value) {
  switch (value) {
    case ShapePaint::kSolid:
      return FM_SHAPE_PAINT_SOLID;
    case ShapePaint::kNone:
      return FM_SHAPE_PAINT_NONE;
    case ShapePaint::kOther:
      return FM_SHAPE_PAINT_OTHER;
    case ShapePaint::kUnset:
      break;
  }
  return FM_SHAPE_PAINT_UNSET;
}

// --- sort_order family -------------------------------------------------------

inline SlicerSortOrder slicer_sort_order_from_c(std::int32_t value) {
  return value == FM_SLICER_SORT_ORDER_DESCENDING ? SlicerSortOrder::kDescending : SlicerSortOrder::kAscending;
}

inline std::int32_t to_c(SlicerSortOrder value) {
  return value == SlicerSortOrder::kDescending ? FM_SLICER_SORT_ORDER_DESCENDING : FM_SLICER_SORT_ORDER_ASCENDING;
}

// --- blanks family (`-1` = no `c:dispBlanksAs`) -----------------------------

inline ChartBlanksAs chart_blanks_as_from_c(std::int32_t value) {
  switch (value) {
    case FM_CHART_BLANKS_AS_GAP:
      return ChartBlanksAs::kGap;
    case FM_CHART_BLANKS_AS_ZERO:
      return ChartBlanksAs::kZero;
    case FM_CHART_BLANKS_AS_SPAN:
      return ChartBlanksAs::kSpan;
    default:
      return ChartBlanksAs::kAbsent;
  }
}

inline std::int32_t to_c(ChartBlanksAs value) {
  switch (value) {
    case ChartBlanksAs::kGap:
      return FM_CHART_BLANKS_AS_GAP;
    case ChartBlanksAs::kZero:
      return FM_CHART_BLANKS_AS_ZERO;
    case ChartBlanksAs::kSpan:
      return FM_CHART_BLANKS_AS_SPAN;
    case ChartBlanksAs::kAbsent:
      break;
  }
  return FM_CHART_BLANKS_AS_ABSENT;
}

inline io::SparklineEmptyCells sparkline_empty_cells_from_c(std::int32_t value) {
  switch (value) {
    case FM_SPARKLINE_EMPTY_CELLS_ZERO:
      return io::SparklineEmptyCells::kZero;
    case FM_SPARKLINE_EMPTY_CELLS_SPAN:
      return io::SparklineEmptyCells::kSpan;
    default:
      return io::SparklineEmptyCells::kGap;
  }
}

inline std::int32_t to_c(io::SparklineEmptyCells value) {
  switch (value) {
    case io::SparklineEmptyCells::kZero:
      return FM_SPARKLINE_EMPTY_CELLS_ZERO;
    case io::SparklineEmptyCells::kSpan:
      return FM_SPARKLINE_EMPTY_CELLS_SPAN;
    case io::SparklineEmptyCells::kGap:
      break;
  }
  return FM_SPARKLINE_EMPTY_CELLS_GAP;
}

// --- side family -------------------------------------------------------------

inline std::int32_t to_c(ChartAxisPosition value) {
  switch (value) {
    case ChartAxisPosition::kTop:
      return FM_CHART_AXIS_POSITION_TOP;
    case ChartAxisPosition::kLeft:
      return FM_CHART_AXIS_POSITION_LEFT;
    case ChartAxisPosition::kRight:
      return FM_CHART_AXIS_POSITION_RIGHT;
    case ChartAxisPosition::kBottom:
      break;
  }
  return FM_CHART_AXIS_POSITION_BOTTOM;
}

// `FM_CHART_LEGEND_POSITION_NONE` is the patch instruction that removes the legend.
inline ChartLegendPosition chart_legend_position_from_c(std::int32_t value) {
  switch (value) {
    case FM_CHART_LEGEND_POSITION_BOTTOM:
      return ChartLegendPosition::kBottom;
    case FM_CHART_LEGEND_POSITION_TOP:
      return ChartLegendPosition::kTop;
    case FM_CHART_LEGEND_POSITION_LEFT:
      return ChartLegendPosition::kLeft;
    case FM_CHART_LEGEND_POSITION_RIGHT:
      return ChartLegendPosition::kRight;
    case FM_CHART_LEGEND_POSITION_TOP_RIGHT:
      return ChartLegendPosition::kTopRight;
    default:
      return ChartLegendPosition::kNone;
  }
}

inline std::int32_t to_c(ChartLegendPosition value) {
  switch (value) {
    case ChartLegendPosition::kBottom:
      return FM_CHART_LEGEND_POSITION_BOTTOM;
    case ChartLegendPosition::kTop:
      return FM_CHART_LEGEND_POSITION_TOP;
    case ChartLegendPosition::kLeft:
      return FM_CHART_LEGEND_POSITION_LEFT;
    case ChartLegendPosition::kRight:
      return FM_CHART_LEGEND_POSITION_RIGHT;
    case ChartLegendPosition::kTopRight:
      return FM_CHART_LEGEND_POSITION_TOP_RIGHT;
    case ChartLegendPosition::kNone:
      break;
  }
  return FM_CHART_LEGEND_POSITION_NONE;
}

// --- chart enums with an absent member ----------------------------------------

// `FM_CHART_TITLE_KIND_NONE` is the patch instruction that removes the title.
inline ChartTitleKind chart_title_kind_from_c(std::int32_t value) {
  switch (value) {
    case FM_CHART_TITLE_KIND_TEXT:
      return ChartTitleKind::kText;
    case FM_CHART_TITLE_KIND_REF:
      return ChartTitleKind::kRef;
    case FM_CHART_TITLE_KIND_AUTO:
      return ChartTitleKind::kAuto;
    default:
      return ChartTitleKind::kNone;
  }
}

inline std::int32_t to_c(ChartTitleKind value) {
  switch (value) {
    case ChartTitleKind::kText:
      return FM_CHART_TITLE_KIND_TEXT;
    case ChartTitleKind::kRef:
      return FM_CHART_TITLE_KIND_REF;
    case ChartTitleKind::kAuto:
      return FM_CHART_TITLE_KIND_AUTO;
    case ChartTitleKind::kNone:
      break;
  }
  return FM_CHART_TITLE_KIND_NONE;
}

inline ChartAxisCrosses chart_axis_crosses_from_c(std::int32_t value) {
  switch (value) {
    case FM_CHART_AXIS_CROSSES_AUTO_ZERO:
      return ChartAxisCrosses::kAutoZero;
    case FM_CHART_AXIS_CROSSES_MIN:
      return ChartAxisCrosses::kMin;
    case FM_CHART_AXIS_CROSSES_MAX:
      return ChartAxisCrosses::kMax;
    case FM_CHART_AXIS_CROSSES_AT:
      return ChartAxisCrosses::kAt;
    default:
      return ChartAxisCrosses::kAbsent;
  }
}

inline std::int32_t to_c(ChartAxisCrosses value) {
  switch (value) {
    case ChartAxisCrosses::kAutoZero:
      return FM_CHART_AXIS_CROSSES_AUTO_ZERO;
    case ChartAxisCrosses::kMin:
      return FM_CHART_AXIS_CROSSES_MIN;
    case ChartAxisCrosses::kMax:
      return FM_CHART_AXIS_CROSSES_MAX;
    case ChartAxisCrosses::kAt:
      return FM_CHART_AXIS_CROSSES_AT;
    case ChartAxisCrosses::kAbsent:
      break;
  }
  return FM_CHART_AXIS_CROSSES_ABSENT;
}

inline std::int32_t to_c(ChartBarDirection value) {
  switch (value) {
    case ChartBarDirection::kCol:
      return FM_CHART_BAR_DIRECTION_COLUMN;
    case ChartBarDirection::kBar:
      return FM_CHART_BAR_DIRECTION_BAR;
    case ChartBarDirection::kAbsent:
      break;
  }
  return FM_CHART_BAR_DIRECTION_ABSENT;
}

// --- data bar ------------------------------------------------------------------

inline cf::DataBarDirection data_bar_direction_from_c(std::int32_t value) {
  switch (value) {
    case FM_CF_DATA_BAR_DIRECTION_LEFT_TO_RIGHT:
      return cf::DataBarDirection::LeftToRight;
    case FM_CF_DATA_BAR_DIRECTION_RIGHT_TO_LEFT:
      return cf::DataBarDirection::RightToLeft;
    default:
      return cf::DataBarDirection::Context;
  }
}

inline std::uint8_t to_c(cf::DataBarDirection value) {
  switch (value) {
    case cf::DataBarDirection::LeftToRight:
      return FM_CF_DATA_BAR_DIRECTION_LEFT_TO_RIGHT;
    case cf::DataBarDirection::RightToLeft:
      return FM_CF_DATA_BAR_DIRECTION_RIGHT_TO_LEFT;
    case cf::DataBarDirection::Context:
      break;
  }
  return FM_CF_DATA_BAR_DIRECTION_CONTEXT;
}

inline cf::DataBarAxisPosition data_bar_axis_position_from_c(std::int32_t value) {
  switch (value) {
    case FM_CF_DATA_BAR_AXIS_POSITION_MIDDLE:
      return cf::DataBarAxisPosition::Middle;
    case FM_CF_DATA_BAR_AXIS_POSITION_NONE:
      return cf::DataBarAxisPosition::None;
    default:
      return cf::DataBarAxisPosition::Automatic;
  }
}

inline std::uint8_t to_c(cf::DataBarAxisPosition value) {
  switch (value) {
    case cf::DataBarAxisPosition::Middle:
      return FM_CF_DATA_BAR_AXIS_POSITION_MIDDLE;
    case cf::DataBarAxisPosition::None:
      return FM_CF_DATA_BAR_AXIS_POSITION_NONE;
    case cf::DataBarAxisPosition::Automatic:
      break;
  }
  return FM_CF_DATA_BAR_AXIS_POSITION_AUTOMATIC;
}

}  // namespace parts
}  // namespace c_api
}  // namespace formulon

#endif  // FORMULON_C_API_PARTS_ENUM_MAP_H_
