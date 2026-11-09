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

#include <array>
#include <cstddef>
#include <cstdint>
#include <string_view>

#include "auto_filter.h"
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
  return FM_SHAPE_PAINT_ABSENT;
}

// --- sort_order family -------------------------------------------------------

inline SlicerSortOrder slicer_sort_order_from_c(std::int32_t value) {
  return value == FM_SLICER_SORT_ORDER_DESCENDING ? SlicerSortOrder::kDescending : SlicerSortOrder::kAscending;
}

inline std::int32_t to_c(SlicerSortOrder value) {
  return value == SlicerSortOrder::kDescending ? FM_SLICER_SORT_ORDER_DESCENDING : FM_SLICER_SORT_ORDER_ASCENDING;
}

// --- blanks (`-1` = no `c:dispBlanksAs`) -----------------------------------

inline ChartBlanksAs chart_blanks_as_from_c(std::int32_t value) {
  switch (value) {
    case FM_BLANKS_AS_GAP:
      return ChartBlanksAs::kGap;
    case FM_BLANKS_AS_ZERO:
      return ChartBlanksAs::kZero;
    case FM_BLANKS_AS_SPAN:
      return ChartBlanksAs::kSpan;
    default:
      return ChartBlanksAs::kAbsent;
  }
}

inline std::int32_t to_c(ChartBlanksAs value) {
  switch (value) {
    case ChartBlanksAs::kGap:
      return FM_BLANKS_AS_GAP;
    case ChartBlanksAs::kZero:
      return FM_BLANKS_AS_ZERO;
    case ChartBlanksAs::kSpan:
      return FM_BLANKS_AS_SPAN;
    case ChartBlanksAs::kAbsent:
      break;
  }
  return FM_BLANKS_AS_ABSENT;
}

inline io::SparklineEmptyCells sparkline_empty_cells_from_c(std::int32_t value) {
  switch (value) {
    case FM_BLANKS_AS_GAP:
      return io::SparklineEmptyCells::kGap;
    case FM_BLANKS_AS_SPAN:
      return io::SparklineEmptyCells::kSpan;
    default:
      return io::SparklineEmptyCells::kZero;
  }
}

inline std::int32_t to_c(io::SparklineEmptyCells value) {
  switch (value) {
    case io::SparklineEmptyCells::kGap:
      return FM_BLANKS_AS_GAP;
    case io::SparklineEmptyCells::kSpan:
      return FM_BLANKS_AS_SPAN;
    case io::SparklineEmptyCells::kZero:
      break;
  }
  return FM_BLANKS_AS_ZERO;
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

// `FM_CHART_LEGEND_POSITION_ABSENT` is the patch instruction that removes the legend.
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
  return FM_CHART_LEGEND_POSITION_ABSENT;
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

// --- comparison operator (one C table, three model encodings) ----------------

inline cf::CellIsOperator cf_operator_from_c(std::int32_t value) {
  switch (value) {
    case FM_COMPARISON_OPERATOR_BETWEEN:
      return cf::CellIsOperator::Between;
    case FM_COMPARISON_OPERATOR_NOT_BETWEEN:
      return cf::CellIsOperator::NotBetween;
    case FM_COMPARISON_OPERATOR_NOT_EQUAL:
      return cf::CellIsOperator::NotEqual;
    case FM_COMPARISON_OPERATOR_LESS_THAN:
      return cf::CellIsOperator::LessThan;
    case FM_COMPARISON_OPERATOR_LESS_THAN_OR_EQUAL:
      return cf::CellIsOperator::LessThanOrEqual;
    case FM_COMPARISON_OPERATOR_GREATER_THAN:
      return cf::CellIsOperator::GreaterThan;
    case FM_COMPARISON_OPERATOR_GREATER_THAN_OR_EQUAL:
      return cf::CellIsOperator::GreaterThanOrEqual;
    default:
      return cf::CellIsOperator::Equal;
  }
}

inline std::int32_t to_c(cf::CellIsOperator value) {
  switch (value) {
    case cf::CellIsOperator::Between:
      return FM_COMPARISON_OPERATOR_BETWEEN;
    case cf::CellIsOperator::NotBetween:
      return FM_COMPARISON_OPERATOR_NOT_BETWEEN;
    case cf::CellIsOperator::NotEqual:
      return FM_COMPARISON_OPERATOR_NOT_EQUAL;
    case cf::CellIsOperator::LessThan:
      return FM_COMPARISON_OPERATOR_LESS_THAN;
    case cf::CellIsOperator::LessThanOrEqual:
      return FM_COMPARISON_OPERATOR_LESS_THAN_OR_EQUAL;
    case cf::CellIsOperator::GreaterThan:
      return FM_COMPARISON_OPERATOR_GREATER_THAN;
    case cf::CellIsOperator::GreaterThanOrEqual:
      return FM_COMPARISON_OPERATOR_GREATER_THAN_OR_EQUAL;
    case cf::CellIsOperator::Equal:
      break;
  }
  return FM_COMPARISON_OPERATOR_EQUAL;
}

// The data-validation model keeps the OOXML ordinals of `sheet_metadata.h`:
// between, notBetween, equal, notEqual, greaterThan, lessThan,
// greaterThanOrEqual, lessThanOrEqual.
inline std::uint8_t validation_operator_from_c(std::int32_t value) {
  switch (value) {
    case FM_COMPARISON_OPERATOR_LESS_THAN:
      return 5U;
    case FM_COMPARISON_OPERATOR_LESS_THAN_OR_EQUAL:
      return 7U;
    case FM_COMPARISON_OPERATOR_GREATER_THAN:
      return 4U;
    case FM_COMPARISON_OPERATOR_GREATER_THAN_OR_EQUAL:
      return 6U;
    default:
      return static_cast<std::uint8_t>(value);
  }
}

inline std::uint8_t validation_operator_to_c(std::uint8_t model) {
  switch (model) {
    case 4U:
      return FM_COMPARISON_OPERATOR_GREATER_THAN;
    case 5U:
      return FM_COMPARISON_OPERATOR_LESS_THAN;
    case 6U:
      return FM_COMPARISON_OPERATOR_GREATER_THAN_OR_EQUAL;
    case 7U:
      return FM_COMPARISON_OPERATOR_LESS_THAN_OR_EQUAL;
    default:
      return model;
  }
}

// AutoFilter conditions take only the six relational members; the caller
// rejects `BETWEEN` / `NOT_BETWEEN` first.
inline FilterOperator filter_operator_from_c(std::int32_t value) {
  switch (value) {
    case FM_COMPARISON_OPERATOR_NOT_EQUAL:
      return FilterOperator::kNotEqual;
    case FM_COMPARISON_OPERATOR_LESS_THAN:
      return FilterOperator::kLessThan;
    case FM_COMPARISON_OPERATOR_LESS_THAN_OR_EQUAL:
      return FilterOperator::kLessThanOrEqual;
    case FM_COMPARISON_OPERATOR_GREATER_THAN:
      return FilterOperator::kGreaterThan;
    case FM_COMPARISON_OPERATOR_GREATER_THAN_OR_EQUAL:
      return FilterOperator::kGreaterThanOrEqual;
    default:
      return FilterOperator::kEqual;
  }
}

inline std::int32_t to_c(FilterOperator value) {
  switch (value) {
    case FilterOperator::kNotEqual:
      return FM_COMPARISON_OPERATOR_NOT_EQUAL;
    case FilterOperator::kLessThan:
      return FM_COMPARISON_OPERATOR_LESS_THAN;
    case FilterOperator::kLessThanOrEqual:
      return FM_COMPARISON_OPERATOR_LESS_THAN_OR_EQUAL;
    case FilterOperator::kGreaterThan:
      return FM_COMPARISON_OPERATOR_GREATER_THAN;
    case FilterOperator::kGreaterThanOrEqual:
      return FM_COMPARISON_OPERATOR_GREATER_THAN_OR_EQUAL;
    case FilterOperator::kEqual:
      break;
  }
  return FM_COMPARISON_OPERATOR_EQUAL;
}

// --- date period --------------------------------------------------------------

inline cf::TimePeriod date_period_from_c(std::int32_t value) {
  switch (value) {
    case FM_DATE_PERIOD_YESTERDAY:
      return cf::TimePeriod::Yesterday;
    case FM_DATE_PERIOD_TOMORROW:
      return cf::TimePeriod::Tomorrow;
    case FM_DATE_PERIOD_LAST_7_DAYS:
      return cf::TimePeriod::Last7Days;
    case FM_DATE_PERIOD_THIS_MONTH:
      return cf::TimePeriod::ThisMonth;
    case FM_DATE_PERIOD_LAST_MONTH:
      return cf::TimePeriod::LastMonth;
    case FM_DATE_PERIOD_NEXT_MONTH:
      return cf::TimePeriod::NextMonth;
    case FM_DATE_PERIOD_THIS_WEEK:
      return cf::TimePeriod::ThisWeek;
    case FM_DATE_PERIOD_LAST_WEEK:
      return cf::TimePeriod::LastWeek;
    case FM_DATE_PERIOD_NEXT_WEEK:
      return cf::TimePeriod::NextWeek;
    default:
      return cf::TimePeriod::Today;
  }
}

inline std::int32_t to_c(cf::TimePeriod value) {
  switch (value) {
    case cf::TimePeriod::Yesterday:
      return FM_DATE_PERIOD_YESTERDAY;
    case cf::TimePeriod::Tomorrow:
      return FM_DATE_PERIOD_TOMORROW;
    case cf::TimePeriod::Last7Days:
      return FM_DATE_PERIOD_LAST_7_DAYS;
    case cf::TimePeriod::ThisMonth:
      return FM_DATE_PERIOD_THIS_MONTH;
    case cf::TimePeriod::LastMonth:
      return FM_DATE_PERIOD_LAST_MONTH;
    case cf::TimePeriod::NextMonth:
      return FM_DATE_PERIOD_NEXT_MONTH;
    case cf::TimePeriod::ThisWeek:
      return FM_DATE_PERIOD_THIS_WEEK;
    case cf::TimePeriod::LastWeek:
      return FM_DATE_PERIOD_LAST_WEEK;
    case cf::TimePeriod::NextWeek:
      return FM_DATE_PERIOD_NEXT_WEEK;
    case cf::TimePeriod::Today:
      break;
  }
  return FM_DATE_PERIOD_TODAY;
}

// --- closed sets the model keeps as OOXML tokens ------------------------------
//
// Each table is indexed by the C ordinal. `token_of` expects a value inside
// the table; `ordinal_of` returns `missing` for a spelling outside it.

template <std::size_t N>
constexpr std::string_view token_of(const std::array<std::string_view, N>& table, std::int32_t value) noexcept {
  return table[static_cast<std::size_t>(value)];
}

template <std::size_t N>
constexpr std::int32_t ordinal_of(const std::array<std::string_view, N>& table, std::string_view token,
                                  std::int32_t missing) noexcept {
  for (std::size_t i = 0; i < N; ++i) {
    if (table[i] == token) {
      return static_cast<std::int32_t>(i);
    }
  }
  return missing;
}

/// `fm_table_totals_row_function_t` -> `ST_TotalsRowFunction`.
inline constexpr std::array<std::string_view, 10> kTotalsRowFunctionTokens = {
    "none", "sum", "min", "max", "average", "count", "countNums", "stdDev", "var", "custom"};
static_assert(kTotalsRowFunctionTokens.size() == FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM + 1);

/// `fm_table_style_element_type_t` -> `ST_TableStyleType`.
inline constexpr std::array<std::string_view, 28> kTableStyleElementTypeTokens = {"wholeTable",
                                                                                  "headerRow",
                                                                                  "totalRow",
                                                                                  "firstColumn",
                                                                                  "lastColumn",
                                                                                  "firstRowStripe",
                                                                                  "secondRowStripe",
                                                                                  "firstColumnStripe",
                                                                                  "secondColumnStripe",
                                                                                  "firstHeaderCell",
                                                                                  "lastHeaderCell",
                                                                                  "firstTotalCell",
                                                                                  "lastTotalCell",
                                                                                  "firstSubtotalColumn",
                                                                                  "secondSubtotalColumn",
                                                                                  "thirdSubtotalColumn",
                                                                                  "firstSubtotalRow",
                                                                                  "secondSubtotalRow",
                                                                                  "thirdSubtotalRow",
                                                                                  "blankRow",
                                                                                  "firstColumnSubheading",
                                                                                  "secondColumnSubheading",
                                                                                  "thirdColumnSubheading",
                                                                                  "firstRowSubheading",
                                                                                  "secondRowSubheading",
                                                                                  "thirdRowSubheading",
                                                                                  "pageFieldLabels",
                                                                                  "pageFieldValues"};
static_assert(kTableStyleElementTypeTokens.size() == FM_TABLE_STYLE_ELEMENT_TYPE_PAGE_FIELD_VALUES + 1);

/// `fm_chart_tick_label_position_t` -> `ST_TickLblPos`.
inline constexpr std::array<std::string_view, 4> kTickLabelPositionTokens = {"nextTo", "high", "low", "none"};
static_assert(kTickLabelPositionTokens.size() == FM_CHART_TICK_LABEL_POSITION_NONE + 1);

/// `fm_width_platform_t` -> the width model's platform spelling.
inline constexpr std::array<std::string_view, 2> kWidthPlatformTokens = {"win", "mac"};
static_assert(kWidthPlatformTokens.size() == FM_WIDTH_PLATFORM_MAC + 1);

/// `fm_icon_set_name_t` -> `ST_IconSetType`: the 17 built-ins, then the
/// x14-only sets a custom icon level can name.
inline constexpr std::array<std::string_view, 21> kIconSetNameTokens = {
    "3Arrows",     "3ArrowsGray", "3Flags",      "3TrafficLights1", "3TrafficLights2", "3Signs",         "3Symbols",
    "3Symbols2",   "4Arrows",     "4ArrowsGray", "4RedToBlack",     "4Rating",         "4TrafficLights", "5Arrows",
    "5ArrowsGray", "5Rating",     "5Quarters",   "3Stars",          "3Triangles",      "5Boxes",         "NoIcons"};
static_assert(kIconSetNameTokens.size() == FM_ICON_SET_NAME_NO_ICONS + 1);
static_assert(
    [] {
      for (std::size_t i = 0; i < cf::kIconSetNames.size(); ++i) {
        if (kIconSetNameTokens[i] != cf::kIconSetNames[i]) {
          return false;
        }
      }
      return true;
    }(),
    "the built-in icon sets lead the C table in model order");

}  // namespace parts
}  // namespace c_api
}  // namespace formulon

#endif  // FORMULON_C_API_PARTS_ENUM_MAP_H_
