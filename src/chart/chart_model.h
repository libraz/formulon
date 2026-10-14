//
// Read-only typed view of a chart part, the resolved series data, and the
// option / patch types of the chart edit API.
//
// The chart part's DOM stays the source of truth: `ChartModel` is rebuilt from
// it on every request and never written back, so everything it does not model
// survives an edit. Every field that the file may omit keeps "absent"
// distinct from any value (`std::optional`, a tri-state, or NaN for doubles)
// and no schema default is guessed here; the effect of an absent element is
// the resolver's or the renderer's business.
//
// Colors are ARGB (`0xAARRGGBB`) resolved from explicit `a:srgbClr` / theme
// references; no explicit color reads as `std::nullopt`.
//
// Patches follow one rule: an unset field leaves the chart alone, `clear()`
// removes the element, `set(v)` writes it.

#ifndef FORMULON_CHART_CHART_MODEL_H_
#define FORMULON_CHART_CHART_MODEL_H_

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "drawing/drawing_edit.h"
#include "drawing/drawing_xml.h"

namespace formulon {

/// `host_sheet` value for a chart that lives on a chartsheet (no host cell grid).
inline constexpr std::size_t kChartNoHostSheet = static_cast<std::size_t>(-1);

/// Tri-state boolean for elements whose absence differs from either value.
enum class Tri : std::uint8_t { kAbsent = 0, kFalse = 1, kTrue = 2 };

/// A rectangle of cells on one sheet (0-based, inclusive).
struct SheetRect {
  std::size_t sheet = 0;
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t bottom = 0;
  std::uint32_t right = 0;
};

// ---------------------------------------------------------------------------
// Enumerations (numeric values are part of the binding surface)
// ---------------------------------------------------------------------------

enum class ChartPlotType : std::int32_t {
  kBar = 0,
  kBar3D = 1,
  kLine = 2,
  kLine3D = 3,
  kArea = 4,
  kArea3D = 5,
  kPie = 6,
  kPie3D = 7,
  kDoughnut = 8,
  kOfPie = 9,
  kScatter = 10,
  kBubble = 11,
  kRadar = 12,
  kStock = 13,
  kSurface = 14,
  kSurface3D = 15,
  kChartex = 16,
};

/// The six kinds `insert_chart` / `set_chart_type` can produce.
enum class ChartInsertKind : std::int32_t {
  kColumn = 0,
  kBar = 1,
  kLine = 2,
  kArea = 3,
  kPie = 4,
  kScatter = 5,
  kDoughnut = 6,
};

enum class ChartGrouping : std::int32_t { kClustered = 0, kStacked = 1, kPercentStacked = 2, kStandard = 3 };

enum class ChartBlanksAs : std::int32_t { kAbsent = 0, kGap = 1, kZero = 2, kSpan = 3 };

enum class ChartPointState : std::uint8_t { kValue = 0, kGap = 1, kConnect = 2 };

enum class ChartAxisKind : std::int32_t { kCategory = 0, kValue = 1, kDate = 2, kSeries = 3 };

enum class ChartTitleKind : std::int32_t { kNone = 0, kText = 1, kRef = 2, kAuto = 3 };

enum class ChartSeriesDirection : std::int32_t { kAuto = 0, kColumns = 1, kRows = 2 };

enum class ChartLegendPosition : std::int32_t {
  kNone = 0,
  kBottom = 1,
  kTop = 2,
  kLeft = 3,
  kRight = 4,
  kTopRight = 5
};

enum class ChartLineDash : std::int32_t { kSolid = 0, kDash = 1, kDot = 2, kDashDot = 3, kOther = 4 };

enum class ChartMarkerSymbol : std::int32_t {
  kNone = 0,
  kAuto = 1,
  kCircle = 2,
  kSquare = 3,
  kDiamond = 4,
  kTriangle = 5,
  kX = 6,
  kStar = 7,
  kPlus = 8,
  kDot = 9,
  kDash = 10,
  kPicture = 11,  ///< Read-only: writing one needs a blip.
};

enum class ChartLabelPosition : std::int32_t {
  kBestFit = 0,
  kCenter = 1,
  kInsideBase = 2,
  kInsideEnd = 3,
  kOutsideEnd = 4,
  kAbove = 5,
  kBelow = 6,
  kLeft = 7,
  kRight = 8,
};

/// Bar orientation (`c:barDir`).
enum class ChartBarDirection : std::uint8_t { kAbsent = 0, kCol = 1, kBar = 2 };

enum class ChartAxisPosition : std::uint8_t { kBottom = 0, kTop = 1, kLeft = 2, kRight = 3 };

enum class ChartAxisCrosses : std::uint8_t { kAbsent = 0, kAutoZero = 1, kMin = 2, kMax = 3, kAt = 4 };

/// Forced use of a heading row / column when splitting a source range.
enum class ChartHeaderMode : std::uint8_t { kAuto = 0, kYes = 1, kNo = 2 };

// ---------------------------------------------------------------------------
// ChartModel
// ---------------------------------------------------------------------------

/// `c:spPr` solid fill. `color` unset means no explicit fill color.
struct ChartFill {
  std::optional<std::uint32_t> color;
  std::optional<std::uint32_t> alpha;  ///< 0..100000 (`a:alpha/@val`).
};

/// `a:ln` of a series, point or marker.
struct ChartLine {
  std::optional<std::uint32_t> color;
  std::optional<std::uint32_t> alpha;  ///< 0..100000.
  std::optional<std::int64_t> width_emu;
  std::optional<ChartLineDash> dash;
  std::string dash_name;  ///< `a:prstDash/@val` verbatim; empty when absent.
  bool none = false;      ///< `a:ln` holds `a:noFill`.
};

struct ChartMarker {
  ChartMarkerSymbol symbol = ChartMarkerSymbol::kAuto;
  std::string symbol_name;  ///< `c:symbol/@val` verbatim; empty when absent.
  std::optional<std::uint32_t> size;
  ChartFill fill;
  ChartLine line;
};

/// One `c:dPt`.
struct ChartPoint {
  std::uint32_t idx = 0;
  ChartFill fill;
  ChartLine line;
  std::optional<ChartMarker> marker;
};

struct ChartNumFmt {
  std::string code;
  bool source_linked = false;
};

/// `c:dLbls` of a series.
struct ChartLabels {
  bool show_val = false;
  bool show_cat = false;
  bool show_ser = false;
  bool show_pct = false;
  bool show_legend_key = false;
  std::optional<std::string> separator;
  std::optional<ChartLabelPosition> position;
  std::optional<ChartNumFmt> num_fmt;
  std::optional<std::uint32_t> font_color;
  std::optional<std::uint32_t> font_size_hundredths_pt;
  bool deleted = false;  ///< `<c:delete val="1"/>`.
};

/// One cached or literal `c:pt` of a data source. Points absent from the file are not stored.
struct ChartValue {
  enum class Kind : std::uint8_t { kBlank = 0, kNumber = 1, kText = 2 };
  std::uint32_t idx = 0;  ///< `c:pt/@idx`.
  Kind kind = Kind::kBlank;
  double number = 0.0;
  std::string text;
  std::string format_code;  ///< `c:pt/@formatCode`; empty when absent.
};

/// A series' name / category / value / x / y / size input.
struct ChartDataSource {
  enum class Kind : std::uint8_t { kNone = 0, kRef = 1, kLiteral = 2 };
  enum class RefElement : std::uint8_t { kNone = 0, kNumRef = 1, kStrRef = 2, kMultiLvlStrRef = 3, kCxDim = 4 };
  Kind kind = Kind::kNone;
  RefElement ref_element = RefElement::kNone;
  std::string formula;       ///< `c:f` / `cx:f` text.
  std::string full_formula;  ///< `c15:fullRef/c15:sqref` of a filtered ref; empty when absent.
  /// `c:ptCount` of the literal or cache. Whole-column refs carry 1048576, so points stay sparse.
  std::uint32_t point_count = 0;
  std::vector<ChartValue> literal;  ///< `numLit` / `strLit` / `c:v` points, ascending idx.
  std::vector<ChartValue> cache;    ///< `numCache` / `strCache` points of a ref, ascending idx.
  /// Upper levels of a `multiLvlStrRef` cache, outermost first; `cache` holds the lowest level.
  std::vector<std::vector<ChartValue>> cache_levels;
  std::string cache_format_code;  ///< `c:formatCode` of the cache; empty when absent.
};

struct ChartSeries {
  std::uint32_t idx = 0;
  std::uint32_t order = 0;
  bool filtered = false;  ///< Held in `c15:filtered*Series`.
  ChartDataSource name;
  ChartDataSource cat;
  ChartDataSource val;
  ChartDataSource x;
  ChartDataSource y;
  ChartDataSource size;
  ChartFill fill;
  ChartLine line;
  std::optional<ChartMarker> marker;
  Tri smooth = Tri::kAbsent;
  std::vector<ChartPoint> points;
  std::optional<ChartLabels> labels;  ///< Unset when the series has no `c:dLbls`.
};

struct ChartPlot {
  ChartPlotType type = ChartPlotType::kBar;
  std::string chartex_layout;  ///< `layoutId` of a chartex plot.
  ChartBarDirection bar_dir = ChartBarDirection::kAbsent;
  std::optional<ChartGrouping> grouping;
  Tri vary_colors = Tri::kAbsent;
  std::optional<std::uint32_t> gap_width;
  std::optional<std::int32_t> overlap;
  std::optional<std::uint32_t> hole_size;
  std::optional<std::uint32_t> first_slice_angle;
  std::string scatter_style;
  std::string radar_style;
  std::vector<std::uint32_t> axis_ids;
  std::vector<ChartSeries> series;
  std::optional<ChartLabels> labels;  ///< Plot-group `c:dLbls`; unset when absent.
};

struct ChartTitle {
  ChartTitleKind kind = ChartTitleKind::kNone;
  std::string text;         ///< Rich text flattened to plain text (kind == kText).
  std::string ref_formula;  ///< `c:strRef/c:f` (kind == kRef).
};

struct ChartLegend {
  bool present = false;
  ChartLegendPosition position = ChartLegendPosition::kRight;
  bool overlay = false;
  std::vector<std::uint32_t> hidden_series_idx;  ///< `c:legendEntry` with `c:delete`.
};

struct ChartAxis {
  std::uint32_t id = 0;
  ChartAxisKind kind = ChartAxisKind::kCategory;
  ChartAxisPosition position = ChartAxisPosition::kBottom;
  bool deleted = false;
  std::uint32_t cross_ax = 0;
  ChartAxisCrosses crosses = ChartAxisCrosses::kAbsent;
  double crosses_at = std::numeric_limits<double>::quiet_NaN();
  bool reversed = false;  ///< `c:orientation val="maxMin"`.
  double log_base = std::numeric_limits<double>::quiet_NaN();
  double min = std::numeric_limits<double>::quiet_NaN();
  double max = std::numeric_limits<double>::quiet_NaN();
  double major_unit = std::numeric_limits<double>::quiet_NaN();
  double minor_unit = std::numeric_limits<double>::quiet_NaN();
  std::optional<ChartNumFmt> num_fmt;
  ChartTitle title;
  bool major_gridlines = false;
  bool minor_gridlines = false;
  std::string tick_label_pos;  ///< `c:tickLblPos/@val`; empty when absent.
  std::string major_tick;      ///< `c:majorTickMark/@val`; empty when absent.
  std::string minor_tick;
  std::string base_time_unit;  ///< Date axes only.
};

/// Derived "chart data range" (not stored in the file); see `chart_template`.
struct ChartSource {
  std::string range;  ///< Sheet-qualified A1 rectangle shared by all series.
  ChartSeriesDirection direction = ChartSeriesDirection::kColumns;
  bool header_row = false;
  bool header_column = false;
};

struct ChartModel {
  std::uint32_t object_id = 0;
  std::string name;
  std::string descr;
  bool is_chartex = false;
  bool is_pivot_chart = false;
  bool on_chartsheet = false;

  /// Anchor of the owning `graphicFrame`, in `DrawingObject` terms.
  AnchorKind anchor_kind = AnchorKind::kTwoCell;
  EditAs edit_as = EditAs::kTwoCell;
  AnchorPoint from;
  AnchorPoint to;
  std::int64_t cx = 0;
  std::int64_t cy = 0;

  ChartTitle title;
  Tri auto_title_deleted = Tri::kAbsent;
  ChartLegend legend;
  ChartBlanksAs disp_blanks_as = ChartBlanksAs::kAbsent;
  Tri plot_vis_only = Tri::kAbsent;
  Tri disp_na_as_blank = Tri::kAbsent;  ///< `c16r3:dataDisplayOptions16/c16r3:dispNaAsBlank`.
  bool date1904 = false;
  std::optional<std::int32_t> style;  ///< `c14:style` first, else `c:style`.

  std::vector<ChartPlot> plots;
  std::vector<std::uint32_t> category_filter;  ///< Hidden category indices; empty when absent.
  std::vector<ChartAxis> axes;
  std::optional<ChartSource> source;  ///< Unset when the series do not share one rectangle.
};

// ---------------------------------------------------------------------------
// ChartData (resolved series data)
// ---------------------------------------------------------------------------

struct ChartDataPoint {
  double value = 0.0;
  ChartPointState state = ChartPointState::kValue;
};

struct ChartXYPoint {
  double x = 0.0;
  double y = 0.0;
  ChartPointState state = ChartPointState::kValue;
};

struct ChartDataSeries {
  std::size_t plot_index = 0;
  std::uint32_t idx = 0;
  std::optional<std::string> name;  ///< Unset when the series has no name reference.
  std::string format_code;          ///< Category domain: values; XY domain: y values.
  std::string x_format_code;        ///< XY domain: x values; empty otherwise.
  bool filtered = false;
  std::vector<ChartDataPoint> points;  ///< Category domain.
  std::vector<ChartXYPoint> xy;        ///< XY domain.
  std::vector<double> sizes;           ///< Bubble only.
  std::vector<SheetRect> name_ranges;
  std::vector<SheetRect> cat_ranges;
  std::vector<SheetRect> val_ranges;
  std::vector<SheetRect> x_ranges;
  std::vector<SheetRect> y_ranges;
  std::vector<SheetRect> size_ranges;
};

struct ChartDataCategories {
  std::vector<std::string> labels;
  std::vector<std::vector<std::string>> levels;  ///< Upper levels of a multi-level axis.
  std::vector<double> values;                    ///< Numeric categories (date axes); empty for text.
  std::string format_code;
  bool generated = false;  ///< No category reference: 1..n generated.
};

/// A data source that could not be resolved (its points all read as gaps).
struct ChartUnresolved {
  std::uint32_t series_idx = 0;  ///< Series `idx`; 0 for `title` / `axis_title`.
  std::string source;            ///< `name` / `cat` / `val` / `x` / `y` / `size` / `title` / `axis_title`.
  std::uint32_t axis_id = 0;     ///< Axis id when `source` is `axis_title`.
};

struct ChartAxisTitleText {
  std::uint32_t axis_id = 0;
  std::string text;
};

struct ChartData {
  enum class Domain : std::uint8_t { kCategory = 0, kXY = 1 };
  Domain domain = Domain::kCategory;
  std::size_t point_count = 0;
  ChartDataCategories categories;
  std::vector<ChartDataSeries> series;
  std::optional<std::string> title_text;
  std::vector<ChartAxisTitleText> axis_title_texts;
  std::vector<ChartUnresolved> unresolved;
};

// ---------------------------------------------------------------------------
// Insert options
// ---------------------------------------------------------------------------

struct ChartInsertOptions {
  ChartInsertKind kind = ChartInsertKind::kColumn;
  ChartGrouping grouping = ChartGrouping::kClustered;  ///< Column / bar / line / area only.
  std::string source;                                  ///< Single rectangle; unqualified means the target sheet.
  ChartSeriesDirection direction = ChartSeriesDirection::kAuto;
  ChartHeaderMode header_row = ChartHeaderMode::kAuto;
  ChartHeaderMode header_column = ChartHeaderMode::kAuto;
  /// Omitted = the Excel default (auto title); empty = no title.
  std::optional<std::string> title;
  ChartLegendPosition legend = ChartLegendPosition::kBottom;
  /// Placement; `name` / `descr` here are ignored in favour of the fields below.
  /// Defaults to a two-cell anchor, unlike `insert_image`.
  ImageInsertOptions placement = DefaultPlacement();
  std::string name;
  std::string descr;

  static ImageInsertOptions DefaultPlacement() {
    ImageInsertOptions p;
    p.anchor_kind = AnchorKind::kTwoCell;
    return p;
  }
};

// ---------------------------------------------------------------------------
// Patches
// ---------------------------------------------------------------------------

/// A three-way patch field: unset (leave alone), cleared (remove), or set.
template <typename T>
struct ChartPatchField {
  enum class Op : std::uint8_t { kKeep = 0, kClear = 1, kSet = 2 };
  Op op = Op::kKeep;
  T value{};

  static ChartPatchField Set(T v) {
    ChartPatchField f;
    f.op = Op::kSet;
    f.value = std::move(v);
    return f;
  }
  static ChartPatchField Clear() {
    ChartPatchField f;
    f.op = Op::kClear;
    return f;
  }
  bool is_keep() const { return op == Op::kKeep; }
  bool is_clear() const { return op == Op::kClear; }
  bool is_set() const { return op == Op::kSet; }
};

struct ChartFillPatch {
  ChartPatchField<std::uint32_t> color;
  ChartPatchField<double> transparency_pct;  ///< 0..100.
};

struct ChartLinePatch {
  ChartPatchField<std::uint32_t> color;
  ChartPatchField<double> transparency_pct;
  ChartPatchField<double> width_pt;
  ChartPatchField<ChartLineDash> dash;
  std::optional<bool> none;
};

struct ChartMarkerPatch {
  std::optional<ChartMarkerSymbol> symbol;
  ChartPatchField<double> size;
  ChartFillPatch fill;
  ChartLinePatch line;
};

struct ChartLabelsPatch {
  std::optional<bool> show_val;
  std::optional<bool> show_cat;
  std::optional<bool> show_ser;
  std::optional<bool> show_pct;
  std::optional<bool> show_legend_key;
  ChartPatchField<std::string> separator;
  ChartPatchField<ChartLabelPosition> position;
  ChartPatchField<std::string> number_format;
  ChartPatchField<std::uint32_t> font_color;
  ChartPatchField<double> font_size_pt;
};

/// Reference strings: unset keeps, empty removes (name / cat / x), and an
/// empty val / y / size is `kInvalidArgument`.
struct ChartSeriesPatch {
  std::optional<std::string> name;
  std::optional<std::string> cat;
  std::optional<std::string> val;
  std::optional<std::string> x;
  std::optional<std::string> y;
  std::optional<std::string> size;
  ChartFillPatch fill;
  ChartLinePatch line;
  ChartPatchField<ChartMarkerPatch> marker;  ///< Clear removes `c:marker`.
  std::optional<bool> smooth;
  ChartPatchField<ChartLabelsPatch> labels;  ///< Clear removes the series' data labels.
};

/// An all-keep patch removes the `c:dPt`.
struct ChartPointPatch {
  ChartFillPatch fill;
  ChartLinePatch line;
  ChartPatchField<ChartMarkerPatch> marker;
};

struct ChartPlotPatch {
  std::optional<std::uint32_t> gap_width;
  std::optional<std::int32_t> overlap;
  std::optional<std::uint32_t> first_slice_angle;
  std::optional<bool> vary_colors;
  std::optional<std::uint32_t> hole_size;  ///< Doughnut `c:holeSize`, 10..90.
};

struct ChartAxisPatch {
  std::optional<bool> visible;
  ChartPatchField<ChartTitle> title;  ///< Clear, or a kNone title, removes the axis title.
  ChartPatchField<double> min;
  ChartPatchField<double> max;
  ChartPatchField<double> major_unit;
  ChartPatchField<std::string> number_format;
  std::optional<bool> major_gridlines;
  std::optional<bool> reversed;             ///< `c:scaling/c:orientation` maxMin when true.
  ChartPatchField<double> log_base;         ///< Clear makes the axis linear.
  std::optional<ChartAxisCrosses> crosses;  ///< Crossing point on the cross axis; kAt uses `crosses_at`.
  double crosses_at = 0.0;
  ChartPatchField<std::string> tick_label_pos;  ///< `nextTo` / `low` / `high` / `none`; Clear restores the default.
};

struct ChartLegendPatch {
  ChartLegendPosition position = ChartLegendPosition::kRight;  ///< kNone removes the legend.
  bool overlay = false;
  /// Unset leaves the legend entries alone.
  std::optional<std::vector<std::uint32_t>> hidden_series_idx;
};

/// Not-yet-committed changes `preview_chart_data` applies to a copy of the chart.
struct ChartPreviewOverrides {
  std::optional<std::string> source;
  std::optional<ChartSeriesDirection> direction;
  std::optional<bool> header_row;
  std::optional<bool> header_column;
  std::optional<ChartBlanksAs> blanks_as;
  std::optional<bool> plot_vis_only;
};

}  // namespace formulon

#endif  // FORMULON_CHART_CHART_MODEL_H_
