#include "chart/chart_xml.h"

#include <algorithm>
#include <initializer_list>
#include <limits>
#include <optional>
#include <set>
#include <string>
#include <string_view>
#include <utility>

#include "color_resolve.h"
#include "drawing/drawing_xml.h"
#include "io/xml_utils.h"
#include "utils/a1_ref.h"
#include "utils/double_parse.h"
#include "workbook.h"

namespace formulon {
namespace {

using pugi::xml_node;

constexpr double kNaN = std::numeric_limits<double>::quiet_NaN();

Error unparseable(std::string_view what, std::string_view detail) {
  return make_error(FormulonErrorCode::kIoChartUnparseable, std::string("chart: ") + std::string(what),
                    std::string(detail));
}

xml_node kid(const xml_node& parent, std::string_view name) {
  return child_local(parent, name);
}

xml_node path(xml_node node, std::initializer_list<std::string_view> names) {
  for (const std::string_view name : names) {
    node = kid(node, name);
  }
  return node;
}

std::vector<xml_node> kids(const xml_node& parent, std::string_view name) {
  std::vector<xml_node> out;
  for (xml_node c = parent.first_child(); c; c = c.next_sibling()) {
    if (c.type() == pugi::node_element && local_name(c) == name) {
      out.push_back(c);
    }
  }
  return out;
}

std::string_view val_of(const xml_node& node) {
  return node.attribute("val").value();
}

/// `CT_Boolean` reading: no element is absent, a bare element is true.
Tri tri_of(const xml_node& node) {
  if (!node) {
    return Tri::kAbsent;
  }
  const std::string_view v = val_of(node);
  return (v == "0" || v == "false") ? Tri::kFalse : Tri::kTrue;
}

bool flag_of(const xml_node& node) {
  return tri_of(node) == Tri::kTrue;
}

bool attr_flag(const xml_node& node, const char* name) {
  return node.attribute(name).as_bool();
}

std::optional<std::uint32_t> uint_of(const xml_node& node) {
  if (!node) {
    return std::nullopt;
  }
  return static_cast<std::uint32_t>(node.attribute("val").as_uint());
}

std::optional<std::int32_t> int_of(const xml_node& node) {
  if (!node) {
    return std::nullopt;
  }
  return static_cast<std::int32_t>(node.attribute("val").as_int());
}

double double_text(std::string_view text) {
  double out = 0.0;
  return parse_double_exact(text, &out) ? out : kNaN;
}

double double_of(const xml_node& node) {
  return node ? double_text(val_of(node)) : kNaN;
}

// ---------------------------------------------------------------------------
// Colours and line / fill / marker formatting
// ---------------------------------------------------------------------------

std::optional<std::uint32_t> scheme_index(std::string_view name) {
  if (name == "bg1" || name == "lt1")
    return 0U;
  if (name == "tx1" || name == "dk1")
    return 1U;
  if (name == "bg2" || name == "lt2")
    return 2U;
  if (name == "tx2" || name == "dk2")
    return 3U;
  if (name == "hlink")
    return 10U;
  if (name == "folHlink")
    return 11U;
  if (name.size() == 7 && name.substr(0, 6) == "accent" && name[6] >= '1' && name[6] <= '6') {
    return 4U + static_cast<std::uint32_t>(name[6] - '1');
  }
  return std::nullopt;
}

/// Applies `a:lumMod` / `a:lumOff` as the equivalent spreadsheet tint; exact
/// for the forms Excel writes (`lumMod + lumOff == 1`, or `lumOff` absent).
std::uint32_t apply_luminance(std::uint32_t argb, const xml_node& clr) {
  const xml_node mod = kid(clr, "lumMod");
  const xml_node off = kid(clr, "lumOff");
  if (!mod && !off) {
    return argb;
  }
  const double m = mod ? mod.attribute("val").as_double() / 100000.0 : 1.0;
  const double o = off ? off.attribute("val").as_double() / 100000.0 : 0.0;
  const double tint = o > 0.0 ? o : m - 1.0;
  return apply_tint(argb, std::clamp(tint, -1.0, 1.0));
}

struct Paint {
  std::optional<std::uint32_t> color;
  std::optional<std::uint32_t> alpha;
};

/// Colour of an `a:solidFill`; unresolvable colour kinds read as no colour.
Paint read_solid(const Workbook& wb, const xml_node& solid_fill) {
  Paint paint;
  const xml_node clr = kid(solid_fill, "");
  if (!clr) {
    return paint;
  }
  const std::string_view kind = local_name(clr);
  if (kind == "srgbClr") {
    paint.color = apply_luminance(io::parse_rgb_hex(val_of(clr), 0U) | 0xFF000000U, clr);
  } else if (kind == "schemeClr") {
    if (const auto index = scheme_index(val_of(clr))) {
      ColorSpec spec;
      spec.kind = ColorSpec::Kind::kTheme;
      spec.theme = *index;
      paint.color = apply_luminance(resolve_color(wb, spec, ColorContext::kFont).argb, clr);
    }
  } else if (kind == "sysClr") {
    const std::string_view last = clr.attribute("lastClr").value();
    if (!last.empty()) {
      paint.color = apply_luminance(io::parse_rgb_hex(last, 0U) | 0xFF000000U, clr);
    } else {
      paint.color = val_of(clr) == "window" ? 0xFFFFFFFFU : 0xFF000000U;
    }
  }
  if (paint.color.has_value()) {
    if (const xml_node alpha = kid(clr, "alpha")) {
      paint.alpha = static_cast<std::uint32_t>(alpha.attribute("val").as_uint());
    }
  }
  return paint;
}

ChartFill read_fill(const Workbook& wb, const xml_node& sp_pr) {
  const Paint paint = read_solid(wb, kid(sp_pr, "solidFill"));
  ChartFill fill;
  fill.color = paint.color;
  fill.alpha = paint.alpha;
  return fill;
}

ChartLineDash dash_of(std::string_view name) {
  if (name == "solid")
    return ChartLineDash::kSolid;
  if (name == "dash")
    return ChartLineDash::kDash;
  if (name == "dot")
    return ChartLineDash::kDot;
  if (name == "dashDot")
    return ChartLineDash::kDashDot;
  return ChartLineDash::kOther;
}

ChartLine read_line(const Workbook& wb, const xml_node& sp_pr) {
  ChartLine line;
  const xml_node ln = kid(sp_pr, "ln");
  if (!ln) {
    return line;
  }
  if (const pugi::xml_attribute w = ln.attribute("w")) {
    line.width_emu = static_cast<std::int64_t>(w.as_llong());
  }
  line.none = static_cast<bool>(kid(ln, "noFill"));
  const Paint paint = read_solid(wb, kid(ln, "solidFill"));
  line.color = paint.color;
  line.alpha = paint.alpha;
  if (const xml_node dash = kid(ln, "prstDash")) {
    line.dash_name = val_of(dash);
    line.dash = dash_of(line.dash_name);
  }
  return line;
}

ChartMarkerSymbol symbol_of(std::string_view name) {
  if (name == "none")
    return ChartMarkerSymbol::kNone;
  if (name == "auto")
    return ChartMarkerSymbol::kAuto;
  if (name == "circle")
    return ChartMarkerSymbol::kCircle;
  if (name == "square")
    return ChartMarkerSymbol::kSquare;
  if (name == "diamond")
    return ChartMarkerSymbol::kDiamond;
  if (name == "triangle")
    return ChartMarkerSymbol::kTriangle;
  if (name == "x")
    return ChartMarkerSymbol::kX;
  if (name == "star")
    return ChartMarkerSymbol::kStar;
  if (name == "plus")
    return ChartMarkerSymbol::kPlus;
  if (name == "dot")
    return ChartMarkerSymbol::kDot;
  if (name == "dash")
    return ChartMarkerSymbol::kDash;
  if (name == "picture")
    return ChartMarkerSymbol::kPicture;
  return ChartMarkerSymbol::kAuto;  // Not an ST_MarkerStyle value.
}

ChartMarker read_marker(const Workbook& wb, const xml_node& marker) {
  ChartMarker out;
  if (const xml_node symbol = kid(marker, "symbol")) {
    out.symbol_name = val_of(symbol);
    out.symbol = symbol_of(out.symbol_name);
  }
  out.size = uint_of(kid(marker, "size"));
  const xml_node sp_pr = kid(marker, "spPr");
  out.fill = read_fill(wb, sp_pr);
  out.line = read_line(wb, sp_pr);
  return out;
}

std::optional<ChartLabelPosition> label_position_of(std::string_view name) {
  if (name == "bestFit")
    return ChartLabelPosition::kBestFit;
  if (name == "ctr")
    return ChartLabelPosition::kCenter;
  if (name == "inBase")
    return ChartLabelPosition::kInsideBase;
  if (name == "inEnd")
    return ChartLabelPosition::kInsideEnd;
  if (name == "outEnd")
    return ChartLabelPosition::kOutsideEnd;
  if (name == "t")
    return ChartLabelPosition::kAbove;
  if (name == "b")
    return ChartLabelPosition::kBelow;
  if (name == "l")
    return ChartLabelPosition::kLeft;
  if (name == "r")
    return ChartLabelPosition::kRight;
  return std::nullopt;
}

ChartNumFmt read_num_fmt(const xml_node& node) {
  ChartNumFmt fmt;
  fmt.code = node.attribute("formatCode").value();
  fmt.source_linked = attr_flag(node, "sourceLinked");
  return fmt;
}

/// `c:dLbls` of a series. Per-point `c:dLbl` children are not modelled.
ChartLabels read_labels(const Workbook& wb, const xml_node& dlbls) {
  ChartLabels labels;
  labels.deleted = flag_of(kid(dlbls, "delete"));
  labels.show_val = flag_of(kid(dlbls, "showVal"));
  labels.show_cat = flag_of(kid(dlbls, "showCatName"));
  labels.show_ser = flag_of(kid(dlbls, "showSerName"));
  labels.show_pct = flag_of(kid(dlbls, "showPercent"));
  labels.show_legend_key = flag_of(kid(dlbls, "showLegendKey"));
  if (const xml_node sep = kid(dlbls, "separator")) {
    labels.separator = std::string(sep.child_value());
  }
  labels.position = label_position_of(val_of(kid(dlbls, "dLblPos")));
  if (const xml_node fmt = kid(dlbls, "numFmt")) {
    labels.num_fmt = read_num_fmt(fmt);
  }
  const xml_node def_rpr = path(dlbls, {"txPr", "p", "pPr", "defRPr"});
  if (def_rpr) {
    if (const pugi::xml_attribute sz = def_rpr.attribute("sz")) {
      labels.font_size_hundredths_pt = static_cast<std::uint32_t>(sz.as_uint());
    }
    labels.font_color = read_solid(wb, kid(def_rpr, "solidFill")).color;
  }
  return labels;
}

// ---------------------------------------------------------------------------
// Data sources
// ---------------------------------------------------------------------------

ChartValue make_value(std::string_view text, bool numeric) {
  ChartValue value;
  double number = 0.0;
  if (numeric && parse_double_exact(text, &number)) {
    value.kind = ChartValue::Kind::kNumber;
    value.number = number;
  } else {
    value.kind = ChartValue::Kind::kText;
    value.text = std::string(text);
  }
  return value;
}

/// Reads `ptCount` plus `pt` children; a point without a `pt` stays blank.
/// `direct_text` selects the chartex shape (`cx:pt` holds the text itself).
void read_points(const xml_node& cache, bool numeric, bool direct_text, std::vector<ChartValue>& out,
                 std::uint32_t& point_count) {
  if (const xml_node count_node = kid(cache, "ptCount")) {
    point_count = count_node.attribute("val").as_uint();
  } else {
    point_count = cache.attribute("ptCount").as_uint();
  }
  out.clear();
  for (const xml_node& pt : kids(cache, "pt")) {
    ChartValue value = make_value(direct_text ? pt.child_value() : kid(pt, "v").child_value(), numeric);
    value.idx = pt.attribute("idx").as_uint();
    value.format_code = pt.attribute("formatCode").value();
    out.push_back(std::move(value));
  }
  std::stable_sort(out.begin(), out.end(), [](const ChartValue& l, const ChartValue& r) { return l.idx < r.idx; });
}

xml_node ext_child(const xml_node& ref, std::string_view name) {
  for (const xml_node& ext : kids(kid(ref, "extLst"), "ext")) {
    if (const xml_node found = kid(ext, name)) {
      return found;
    }
  }
  return {};
}

/// A series' name / category / value input: `c:tx`, `c:cat`, `c:val`, ...
ChartDataSource read_data_source(const xml_node& holder) {
  ChartDataSource src;
  const xml_node body = kid(holder, "");
  if (!body) {
    return src;
  }
  const std::string_view kind = local_name(body);
  if (kind == "strRef" || kind == "numRef" || kind == "multiLvlStrRef" || kind == "multiLvlNumRef") {
    src.kind = ChartDataSource::Kind::kRef;
    src.ref_element = kind == "strRef"   ? ChartDataSource::RefElement::kStrRef
                      : kind == "numRef" ? ChartDataSource::RefElement::kNumRef
                                         : ChartDataSource::RefElement::kMultiLvlStrRef;
    if (const xml_node f = kid(body, "f")) {
      src.formula = f.child_value();
    } else {
      src.formula = kid(ext_child(body, "formulaRef"), "sqref").child_value();
    }
    src.full_formula = kid(ext_child(body, "fullRef"), "sqref").child_value();
    if (kind == "multiLvlStrRef" || kind == "multiLvlNumRef") {
      const xml_node cache = kid(body, kind == "multiLvlStrRef" ? "multiLvlStrCache" : "multiLvlNumCache");
      std::uint32_t unused = 0;
      src.point_count = kid(cache, "ptCount").attribute("val").as_uint();
      const std::vector<xml_node> levels = kids(cache, "lvl");
      for (std::size_t i = 0; i < levels.size(); ++i) {
        std::vector<ChartValue> level;
        read_points(levels[i], false, false, level, unused);
        if (i == 0) {
          src.cache = std::move(level);
        } else {
          src.cache_levels.insert(src.cache_levels.begin(), std::move(level));
        }
      }
    } else {
      const xml_node cache = kid(body, kind == "strRef" ? "strCache" : "numCache");
      if (cache) {
        read_points(cache, kind == "numRef", false, src.cache, src.point_count);
        src.cache_format_code = kid(cache, "formatCode").child_value();
      }
    }
  } else if (kind == "numLit" || kind == "strLit") {
    src.kind = ChartDataSource::Kind::kLiteral;
    read_points(body, kind == "numLit", false, src.literal, src.point_count);
    src.cache_format_code = kid(body, "formatCode").child_value();
  } else if (kind == "v") {
    src.kind = ChartDataSource::Kind::kLiteral;
    src.literal.push_back(make_value(body.child_value(), false));
    src.point_count = 1;
  }
  return src;
}

// ---------------------------------------------------------------------------
// Titles, legend, axes
// ---------------------------------------------------------------------------

std::string flatten_rich(const xml_node& rich) {
  std::string out;
  bool first = true;
  for (const xml_node& p : kids(rich, "p")) {
    if (!first) {
      out += '\n';
    }
    first = false;
    for (xml_node run = p.first_child(); run; run = run.next_sibling()) {
      if (run.type() != pugi::node_element) {
        continue;
      }
      const std::string_view name = local_name(run);
      if (name == "r" || name == "fld") {
        out += kid(run, "t").child_value();
      } else if (name == "br") {
        out += '\n';
      }
    }
  }
  return out;
}

ChartTitle read_title(const xml_node& title) {
  ChartTitle out;
  if (!title) {
    return out;
  }
  out.kind = ChartTitleKind::kAuto;
  const xml_node tx = kid(title, "tx");
  if (const xml_node rich = kid(tx, "rich")) {
    out.kind = ChartTitleKind::kText;
    out.text = flatten_rich(rich);
  } else if (const xml_node str_ref = kid(tx, "strRef")) {
    out.kind = ChartTitleKind::kRef;
    out.ref_formula = kid(str_ref, "f").child_value();
  }
  return out;
}

ChartLegendPosition legend_position_of(std::string_view name, ChartLegendPosition fallback) {
  if (name == "b")
    return ChartLegendPosition::kBottom;
  if (name == "t")
    return ChartLegendPosition::kTop;
  if (name == "l")
    return ChartLegendPosition::kLeft;
  if (name == "r")
    return ChartLegendPosition::kRight;
  if (name == "tr")
    return ChartLegendPosition::kTopRight;
  return fallback;
}

ChartLegend read_legend(const xml_node& legend) {
  ChartLegend out;
  if (!legend) {
    return out;
  }
  out.present = true;
  out.position = legend_position_of(val_of(kid(legend, "legendPos")), ChartLegendPosition::kRight);
  out.overlay = flag_of(kid(legend, "overlay"));
  for (const xml_node& entry : kids(legend, "legendEntry")) {
    if (flag_of(kid(entry, "delete"))) {
      out.hidden_series_idx.push_back(uint_of(kid(entry, "idx")).value_or(0U));
    }
  }
  return out;
}

ChartAxisPosition axis_position_of(std::string_view name) {
  if (name == "t")
    return ChartAxisPosition::kTop;
  if (name == "l")
    return ChartAxisPosition::kLeft;
  if (name == "r")
    return ChartAxisPosition::kRight;
  return ChartAxisPosition::kBottom;
}

ChartAxis read_axis(const xml_node& node, ChartAxisKind kind) {
  ChartAxis axis;
  axis.kind = kind;
  axis.id = uint_of(kid(node, "axId")).value_or(0U);
  axis.position = axis_position_of(val_of(kid(node, "axPos")));
  axis.deleted = flag_of(kid(node, "delete"));
  axis.cross_ax = uint_of(kid(node, "crossAx")).value_or(0U);
  const std::string_view crosses = val_of(kid(node, "crosses"));
  if (crosses == "autoZero") {
    axis.crosses = ChartAxisCrosses::kAutoZero;
  } else if (crosses == "min") {
    axis.crosses = ChartAxisCrosses::kMin;
  } else if (crosses == "max") {
    axis.crosses = ChartAxisCrosses::kMax;
  }
  if (const xml_node at = kid(node, "crossesAt")) {
    axis.crosses = ChartAxisCrosses::kAt;
    axis.crosses_at = double_of(at);
  }
  const xml_node scaling = kid(node, "scaling");
  axis.reversed = val_of(kid(scaling, "orientation")) == "maxMin";
  axis.log_base = double_of(kid(scaling, "logBase"));
  axis.min = double_of(kid(scaling, "min"));
  axis.max = double_of(kid(scaling, "max"));
  axis.major_unit = double_of(kid(node, "majorUnit"));
  axis.minor_unit = double_of(kid(node, "minorUnit"));
  if (const xml_node fmt = kid(node, "numFmt")) {
    axis.num_fmt = read_num_fmt(fmt);
  }
  axis.title = read_title(kid(node, "title"));
  axis.major_gridlines = static_cast<bool>(kid(node, "majorGridlines"));
  axis.minor_gridlines = static_cast<bool>(kid(node, "minorGridlines"));
  axis.tick_label_pos = val_of(kid(node, "tickLblPos"));
  axis.major_tick = val_of(kid(node, "majorTickMark"));
  axis.minor_tick = val_of(kid(node, "minorTickMark"));
  axis.base_time_unit = val_of(kid(node, "baseTimeUnit"));
  return axis;
}

// ---------------------------------------------------------------------------
// DrawingML chart (c:)
// ---------------------------------------------------------------------------

std::optional<ChartPlotType> plot_type_of(std::string_view name) {
  struct Entry {
    const char* name;
    ChartPlotType type;
  };
  static constexpr Entry kTable[] = {
      {"barChart", ChartPlotType::kBar},           {"bar3DChart", ChartPlotType::kBar3D},
      {"lineChart", ChartPlotType::kLine},         {"line3DChart", ChartPlotType::kLine3D},
      {"areaChart", ChartPlotType::kArea},         {"area3DChart", ChartPlotType::kArea3D},
      {"pieChart", ChartPlotType::kPie},           {"pie3DChart", ChartPlotType::kPie3D},
      {"doughnutChart", ChartPlotType::kDoughnut}, {"ofPieChart", ChartPlotType::kOfPie},
      {"scatterChart", ChartPlotType::kScatter},   {"bubbleChart", ChartPlotType::kBubble},
      {"radarChart", ChartPlotType::kRadar},       {"stockChart", ChartPlotType::kStock},
      {"surfaceChart", ChartPlotType::kSurface},   {"surface3DChart", ChartPlotType::kSurface3D},
  };
  for (const Entry& entry : kTable) {
    if (name == entry.name) {
      return entry.type;
    }
  }
  return std::nullopt;
}

std::optional<ChartGrouping> grouping_of(std::string_view name) {
  if (name == "clustered")
    return ChartGrouping::kClustered;
  if (name == "stacked")
    return ChartGrouping::kStacked;
  if (name == "percentStacked")
    return ChartGrouping::kPercentStacked;
  if (name == "standard")
    return ChartGrouping::kStandard;
  return std::nullopt;
}

ChartSeries read_series(const Workbook& wb, const xml_node& ser, bool filtered) {
  ChartSeries s;
  s.idx = uint_of(kid(ser, "idx")).value_or(0U);
  s.order = uint_of(kid(ser, "order")).value_or(0U);
  s.filtered = filtered;
  s.name = read_data_source(kid(ser, "tx"));
  s.cat = read_data_source(kid(ser, "cat"));
  s.val = read_data_source(kid(ser, "val"));
  s.x = read_data_source(kid(ser, "xVal"));
  s.y = read_data_source(kid(ser, "yVal"));
  s.size = read_data_source(kid(ser, "bubbleSize"));
  const xml_node sp_pr = kid(ser, "spPr");
  s.fill = read_fill(wb, sp_pr);
  s.line = read_line(wb, sp_pr);
  if (const xml_node marker = kid(ser, "marker")) {
    s.marker = read_marker(wb, marker);
  }
  s.smooth = tri_of(kid(ser, "smooth"));
  for (const xml_node& dpt : kids(ser, "dPt")) {
    ChartPoint point;
    point.idx = uint_of(kid(dpt, "idx")).value_or(0U);
    const xml_node point_sp_pr = kid(dpt, "spPr");
    point.fill = read_fill(wb, point_sp_pr);
    point.line = read_line(wb, point_sp_pr);
    if (const xml_node marker = kid(dpt, "marker")) {
      point.marker = read_marker(wb, marker);
    }
    s.points.push_back(std::move(point));
  }
  if (const xml_node dlbls = kid(ser, "dLbls")) {
    s.labels = read_labels(wb, dlbls);
  }
  return s;
}

ChartPlot read_plot(const Workbook& wb, const xml_node& node, ChartPlotType type) {
  ChartPlot plot;
  plot.type = type;
  const std::string_view bar_dir = val_of(kid(node, "barDir"));
  plot.bar_dir = bar_dir == "col"   ? ChartBarDirection::kCol
                 : bar_dir == "bar" ? ChartBarDirection::kBar
                                    : ChartBarDirection::kAbsent;
  plot.grouping = grouping_of(val_of(kid(node, "grouping")));
  plot.vary_colors = tri_of(kid(node, "varyColors"));
  plot.gap_width = uint_of(kid(node, "gapWidth"));
  plot.overlap = int_of(kid(node, "overlap"));
  plot.hole_size = uint_of(kid(node, "holeSize"));
  plot.first_slice_angle = uint_of(kid(node, "firstSliceAng"));
  plot.scatter_style = val_of(kid(node, "scatterStyle"));
  plot.radar_style = val_of(kid(node, "radarStyle"));
  if (const xml_node dlbls = kid(node, "dLbls")) {
    plot.labels = read_labels(wb, dlbls);
  }
  for (const xml_node& ax_id : kids(node, "axId")) {
    plot.axis_ids.push_back(static_cast<std::uint32_t>(ax_id.attribute("val").as_uint()));
  }
  for (const xml_node& ser : kids(node, "ser")) {
    plot.series.push_back(read_series(wb, ser, false));
  }
  // Series hidden through the chart filter move to `c15:filtered*Series`.
  for (const xml_node& ext : kids(kid(node, "extLst"), "ext")) {
    for (xml_node group = ext.first_child(); group; group = group.next_sibling()) {
      const std::string_view name = local_name(group);
      if (group.type() == pugi::node_element && name.size() > 14 && name.substr(0, 8) == "filtered" &&
          name.substr(name.size() - 6) == "Series") {
        for (const xml_node& ser : kids(group, "ser")) {
          plot.series.push_back(read_series(wb, ser, true));
        }
      }
    }
  }
  return plot;
}

/// Cells of an A1 reference list (`Sheet1!$A$2:$A$5` or a parenthesised
/// union), row-major. False for forms other than single-row / column areas.
bool expand_cells(std::string_view text, std::vector<std::pair<std::uint32_t, std::uint32_t>>& out) {
  if (text.size() >= 2 && text.front() == '(' && text.back() == ')') {
    text = text.substr(1, text.size() - 2);
  }
  std::vector<std::string> parts(1);
  bool quoted = false;
  for (const char ch : text) {
    if (ch == '\'') {
      quoted = !quoted;
    }
    if (ch == ',' && !quoted) {
      parts.emplace_back();
    } else {
      parts.back() += ch;
    }
  }
  for (std::string& part : parts) {
    bool in_quote = false;
    std::size_t bang = std::string::npos;
    for (std::size_t i = 0; i < part.size(); ++i) {
      if (part[i] == '\'') {
        in_quote = !in_quote;
      } else if (part[i] == '!' && !in_quote) {
        bang = i;
      }
    }
    std::string area = bang == std::string::npos ? part : part.substr(bang + 1);
    area.erase(std::remove(area.begin(), area.end(), '$'), area.end());
    const std::size_t colon = area.find(':');
    std::uint32_t r1 = 0;
    std::uint32_t c1 = 0;
    std::uint32_t r2 = 0;
    std::uint32_t c2 = 0;
    if (!a1::parse_a1_ref(std::string_view(area).substr(0, colon), &r1, &c1)) {
      return false;
    }
    if (colon == std::string::npos) {
      r2 = r1;
      c2 = c1;
    } else if (!a1::parse_a1_ref(std::string_view(area).substr(colon + 1), &r2, &c2)) {
      return false;
    }
    if (r2 < r1 || c2 < c1 || (r2 > r1 && c2 > c1)) {
      return false;
    }
    for (std::uint32_t r = r1; r <= r2; ++r) {
      for (std::uint32_t c = c1; c <= c2; ++c) {
        out.emplace_back(r, c);
      }
    }
  }
  return true;
}

/// Positions of the full range (`c15:fullRef`) that the visible union in
/// `c:f` leaves out.
std::vector<std::uint32_t> hidden_positions(const xml_node& ref) {
  std::vector<std::uint32_t> hidden;
  const std::string_view full_text = kid(ext_child(ref, "fullRef"), "sqref").child_value();
  const std::string_view visible_text = kid(ref, "f").child_value();
  std::vector<std::pair<std::uint32_t, std::uint32_t>> full;
  std::vector<std::pair<std::uint32_t, std::uint32_t>> visible;
  if (full_text.empty() || !expand_cells(full_text, full) || !expand_cells(visible_text, visible)) {
    return hidden;
  }
  const std::set<std::pair<std::uint32_t, std::uint32_t>> shown(visible.begin(), visible.end());
  for (std::size_t i = 0; i < full.size(); ++i) {
    if (shown.count(full[i]) == 0) {
      hidden.push_back(static_cast<std::uint32_t>(i));
    }
  }
  return hidden;
}

std::vector<std::uint32_t> read_category_filter(const xml_node& plot_area) {
  for (xml_node plot = plot_area.first_child(); plot; plot = plot.next_sibling()) {
    if (plot.type() != pugi::node_element || !plot_type_of(local_name(plot))) {
      continue;
    }
    std::vector<xml_node> sers = kids(plot, "ser");
    for (const xml_node& ext : kids(kid(plot, "extLst"), "ext")) {
      for (xml_node group = ext.first_child(); group; group = group.next_sibling()) {
        for (const xml_node& ser : kids(group, "ser")) {
          sers.push_back(ser);
        }
      }
    }
    for (const xml_node& ser : sers) {
      for (const char* tag : {"cat", "xVal", "val", "yVal"}) {
        const xml_node holder = kid(ser, tag);
        std::vector<std::uint32_t> hidden = hidden_positions(kid(holder, ""));
        if (!hidden.empty()) {
          return hidden;
        }
      }
    }
  }
  return {};
}

std::optional<std::int32_t> read_style(const xml_node& root) {
  for (const xml_node& alt : kids(root, "AlternateContent")) {
    if (const xml_node style = kid(kid(alt, "Choice"), "style")) {
      return int_of(style);
    }
  }
  if (const xml_node style = kid(root, "style")) {
    return int_of(style);
  }
  for (const xml_node& alt : kids(root, "AlternateContent")) {
    if (const xml_node style = kid(kid(alt, "Fallback"), "style")) {
      return int_of(style);
    }
  }
  return std::nullopt;
}

void read_drawingml_chart(const Workbook& wb, const xml_node& root, ChartModel& model) {
  model.is_pivot_chart = static_cast<bool>(kid(root, "pivotSource"));
  const xml_node date1904 = kid(root, "date1904");
  model.date1904 = date1904 ? flag_of(date1904) : wb.date1904();
  model.style = read_style(root);

  const xml_node chart = kid(root, "chart");
  model.title = read_title(kid(chart, "title"));
  model.auto_title_deleted = tri_of(kid(chart, "autoTitleDeleted"));
  model.legend = read_legend(kid(chart, "legend"));
  model.plot_vis_only = tri_of(kid(chart, "plotVisOnly"));
  for (const xml_node& ext : kids(kid(chart, "extLst"), "ext")) {
    if (const Tri na = tri_of(path(ext, {"dataDisplayOptions16", "dispNaAsBlank"})); na != Tri::kAbsent) {
      model.disp_na_as_blank = na;
    }
  }
  const std::string_view blanks = val_of(kid(chart, "dispBlanksAs"));
  model.disp_blanks_as = blanks == "gap"    ? ChartBlanksAs::kGap
                         : blanks == "zero" ? ChartBlanksAs::kZero
                         : blanks == "span" ? ChartBlanksAs::kSpan
                                            : ChartBlanksAs::kAbsent;

  const xml_node plot_area = kid(chart, "plotArea");
  for (xml_node node = plot_area.first_child(); node; node = node.next_sibling()) {
    if (node.type() != pugi::node_element) {
      continue;
    }
    const std::string_view name = local_name(node);
    if (const auto type = plot_type_of(name)) {
      model.plots.push_back(read_plot(wb, node, *type));
    } else if (name == "catAx") {
      model.axes.push_back(read_axis(node, ChartAxisKind::kCategory));
    } else if (name == "valAx") {
      model.axes.push_back(read_axis(node, ChartAxisKind::kValue));
    } else if (name == "dateAx") {
      model.axes.push_back(read_axis(node, ChartAxisKind::kDate));
    } else if (name == "serAx") {
      model.axes.push_back(read_axis(node, ChartAxisKind::kSeries));
    }
  }
  model.category_filter = read_category_filter(plot_area);
}

// ---------------------------------------------------------------------------
// Chartex (cx:)
// ---------------------------------------------------------------------------

/// A `cx:strDim` / `cx:numDim`: `cx:f` formula plus the `cx:lvl` cache.
ChartDataSource read_cx_dim(const xml_node& dim) {
  ChartDataSource src;
  const bool numeric = local_name(dim) == "numDim";
  const xml_node f = kid(dim, "f");
  const xml_node lvl = kid(dim, "lvl");
  if (f) {
    src.kind = ChartDataSource::Kind::kRef;
    src.ref_element = ChartDataSource::RefElement::kCxDim;
    src.formula = f.child_value();
    if (lvl) {
      read_points(lvl, numeric, true, src.cache, src.point_count);
      src.cache_format_code = lvl.attribute("formatCode").value();
    }
  } else if (lvl) {
    src.kind = ChartDataSource::Kind::kLiteral;
    read_points(lvl, numeric, true, src.literal, src.point_count);
    src.cache_format_code = lvl.attribute("formatCode").value();
  }
  return src;
}

ChartDataSource read_cx_tx(const xml_node& tx) {
  ChartDataSource src;
  const xml_node data = kid(tx, "txData");
  if (!data) {
    return src;
  }
  const xml_node v = kid(data, "v");
  if (const xml_node f = kid(data, "f")) {
    src.kind = ChartDataSource::Kind::kRef;
    src.ref_element = ChartDataSource::RefElement::kCxDim;
    src.formula = f.child_value();
    if (v) {
      src.cache.push_back(make_value(v.child_value(), false));
      src.point_count = 1;
    }
  } else if (v) {
    src.kind = ChartDataSource::Kind::kLiteral;
    src.literal.push_back(make_value(v.child_value(), false));
    src.point_count = 1;
  }
  return src;
}

ChartTitle read_cx_title(const xml_node& title) {
  ChartTitle out;
  if (!title) {
    return out;
  }
  out.kind = ChartTitleKind::kAuto;
  const xml_node tx = kid(title, "tx");
  if (const xml_node rich = kid(tx, "rich")) {
    out.kind = ChartTitleKind::kText;
    out.text = flatten_rich(rich);
  } else if (const xml_node data = kid(tx, "txData")) {
    if (const xml_node f = kid(data, "f")) {
      out.kind = ChartTitleKind::kRef;
      out.ref_formula = f.child_value();
    } else if (const xml_node v = kid(data, "v")) {
      out.kind = ChartTitleKind::kText;
      out.text = v.child_value();
    }
  }
  return out;
}

ChartLabels read_cx_labels(const xml_node& labels_node) {
  ChartLabels labels;
  const xml_node visibility = kid(labels_node, "visibility");
  labels.show_val = attr_flag(visibility, "value");
  labels.show_cat = attr_flag(visibility, "categoryName");
  labels.show_ser = attr_flag(visibility, "seriesName");
  labels.position = label_position_of(labels_node.attribute("pos").value());
  if (const xml_node fmt = kid(labels_node, "numFmt")) {
    labels.num_fmt = read_num_fmt(fmt);
  }
  if (const xml_node sep = kid(labels_node, "separator")) {
    labels.separator = std::string(sep.child_value());
  }
  return labels;
}

ChartAxis read_cx_axis(const xml_node& node) {
  ChartAxis axis;
  axis.id = static_cast<std::uint32_t>(node.attribute("id").as_uint());
  axis.deleted = attr_flag(node, "hidden");
  const xml_node val_scaling = kid(node, "valScaling");
  axis.kind = val_scaling                ? ChartAxisKind::kValue
              : kid(node, "dateScaling") ? ChartAxisKind::kDate
                                         : ChartAxisKind::kCategory;
  if (val_scaling) {
    axis.min = double_text(val_scaling.attribute("min").value());
    axis.max = double_text(val_scaling.attribute("max").value());
    axis.major_unit = double_text(val_scaling.attribute("majorUnit").value());
    axis.minor_unit = double_text(val_scaling.attribute("minorUnit").value());
  }
  if (const xml_node fmt = kid(node, "numFmt")) {
    axis.num_fmt = read_num_fmt(fmt);
  }
  axis.title = read_cx_title(kid(node, "title"));
  axis.major_gridlines = static_cast<bool>(kid(node, "majorGridlines"));
  axis.minor_gridlines = static_cast<bool>(kid(node, "minorGridlines"));
  return axis;
}

void read_chartex(const xml_node& root, ChartModel& model) {
  model.is_chartex = true;
  model.date1904 = false;

  struct Dim {
    std::string type;
    ChartDataSource source;
  };
  struct Data {
    std::string id;
    std::vector<Dim> dims;
  };
  std::vector<Data> data_sets;
  for (const xml_node& data : kids(kid(root, "chartData"), "data")) {
    Data entry;
    entry.id = data.attribute("id").value();
    for (xml_node dim = data.first_child(); dim; dim = dim.next_sibling()) {
      const std::string_view name = local_name(dim);
      if (dim.type() == pugi::node_element && (name == "strDim" || name == "numDim")) {
        entry.dims.push_back({dim.attribute("type").value(), read_cx_dim(dim)});
      }
    }
    data_sets.push_back(std::move(entry));
  }

  const xml_node chart = kid(root, "chart");
  const xml_node plot_area = kid(chart, "plotArea");
  ChartPlot plot;
  plot.type = ChartPlotType::kChartex;
  for (const xml_node& ser : kids(kid(plot_area, "plotAreaRegion"), "series")) {
    if (plot.series.empty()) {
      plot.chartex_layout = ser.attribute("layoutId").value();
    }
    ChartSeries s;
    s.idx = static_cast<std::uint32_t>(plot.series.size());
    s.order = s.idx;
    s.name = read_cx_tx(kid(ser, "tx"));
    if (const xml_node labels = kid(ser, "dataLabels")) {
      s.labels = read_cx_labels(labels);
    }
    const std::string_view data_id = val_of(kid(ser, "dataId"));
    for (const Data& data : data_sets) {
      if (data.id != data_id) {
        continue;
      }
      for (const Dim& dim : data.dims) {
        if (dim.type == "cat") {
          s.cat = dim.source;
        } else if (dim.type == "val") {
          s.val = dim.source;
        } else if (dim.type == "x") {
          s.x = dim.source;
        } else if (dim.type == "y") {
          s.y = dim.source;
        } else if (dim.type == "size") {
          s.size = dim.source;
        }
      }
    }
    for (const xml_node& ax_id : kids(ser, "axisId")) {
      plot.axis_ids.push_back(static_cast<std::uint32_t>(ax_id.attribute("val").as_uint()));
    }
    plot.series.push_back(std::move(s));
  }
  model.plots.push_back(std::move(plot));

  model.title = read_cx_title(kid(chart, "title"));
  if (const xml_node legend = kid(chart, "legend")) {
    model.legend.present = true;
    model.legend.position = legend_position_of(legend.attribute("pos").value(), ChartLegendPosition::kRight);
    model.legend.overlay = attr_flag(legend, "overlay");
  }
  for (const xml_node& axis : kids(plot_area, "axis")) {
    model.axes.push_back(read_cx_axis(axis));
  }
}

}  // namespace

Expected<void, Error> parse_chart_part(const std::vector<std::uint8_t>& bytes, pugi::xml_document& doc) {
  if (!io::load_xml_buffer(doc, bytes, "chart", "chart")) {
    return unparseable("part is not well-formed XML", {});
  }
  if (local_name(doc.document_element()) != "chartSpace") {
    return unparseable("part root is not chartSpace", doc.document_element().name());
  }
  return Expected<void, Error>::Ok();
}

Expected<ChartModel, Error> read_chart_model(const Workbook& wb, const pugi::xml_document& doc) {
  const xml_node root = doc.document_element();
  if (local_name(root) != "chartSpace") {
    return unparseable("part root is not chartSpace", root.name());
  }
  ChartModel model;
  if (kid(root, "chartData")) {
    read_chartex(root, model);
  } else {
    read_drawingml_chart(wb, root, model);
  }
  return model;
}

}  // namespace formulon
