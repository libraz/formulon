#ifndef FORMULON_CHART_CHART_EDIT_INTERNAL_H_
#define FORMULON_CHART_CHART_EDIT_INTERNAL_H_

// Chart-part editing primitives shared by the `chart_edit*.cpp` translation units.

#include <cstddef>
#include <cstdint>
#include <set>
#include <string>
#include <string_view>
#include <vector>

#include "chart/chart_data.h"
#include "chart/chart_edit.h"
#include "chart/chart_ns.h"
#include "chart/chart_template.h"
#include "drawing/drawing_xml.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "utils/double_format.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/number_text.h"

namespace formulon {

class Workbook;

namespace chart_edit {

using pugi::xml_node;
using Order = const char* const*;
using chart_ns::kNsA;
using chart_ns::kNsC;
using chart_ns::kNsC15;
using chart_ns::kUriC15;
using chart_ns::kUriLeader;

// Child orders of the CT_ types edits insert into (unions where siblings share a parent tag).
// clang-format off
inline constexpr const char* const kChartOrder[] = {"title", "autoTitleDeleted", "pivotFmts", "view3D", "floor", "sideWall",
    "backWall", "plotArea", "legend", "plotVisOnly", "dispBlanksAs", "showDLblsOverMax", "extLst", nullptr};
inline constexpr const char* const kTitleOrder[] = {"tx", "layout", "overlay", "spPr", "txPr", "extLst", nullptr};
inline constexpr const char* const kLegendOrder[] = {"legendPos", "legendEntry", "layout", "overlay", "spPr", "txPr", "extLst",
    nullptr};
inline constexpr const char* const kPlotOrder[] = {"barDir", "grouping", "scatterStyle", "radarStyle", "varyColors", "ser",
    "dLbls", "gapWidth", "overlap", "serLines", "dropLines", "hiLowLines", "upDownBars", "marker", "smooth",
    "firstSliceAng", "holeSize", "axId", "extLst", nullptr};
inline constexpr const char* const kSerOrder[] = {"idx", "order", "tx", "spPr", "invertIfNegative", "pictureOptions", "marker",
    "explosion", "dPt", "dLbls", "trendline", "errBars", "cat", "val", "xVal", "yVal", "shape", "smooth",
    "bubbleSize", "bubble3D", "extLst", nullptr};
inline constexpr const char* const kPointOrder[] = {"idx", "invertIfNegative", "marker", "bubble3D", "explosion", "spPr",
    "pictureOptions", "extLst", nullptr};
inline constexpr const char* const kMarkerOrder[] = {"symbol", "size", "spPr", "extLst", nullptr};
inline constexpr const char* const kLabelsOrder[] = {"dLbl", "delete", "numFmt", "spPr", "txPr", "dLblPos", "showLegendKey",
    "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize", "separator", "showLeaderLines",
    "leaderLines", "extLst", nullptr};
inline constexpr const char* const kAxisOrder[] = {"axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines",
    "title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses",
    "crossesAt", "crossBetween", "auto", "lblAlgn", "lblOffset", "baseTimeUnit", "majorUnit", "majorTimeUnit",
    "minorUnit", "minorTimeUnit", "tickLblSkip", "tickMarkSkip", "noMultiLvlLbl", "dispUnits", "extLst", nullptr};
inline constexpr const char* const kScalingOrder[] = {"logBase", "orientation", "max", "min", "extLst", nullptr};
inline constexpr const char* const kShapeOrder[] = {"xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "gradFill", "blipFill",
    "pattFill", "grpFill", "ln", "effectLst", "effectDag", "scene3d", "sp3d", "extLst", nullptr};
inline constexpr const char* const kLineOrder[] = {"noFill", "solidFill", "gradFill", "pattFill", "prstDash", "custDash", "round",
    "bevel", "miter", "headEnd", "tailEnd", "extLst", nullptr};
inline constexpr const char* const kRunOrder[] = {"ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill",
    "effectLst", "effectDag", "highlight", "uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea", "cs", "sym",
    "hlinkClick", "hlinkMouseOver", "rtl", "extLst", nullptr};
inline constexpr const char* const kTextOrder[] = {"bodyPr", "lstStyle", "p", nullptr};

inline constexpr const char* const kFills[] = {"noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", nullptr};
// clang-format on

inline constexpr const char* const kRoles[] = {"tx", "cat", "val", "xVal", "yVal", "bubbleSize"};

struct Chart {
  std::string path;
  std::size_t host = kChartNoHostSheet;
  bool chartex = false;
  pugi::xml_document doc;
  std::string before;  // Serialised root at open; commit writes only when it differs.

  xml_node root() const { return doc.document_element(); }
  xml_node chart() const { return child_local(root(), "chart"); }
  xml_node area() const { return child_local(chart(), "plotArea"); }
  bool pivot() const { return static_cast<bool>(child_local(root(), "pivotSource")); }
  std::string lang() const {
    const std::string_view v = child_local(root(), "lang").attribute("val").value();
    return v.empty() ? std::string("en-US") : std::string(v);
  }
};

/// A chart's style parts as marker settling needs them: Excel completes marker outlines only in a chart with
/// a chart style part, in the colours of its colour style.
struct Palette {
  bool styled = false;
  pugi::xml_document colors;
};

// `c:spPr` of the data-label leader lines.
inline std::string leader_sp_pr() {
  std::string x = "<c:spPr>";
  append_grey_ln(x, 35);
  return x + "<a:effectLst/></c:spPr>";
}

inline std::string gridlines_xml() {
  std::string x = "<c:majorGridlines><c:spPr>";
  append_grey_ln(x, 15);
  return x + "<a:effectLst/></c:spPr></c:majorGridlines>";
}

inline Expected<void, Error> ok() {
  return Expected<void, Error>::Ok();
}

inline Error bad_arg(const char* message) {
  return make_error(FormulonErrorCode::kInvalidArgument, message);
}

inline Error unsupported(const char* message) {
  return make_error(FormulonErrorCode::kIoChartUnsupportedKind, message);
}

inline Error unparseable(const char* message) {
  return make_error(FormulonErrorCode::kIoChartUnparseable, message);
}

int rank(Order order, std::string_view name);

/// The child a new `name` element goes before under `order`; null appends.
xml_node slot(const xml_node& parent, std::string_view name, Order order);

inline xml_node ensure(xml_node parent, const char* qname, Order order) {
  if (const xml_node found = child_local(parent, local_name(qname))) {
    return found;
  }
  const xml_node before = slot(parent, local_name(qname), order);
  return before ? parent.insert_child_before(qname, before) : parent.append_child(qname);
}

/// Copies `src` into `parent` at its schema position, renamed to `qname` when given.
xml_node copy_at(xml_node parent, const xml_node& src, Order order, const char* qname = nullptr);

inline xml_node put_val(xml_node parent, const char* qname, Order order, const std::string& value) {
  xml_node node = ensure(parent, qname, order);
  io::set_attr(node, "val", value);
  return node;
}

inline void drop(xml_node parent, std::string_view name) {
  for (const xml_node& node : children_local(parent, name)) {
    parent.remove_child(node);
  }
}

inline std::uint32_t val_u(const xml_node& parent, std::string_view name) {
  return child_local(parent, name).attribute("val").as_uint();
}

inline bool flag_attr(const xml_node& node, const char* name) {
  const std::string_view v = node.attribute(name).value();
  return v == "1" || v == "true";
}

inline bool flag(const xml_node& node) {
  return flag_attr(node, "val");
}

inline bool load_xml(pugi::xml_document& doc, const std::string& xml) {
  return static_cast<bool>(doc.load_buffer(xml.data(), xml.size()));
}

inline std::string num(double v) {
  std::string out;
  format_double(out, v);
  return out;
}

inline std::string rounded(double v) {
  return std::to_string(std::llround(v));
}

inline bool in_range(double v, double lo, double hi) {
  return std::isfinite(v) && v >= lo && v <= hi;
}

inline std::string hex6(std::uint32_t argb) {
  char buf[8];
  format_hex(buf, sizeof(buf), argb & 0xFFFFFFU, 6, true);
  return buf;
}

Expected<void, Error> open_chart(const Workbook& wb, std::size_t sheet, std::uint32_t object_id, Chart& c);

Expected<void, Error> open_edit(const Workbook& wb, std::size_t sheet, std::uint32_t object_id, Chart& c,
                                bool series_op);

/// Stores the part when the edit changed it, refreshing caches after reference edits.
Expected<void, Error> commit(Workbook& wb, Chart& c, bool refresh);

inline bool is_plot(const xml_node& node) {
  const std::string_view name = local_name(node);
  return node.type() == pugi::node_element && name.size() > 5 && name.substr(name.size() - 5) == "Chart";
}

std::vector<xml_node> plots_of(const Chart& c);

inline Expected<xml_node, Error> plot_at(const Chart& c, std::size_t index) {
  const std::vector<xml_node> plots = plots_of(c);
  if (index >= plots.size()) {
    return bad_arg("chart: plot index out of range");
  }
  return plots[index];
}

/// The `ChartInsertKind` a plot is, or -1 outside the six.
int kind_of(const xml_node& plot);

inline ChartGrouping grouping_of(const xml_node& plot) {
  const std::string_view v = child_local(plot, "grouping").attribute("val").value();
  return v == "percentStacked" ? ChartGrouping::kPercentStacked
         : v == "stacked"      ? ChartGrouping::kStacked
                               : ChartGrouping::kClustered;
}

inline bool has_markers(const xml_node& plot) {
  const std::string_view n = local_name(plot);
  return n == "lineChart" || n == "line3DChart" || n == "scatterChart" || n == "radarChart" || n == "stockChart";
}

inline bool is_bar(const xml_node& plot) {
  const std::string_view n = local_name(plot);
  return n == "barChart" || n == "bar3DChart";
}

/// A line or scatter series and its points are filled through the marker; Excel drops a fill on their own
/// shape properties, with or without transparency.
inline bool marker_filled(const xml_node& plot) {
  const std::string_view n = local_name(plot);
  return n == "lineChart" || n == "scatterChart";
}

inline bool is_pie(const xml_node& plot) {
  const std::string_view n = local_name(plot);
  return n == "pieChart" || n == "pie3DChart" || n == "doughnutChart" || n == "ofPieChart";
}

/// The `c15:filtered*Series` group of a plot, or null.
xml_node filter_group(const xml_node& plot);

inline std::vector<xml_node> plot_series(const xml_node& plot) {
  std::vector<xml_node> out = children_local(plot, "ser");
  for (const xml_node& s : children_local(filter_group(plot), "ser")) {
    out.push_back(s);
  }
  return out;
}

Expected<xml_node, Error> series_at(const xml_node& plot, std::uint32_t idx);

/// The `strRef` / `numRef` / `multiLvlStrRef` of a series input, or null.
inline xml_node ref_of(const xml_node& holder) {
  const xml_node ref = io::first_element_child(holder);
  const std::string_view n = local_name(ref);
  return n.size() > 3 && n.substr(n.size() - 3) == "Ref" ? ref : xml_node();
}

/// Declares `c15` on `node` unless it already resolves there to the c15 namespace.
inline void bind_c15(xml_node node) {
  if (io::find_namespace_binding(node, "c15") != kNsC15) {
    node.append_attribute("xmlns:c15").set_value(kNsC15);
  }
}

/// The `c:ext` holding a ref's c15 children; `create` makes it the ref's first
/// child and binds `c15` on it for the children the caller appends.
xml_node c15_ext(xml_node ref, bool create);

/// Removes an empty c15 ext and an empty extLst from `ref`.
void prune_ext(xml_node ref);

/// The element holding a ref's formula: `c:f`, or a filtered series' `c15:sqref`.
inline xml_node formula_node(const xml_node& ref) {
  if (const xml_node f = child_local(ref, "f")) {
    return f;
  }
  return child_local(child_local(c15_ext(ref, false), "formulaRef"), "sqref");
}

inline std::string formula_of(const xml_node& holder) {
  return formula_node(ref_of(holder)).child_value();
}

std::uint32_t point_count(const xml_node& ser);

inline xml_node fragment(pugi::xml_document& frag, const std::string& xml) {
  return load_xml(frag, xml) ? frag.first_child() : xml_node();
}

void load_palette(const Workbook& wb, const std::string& chart_path, Palette& out);

/// A chart part from the insert template, for its default elements, in the colours of chart `c`.
Expected<void, Error> template_doc(const Workbook& wb, const Chart& c, ChartInsertKind kind, ChartGrouping grouping,
                                   const std::optional<std::string>& title, ChartLegendPosition legend,
                                   ChartSplit split, pugi::xml_document& out);

inline xml_node template_chart(const pugi::xml_document& t) {
  return child_local(t.document_element(), "chart");
}

xml_node template_plot(const pugi::xml_document& t);

inline bool fill_empty(const ChartFillPatch& p) {
  return p.color.is_keep() && p.transparency_pct.is_keep();
}

inline bool line_empty(const ChartLinePatch& p) {
  return fill_empty(ChartFillPatch{p.color, p.transparency_pct}) && p.width_pt.is_keep() && p.dash.is_keep() &&
         !p.none.has_value();
}

Expected<void, Error> apply_shape(xml_node owner, Order owner_order, const ChartFillPatch& fill,
                                  const ChartLinePatch& line);

Expected<void, Error> apply_marker(xml_node owner, Order order, const ChartPatchField<ChartMarkerPatch>& patch);

/// Completes the marker of `owner` (series `ser` or one of its points) to the form Excel re-saves unchanged.
/// An automatic symbol becomes the series' own (a series' by idx); a hidden marker keeps only its symbol; a
/// shown one takes size 5 and, in a styled chart, a fill and a 0.75 pt outline in the series colour where unset.
void settle_marker(xml_node ser, xml_node owner, const Palette& p);

Expected<void, Error> apply_labels(xml_node ser, bool pie, const std::string& lang,
                                   const ChartPatchField<ChartLabelsPatch>& patch);

void drop_empty_group(xml_node plot);

/// Moves series between the plot and its `c15:filtered*Series` group, in the form Excel saves.
void patch_series_filter(const Chart& c, const std::set<std::uint32_t>& hidden);

/// Points each series' category / value refs at the visible cells, keeping
/// the full range in `c15:fullRef`, as Excel saves it. Positions index the full range.
Expected<void, Error> patch_category_filter(const Workbook& wb, const Chart& c, const std::set<std::uint32_t>& hidden);

std::set<std::uint32_t> filtered_series(const Chart& c);

Expected<void, Error> patch_series(const Workbook& wb, const Chart& c, xml_node plot, xml_node ser,
                                   const ChartSeriesPatch& p, bool* refs);

/// Inserts a copy of `src` among `ser`'s `c:dPt` in idx order.
xml_node insert_point(xml_node ser, const xml_node& src, std::uint32_t idx);

xml_node find_point(const xml_node& ser, std::uint32_t idx);

/// Excel's column vary-colours form: per-point accent fills replace the series fill.
void vary_bar(xml_node plot, bool on, const Palette& palette);

void patch_display(const Chart& c, const ChartDisplayPatch& patch);

std::vector<xml_node> axes_of(const Chart& c, const std::vector<std::uint32_t>& ids);

/// Puts `ax` on the side Excel derives on save: the far side (t / r) when it crosses at the maximum of its
/// crossing axis, with the sides swapped while that axis is reversed.
void place_axis(const Chart& c, xml_node ax);

Expected<void, Error> patch_type(const Workbook& wb, Chart& c, std::size_t plot_index, ChartInsertKind kind,
                                 ChartGrouping grouping);

}  // namespace chart_edit
}  // namespace formulon

#endif  // FORMULON_CHART_CHART_EDIT_INTERNAL_H_
