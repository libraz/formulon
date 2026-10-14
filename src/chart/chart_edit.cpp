#include "chart/chart_edit.h"

#include <algorithm>
#include <cmath>
#include <cstring>
#include <initializer_list>
#include <optional>
#include <set>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "chart/chart_cache.h"
#include "chart/chart_data.h"
#include "chart/chart_ns.h"
#include "chart/chart_source.h"
#include "chart/chart_template.h"
#include "chart/chart_xml.h"
#include "drawing/drawing_edit.h"
#include "drawing/drawing_xml.h"
#include "eval/function_registry.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml_defs.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "unknown_relationship.h"
#include "utils/double_format.h"
#include "utils/status_macros.h"
#include "workbook.h"

namespace formulon {
namespace {

using pugi::xml_node;
using Order = const char* const*;

constexpr const char* kRelChart = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
constexpr const char* kCtChart = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
using chart_ns::kNsA;
using chart_ns::kNsC;
using chart_ns::kNsC15;
using chart_ns::kUriC15;
using chart_ns::kUriLeader;
// Excel's default chart frame, 360 x 216 pt.
constexpr std::int64_t kFrameCx = 4572000;
constexpr std::int64_t kFrameCy = 2743200;

// Child orders of the CT_ types edits insert into (unions where siblings share a parent tag).
// clang-format off
constexpr const char* kChartOrder[] = {"title", "autoTitleDeleted", "pivotFmts", "view3D", "floor", "sideWall",
    "backWall", "plotArea", "legend", "plotVisOnly", "dispBlanksAs", "showDLblsOverMax", "extLst", nullptr};
constexpr const char* kTitleOrder[] = {"tx", "layout", "overlay", "spPr", "txPr", "extLst", nullptr};
constexpr const char* kLegendOrder[] = {"legendPos", "legendEntry", "layout", "overlay", "spPr", "txPr", "extLst",
    nullptr};
constexpr const char* kPlotOrder[] = {"barDir", "grouping", "scatterStyle", "radarStyle", "varyColors", "ser",
    "dLbls", "gapWidth", "overlap", "serLines", "dropLines", "hiLowLines", "upDownBars", "marker", "smooth",
    "firstSliceAng", "holeSize", "axId", "extLst", nullptr};
constexpr const char* kSerOrder[] = {"idx", "order", "tx", "spPr", "invertIfNegative", "pictureOptions", "marker",
    "explosion", "dPt", "dLbls", "trendline", "errBars", "cat", "val", "xVal", "yVal", "shape", "smooth",
    "bubbleSize", "bubble3D", "extLst", nullptr};
constexpr const char* kPointOrder[] = {"idx", "invertIfNegative", "marker", "bubble3D", "explosion", "spPr",
    "pictureOptions", "extLst", nullptr};
constexpr const char* kMarkerOrder[] = {"symbol", "size", "spPr", "extLst", nullptr};
constexpr const char* kLabelsOrder[] = {"dLbl", "delete", "numFmt", "spPr", "txPr", "dLblPos", "showLegendKey",
    "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize", "separator", "showLeaderLines",
    "leaderLines", "extLst", nullptr};
constexpr const char* kAxisOrder[] = {"axId", "scaling", "delete", "axPos", "majorGridlines", "minorGridlines",
    "title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crosses",
    "crossesAt", "crossBetween", "auto", "lblAlgn", "lblOffset", "baseTimeUnit", "majorUnit", "majorTimeUnit",
    "minorUnit", "minorTimeUnit", "tickLblSkip", "tickMarkSkip", "noMultiLvlLbl", "dispUnits", "extLst", nullptr};
constexpr const char* kScalingOrder[] = {"logBase", "orientation", "max", "min", "extLst", nullptr};
constexpr const char* kShapeOrder[] = {"xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "gradFill", "blipFill",
    "pattFill", "grpFill", "ln", "effectLst", "effectDag", "scene3d", "sp3d", "extLst", nullptr};
constexpr const char* kLineOrder[] = {"noFill", "solidFill", "gradFill", "pattFill", "prstDash", "custDash", "round",
    "bevel", "miter", "headEnd", "tailEnd", "extLst", nullptr};
constexpr const char* kRunOrder[] = {"ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill",
    "effectLst", "effectDag", "highlight", "uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea", "cs", "sym",
    "hlinkClick", "hlinkMouseOver", "rtl", "extLst", nullptr};
constexpr const char* kTextOrder[] = {"bodyPr", "lstStyle", "p", nullptr};
constexpr const char* kFills[] = {"noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", nullptr};
// clang-format on

constexpr const char* kLeaderSpPr =
    "<c:spPr><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"tx1\">"
    "<a:lumMod val=\"35000\"/><a:lumOff val=\"65000\"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/>"
    "</c:spPr>";
constexpr const char* kGridlines =
    "<c:majorGridlines><c:spPr><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill>"
    "<a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill>"
    "<a:round/></a:ln><a:effectLst/></c:spPr></c:majorGridlines>";

Expected<void, Error> ok() {
  return Expected<void, Error>::Ok();
}
Error bad_arg(const char* message) {
  return make_error(FormulonErrorCode::kInvalidArgument, message);
}
Error unsupported(const char* message) {
  return make_error(FormulonErrorCode::kIoChartUnsupportedKind, message);
}
Error unparseable(const char* message) {
  return make_error(FormulonErrorCode::kIoChartUnparseable, message);
}

// ---------------------------------------------------------------------------
// DOM helpers
// ---------------------------------------------------------------------------

std::string_view local_of(const char* qname) {
  const char* colon = std::strchr(qname, ':');
  return colon != nullptr ? std::string_view(colon + 1) : std::string_view(qname);
}

xml_node first_element(const xml_node& parent) {
  for (xml_node c = parent.first_child(); c; c = c.next_sibling()) {
    if (c.type() == pugi::node_element) {
      return c;
    }
  }
  return {};
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

int rank(Order order, std::string_view name) {
  for (int i = 0; order[i] != nullptr; ++i) {
    if (name == order[i]) {
      return i;
    }
  }
  return -1;
}

/// The child a new `name` element goes before under `order`; null appends.
xml_node slot(const xml_node& parent, std::string_view name, Order order) {
  const int r = rank(order, name);
  for (xml_node c = parent.first_child(); c; c = c.next_sibling()) {
    if (c.type() == pugi::node_element && rank(order, local_name(c)) > r) {
      return c;
    }
  }
  return {};
}

xml_node ensure(xml_node parent, const char* qname, Order order) {
  if (const xml_node found = child_local(parent, local_of(qname))) {
    return found;
  }
  const xml_node before = slot(parent, local_of(qname), order);
  return before ? parent.insert_child_before(qname, before) : parent.append_child(qname);
}

/// Copies `src` into `parent` at its schema position, renamed to `qname` when given.
xml_node copy_at(xml_node parent, const xml_node& src, Order order, const char* qname = nullptr) {
  xml_node node = parent.append_copy(src);
  if (qname != nullptr) {
    node.set_name(qname);
  }
  const xml_node before = slot(parent, local_name(node), order);
  return before ? parent.insert_move_before(node, before) : node;
}

void set_attr(xml_node node, const char* name, const std::string& value) {
  pugi::xml_attribute attr = node.attribute(name);
  if (!attr) {
    attr = node.append_attribute(name);
  }
  attr.set_value(value.c_str());
}

xml_node put_val(xml_node parent, const char* qname, Order order, const std::string& value) {
  xml_node node = ensure(parent, qname, order);
  set_attr(node, "val", value);
  return node;
}

void drop(xml_node parent, std::string_view name) {
  for (const xml_node& node : kids(parent, name)) {
    parent.remove_child(node);
  }
}

std::uint32_t val_u(const xml_node& parent, std::string_view name) {
  return child_local(parent, name).attribute("val").as_uint();
}

bool flag_attr(const xml_node& node, const char* name) {
  const std::string_view v = node.attribute(name).value();
  return v == "1" || v == "true";
}

bool flag(const xml_node& node) {
  return flag_attr(node, "val");
}

bool load_xml(pugi::xml_document& doc, const std::string& xml) {
  return static_cast<bool>(doc.load_buffer(xml.data(), xml.size()));
}

std::string num(double v) {
  std::string out;
  format_double(out, v);
  return out;
}

std::string rounded(double v) {
  return std::to_string(std::llround(v));
}

bool in_range(double v, double lo, double hi) {
  return std::isfinite(v) && v >= lo && v <= hi;
}

std::string hex6(std::uint32_t argb) {
  static const char kHex[] = "0123456789ABCDEF";
  std::string out(6, '0');
  for (std::size_t i = 0; i < 6; ++i) {
    out[5 - i] = kHex[(argb >> (4 * i)) & 0xFU];
  }
  return out;
}

// ---------------------------------------------------------------------------
// The chart being read or edited
// ---------------------------------------------------------------------------

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

Expected<void, Error> open_chart(const Workbook& wb, std::size_t sheet, std::uint32_t object_id, Chart& c) {
  ASSIGN_OR_RETURN(std::vector<ChartPartRef> parts, list_chart_parts(wb, sheet));
  const auto it =
      std::find_if(parts.begin(), parts.end(), [&](const ChartPartRef& p) { return p.object_id == object_id; });
  if (it == parts.end()) {
    return bad_arg("chart: no chart with that id");
  }
  const PassthroughPart* part = it->part_path.empty() ? nullptr : io::ooxml::find_passthrough_part(wb, it->part_path);
  if (part == nullptr) {
    return unparseable("chart: chart part is missing");
  }
  RETURN_IF_ERROR(parse_chart_part(part->bytes, c.doc));
  c.path = it->part_path;
  c.chartex = it->is_chartex || static_cast<bool>(child_local(c.root(), "chartData"));
  c.host = wb.sheet(sheet).is_opaque_ooxml_sheet() ? kChartNoHostSheet : sheet;
  return ok();
}

Expected<void, Error> open_edit(const Workbook& wb, std::size_t sheet, std::uint32_t object_id, Chart& c,
                                bool series_op) {
  RETURN_IF_ERROR(open_chart(wb, sheet, object_id, c));
  const auto& rels = wb.sheet(sheet).unknown_relationships();
  if (std::any_of(rels.begin(), rels.end(), [](const UnknownRelationship& r) { return r.type == io::kRelDrawing; })) {
    return make_error(FormulonErrorCode::kIoDrawingUnparseable, "chart: a chart retained from xlsb is not editable");
  }
  if (c.chartex) {
    return unsupported("chart: chartex charts are not editable");
  }
  if (series_op && c.pivot()) {
    return unsupported("chart: a pivot chart's series follow its pivot table");
  }
  c.before = io::raw_xml(c.root());
  return ok();
}

/// Stores the part when the edit changed it, refreshing caches after reference edits.
Expected<void, Error> commit(Workbook& wb, Chart& c, bool refresh) {
  xml_node root = c.root();
  if (io::raw_xml(root) == c.before) {
    return ok();
  }
  if (!root.attribute("xmlns:c")) {
    root.append_attribute("xmlns:c").set_value(kNsC);
  }
  if (!root.attribute("xmlns:a")) {
    root.append_attribute("xmlns:a").set_value(kNsA);
  }
  if (refresh) {
    refresh_chart_caches(wb, c.host, c.doc);
  }
  return io::ooxml::store_part_dom(wb, c.path, c.doc);
}

bool is_plot(const xml_node& node) {
  const std::string_view name = local_name(node);
  return node.type() == pugi::node_element && name.size() > 5 && name.substr(name.size() - 5) == "Chart";
}

std::vector<xml_node> plots_of(const Chart& c) {
  std::vector<xml_node> out;
  for (xml_node n = c.area().first_child(); n; n = n.next_sibling()) {
    if (is_plot(n)) {
      out.push_back(n);
    }
  }
  return out;
}

Expected<xml_node, Error> plot_at(const Chart& c, std::size_t index) {
  const std::vector<xml_node> plots = plots_of(c);
  if (index >= plots.size()) {
    return bad_arg("chart: plot index out of range");
  }
  return plots[index];
}

/// The `ChartInsertKind` a plot is, or -1 outside the six.
int kind_of(const xml_node& plot) {
  const std::string_view name = local_name(plot);
  if (name == "barChart") {
    return std::string_view(child_local(plot, "barDir").attribute("val").value()) == "bar"
               ? static_cast<int>(ChartInsertKind::kBar)
               : static_cast<int>(ChartInsertKind::kColumn);
  }
  static const char* const kNames[] = {"lineChart", "areaChart", "pieChart", "scatterChart", "doughnutChart"};
  for (int i = 0; i < 5; ++i) {
    if (name == kNames[i]) {
      return static_cast<int>(ChartInsertKind::kLine) + i;
    }
  }
  return -1;
}

ChartGrouping grouping_of(const xml_node& plot) {
  const std::string_view v = child_local(plot, "grouping").attribute("val").value();
  return v == "percentStacked" ? ChartGrouping::kPercentStacked
         : v == "stacked"      ? ChartGrouping::kStacked
                               : ChartGrouping::kClustered;
}

bool has_markers(const xml_node& plot) {
  const std::string_view n = local_name(plot);
  return n == "lineChart" || n == "line3DChart" || n == "scatterChart" || n == "radarChart" || n == "stockChart";
}

bool is_bar(const xml_node& plot) {
  const std::string_view n = local_name(plot);
  return n == "barChart" || n == "bar3DChart";
}

/// A line or scatter series and its points are filled through the marker; Excel drops a fill on their own
/// shape properties, with or without transparency.
bool marker_filled(const xml_node& plot) {
  const std::string_view n = local_name(plot);
  return n == "lineChart" || n == "scatterChart";
}

bool is_pie(const xml_node& plot) {
  const std::string_view n = local_name(plot);
  return n == "pieChart" || n == "pie3DChart" || n == "doughnutChart" || n == "ofPieChart";
}

/// The `c15:filtered*Series` group of a plot, or null.
xml_node filter_group(const xml_node& plot) {
  for (const xml_node& ext : kids(child_local(plot, "extLst"), "ext")) {
    for (xml_node g = ext.first_child(); g; g = g.next_sibling()) {
      const std::string_view n = local_name(g);
      if (g.type() == pugi::node_element && n.size() > 14 && n.substr(0, 8) == "filtered" &&
          n.substr(n.size() - 6) == "Series") {
        return g;
      }
    }
  }
  return {};
}

std::vector<xml_node> plot_series(const xml_node& plot) {
  std::vector<xml_node> out = kids(plot, "ser");
  for (const xml_node& s : kids(filter_group(plot), "ser")) {
    out.push_back(s);
  }
  return out;
}

Expected<xml_node, Error> series_at(const xml_node& plot, std::uint32_t idx) {
  for (const xml_node& s : plot_series(plot)) {
    if (val_u(s, "idx") == idx) {
      return s;
    }
  }
  return bad_arg("chart: no series with that idx");
}

/// The `strRef` / `numRef` / `multiLvlStrRef` of a series input, or null.
xml_node ref_of(const xml_node& holder) {
  const xml_node ref = first_element(holder);
  const std::string_view n = local_name(ref);
  return n.size() > 3 && n.substr(n.size() - 3) == "Ref" ? ref : xml_node();
}

/// The `c:ext` holding a ref's c15 children; `create` makes it the ref's first child.
xml_node c15_ext(xml_node ref, bool create, bool declare) {
  xml_node list = child_local(ref, "extLst");
  for (const xml_node& ext : kids(list, "ext")) {
    if (std::string_view(ext.attribute("uri").value()) == kUriC15) {
      return ext;
    }
  }
  if (!create) {
    return {};
  }
  if (!list) {
    list = ref.prepend_child("c:extLst");
  }
  xml_node ext = list.prepend_child("c:ext");
  ext.append_attribute("uri").set_value(kUriC15);
  if (declare) {
    ext.append_attribute("xmlns:c15").set_value(kNsC15);
  }
  return ext;
}

/// Removes an empty c15 ext and an empty extLst from `ref`.
void prune_ext(xml_node ref) {
  xml_node list = child_local(ref, "extLst");
  for (const xml_node& ext : kids(list, "ext")) {
    if (!first_element(ext)) {
      list.remove_child(ext);
    }
  }
  if (list && !first_element(list)) {
    ref.remove_child(list);
  }
}

/// The element holding a ref's formula: `c:f`, or a filtered series' `c15:sqref`.
xml_node formula_node(const xml_node& ref) {
  if (const xml_node f = child_local(ref, "f")) {
    return f;
  }
  return child_local(child_local(c15_ext(ref, false, false), "formulaRef"), "sqref");
}

std::string formula_of(const xml_node& holder) {
  return formula_node(ref_of(holder)).child_value();
}

std::uint32_t point_count(const xml_node& ser) {
  const xml_node val = child_local(ser, "val") ? child_local(ser, "val") : child_local(ser, "yVal");
  const xml_node ref = ref_of(val);
  for (xml_node c = ref.first_child(); c; c = c.next_sibling()) {
    if (child_local(c, "ptCount")) {
      return val_u(c, "ptCount");
    }
  }
  return 0;
}

xml_node fragment(pugi::xml_document& frag, const std::string& xml) {
  return load_xml(frag, xml) ? frag.first_child() : xml_node();
}

/// A chart part from the insert template, for its default elements.
Expected<void, Error> template_doc(const Workbook& wb, ChartInsertKind kind, ChartGrouping grouping,
                                   const std::optional<std::string>& title, ChartLegendPosition legend,
                                   ChartSplit split, pugi::xml_document& out) {
  if (split.series.empty()) {
    split.series.push_back(ChartSeriesRefs{"", "", "A1", "A1", "A1"});
  }
  ChartInsertOptions options;
  options.kind = kind;
  options.grouping = grouping;
  options.title = title;
  options.legend = legend;
  ASSIGN_OR_RETURN(std::string xml, build_chart_xml(wb, options, split));
  if (!load_xml(out, xml)) {
    return unparseable("chart: template does not parse");
  }
  return ok();
}

xml_node template_chart(const pugi::xml_document& t) {
  return child_local(t.document_element(), "chart");
}

xml_node template_plot(const pugi::xml_document& t) {
  for (xml_node n = child_local(template_chart(t), "plotArea").first_child(); n; n = n.next_sibling()) {
    if (is_plot(n)) {
      return n;
    }
  }
  return {};
}

// ---------------------------------------------------------------------------
// Shape properties, markers, labels
// ---------------------------------------------------------------------------

bool fill_empty(const ChartFillPatch& p) {
  return p.color.is_keep() && p.transparency_pct.is_keep();
}

bool line_empty(const ChartLinePatch& p) {
  return fill_empty(ChartFillPatch{p.color, p.transparency_pct}) && p.width_pt.is_keep() && p.dash.is_keep() &&
         !p.none.has_value();
}

/// Colour and transparency of `holder`'s solid fill (`a:spPr`, `a:ln` or `a:defRPr`).
Expected<void, Error> apply_paint(xml_node holder, Order order, const ChartPatchField<std::uint32_t>& color,
                                  const ChartPatchField<double>& transparency) {
  if (color.is_clear()) {
    drop(holder, "solidFill");
  } else if (color.is_set()) {
    const std::string alpha =
        child_local(first_element(child_local(holder, "solidFill")), "alpha").attribute("val").value();
    for (Order f = kFills; *f != nullptr; ++f) {
      drop(holder, *f);
    }
    xml_node clr = ensure(holder, "a:solidFill", order).append_child("a:srgbClr");
    set_attr(clr, "val", hex6(color.value));
    if (!alpha.empty()) {
      set_attr(clr.append_child("a:alpha"), "val", alpha);
    }
  }
  if (transparency.is_keep()) {
    return ok();
  }
  xml_node clr = first_element(child_local(holder, "solidFill"));
  if (transparency.is_clear()) {
    drop(clr, "alpha");
    return ok();
  }
  if (!clr) {
    return bad_arg("chart: transparency needs a solid colour");
  }
  if (!in_range(transparency.value, 0.0, 100.0)) {
    return bad_arg("chart: transparency must be 0..100");
  }
  xml_node alpha = child_local(clr, "alpha");
  set_attr(alpha ? alpha : clr.append_child("a:alpha"), "val", rounded((100.0 - transparency.value) * 1000.0));
  return ok();
}

Expected<void, Error> apply_shape(xml_node owner, Order owner_order, const ChartFillPatch& fill,
                                  const ChartLinePatch& line) {
  if (fill_empty(fill) && line_empty(line)) {
    return ok();
  }
  const bool fresh = !child_local(owner, "spPr");
  xml_node sp = ensure(owner, "c:spPr", owner_order);
  if (fresh) {
    sp.append_child("a:effectLst");  // Excel writes one in every shape-properties element.
  }
  RETURN_IF_ERROR(apply_paint(sp, kShapeOrder, fill.color, fill.transparency_pct));
  if (line_empty(line)) {
    return ok();
  }
  xml_node ln = ensure(sp, "a:ln", kShapeOrder);
  if (line.none.has_value()) {
    if (*line.none && line.color.is_set()) {
      return bad_arg("chart: a hidden line takes no colour");
    }
    if (*line.none) {
      for (Order f = kFills; *f != nullptr; ++f) {
        drop(ln, *f);
      }
      ensure(ln, "a:noFill", kLineOrder);
    } else {
      drop(ln, "noFill");
    }
  }
  RETURN_IF_ERROR(apply_paint(ln, kLineOrder, line.color, line.transparency_pct));
  if (line.width_pt.is_clear()) {
    ln.remove_attribute("w");
  } else if (line.width_pt.is_set()) {
    if (!in_range(line.width_pt.value, 0.0, 1584.0)) {
      return bad_arg("chart: line width must be 0..1584 pt");
    }
    pugi::xml_attribute w = ln.attribute("w");
    (w ? w : ln.prepend_attribute("w")).set_value(rounded(line.width_pt.value * 12700.0).c_str());
  }
  if (line.dash.is_clear()) {
    drop(ln, "prstDash");
  } else if (line.dash.is_set()) {
    static const char* const kDash[] = {"solid", "dash", "dot", "dashDot"};
    const int d = static_cast<int>(line.dash.value);
    if (d < 0 || d > 3) {
      return bad_arg("chart: line dash must be solid, dash, dot or dashDot");
    }
    drop(ln, "custDash");
    put_val(ln, "a:prstDash", kLineOrder, kDash[d]);
  }
  return ok();
}

Expected<void, Error> apply_marker(xml_node owner, Order order, const ChartPatchField<ChartMarkerPatch>& patch) {
  if (patch.is_clear()) {
    drop(owner, "marker");
  }
  if (!patch.is_set()) {
    return ok();
  }
  const ChartMarkerPatch& m = patch.value;
  xml_node node = ensure(owner, "c:marker", order);
  if (m.symbol.has_value()) {
    static const char* const kSymbols[] = {"none", "auto", "circle", "square", "diamond", "triangle"};
    const int s = static_cast<int>(*m.symbol);
    if (s < 0 || s > 5) {
      return bad_arg("chart: marker symbol must be a named shape");
    }
    put_val(node, "c:symbol", kMarkerOrder, kSymbols[s]);
  }
  if (m.size.is_clear()) {
    drop(node, "size");
  } else if (m.size.is_set()) {
    if (!in_range(m.size.value, 2.0, 72.0)) {
      return bad_arg("chart: marker size must be 2..72");
    }
    put_val(node, "c:size", kMarkerOrder, rounded(m.size.value));
  }
  return apply_shape(node, kMarkerOrder, m.fill, m.line);
}

/// Completes a series marker with a named symbol to the form Excel re-saves unchanged: size 5 when it has
/// none, and with a fill and an outline, a 0.75 pt outline width and the series' accent colour.
void settle_marker(xml_node ser) {
  xml_node m = child_local(ser, "marker");
  const std::string_view symbol = child_local(m, "symbol").attribute("val").value();
  if (symbol.empty() || symbol == "none" || symbol == "auto") {
    return;
  }
  if (!child_local(m, "size")) {
    put_val(m, "c:size", kMarkerOrder, "5");
  }
  const xml_node sp = child_local(m, "spPr");
  xml_node ln = child_local(sp, "ln");
  const auto filled = [](const xml_node& n) {
    return std::any_of(std::begin(kFills), std::end(kFills) - 1,
                       [&](const char* f) { return static_cast<bool>(child_local(n, f)); });
  };
  if (!ln || !filled(sp)) {
    return;
  }
  if (!ln.attribute("w")) {
    ln.prepend_attribute("w").set_value("9525");
  }
  if (!filled(ln)) {
    set_attr(ensure(ln, "a:solidFill", kLineOrder).append_child("a:schemeClr"), "val",
             "accent" + std::to_string(val_u(ser, "idx") % 6U + 1U));
  }
}

/// The series `c:dLbls` Excel writes when labels are turned on.
std::string labels_xml(bool pie, const std::string& lang) {
  std::string x =
      "<c:dLbls><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr><c:txPr><a:bodyPr rot=\"0\" "
      "spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" lIns=\"38100\" "
      "tIns=\"19050\" rIns=\"38100\" bIns=\"19050\" anchor=\"ctr\" anchorCtr=\"1\"><a:spAutoFit/></a:bodyPr>"
      "<a:lstStyle/><a:p><a:pPr><a:defRPr sz=\"900\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" "
      "baseline=\"0\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"75000\"/><a:lumOff val=\"25000\"/>"
      "</a:schemeClr></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/>"
      "<a:cs typeface=\"+mn-cs\"/></a:defRPr></a:pPr><a:endParaRPr lang=\"" +
      lang +
      "\"/></a:p></c:txPr><c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/><c:showCatName val=\"0\"/>"
      "<c:showSerName val=\"0\"/><c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/>";
  const std::string ext = std::string("<c:extLst><c:ext uri=\"") + kUriLeader + "\" xmlns:c15=\"" + kNsC15 + "\"";
  if (pie) {
    x += std::string("<c:showLeaderLines val=\"1\"/><c:leaderLines>") + kLeaderSpPr + "</c:leaderLines>" + ext +
         "/></c:extLst>";
  } else {
    x += "<c:showLeaderLines val=\"0\"/>" + ext + "><c15:showLeaderLines val=\"1\"/><c15:leaderLines>" + kLeaderSpPr +
         "</c15:leaderLines></c:ext></c:extLst>";
  }
  return x + "</c:dLbls>";
}

Expected<void, Error> apply_labels(xml_node ser, bool pie, const std::string& lang,
                                   const ChartPatchField<ChartLabelsPatch>& patch) {
  if (patch.is_clear()) {
    drop(ser, "dLbls");
  }
  if (!patch.is_set()) {
    return ok();
  }
  const ChartLabelsPatch& p = patch.value;
  pugi::xml_document frag;
  const xml_node defaults = fragment(frag, labels_xml(pie, lang));
  xml_node d = child_local(ser, "dLbls");
  if (!d) {
    d = copy_at(ser, defaults, kSerOrder);
  }
  drop(d, "delete");
  static const char* const kShow[] = {"c:showLegendKey", "c:showVal",     "c:showCatName",
                                      "c:showSerName",   "c:showPercent", "c:showBubbleSize"};
  const std::optional<bool>* flags[] = {&p.show_legend_key, &p.show_val, &p.show_cat, &p.show_ser, &p.show_pct};
  for (int i = 0; i < 6; ++i) {
    if (i < 5 && flags[i]->has_value()) {
      put_val(d, kShow[i], kLabelsOrder, **flags[i] ? "1" : "0");
    } else if (!child_local(d, local_of(kShow[i]))) {
      put_val(d, kShow[i], kLabelsOrder, "0");
    }
  }
  if (p.separator.is_clear()) {
    drop(d, "separator");
  } else if (p.separator.is_set()) {
    ensure(d, "c:separator", kLabelsOrder).text().set(p.separator.value.c_str());
  }
  if (p.position.is_clear()) {
    drop(d, "dLblPos");
  } else if (p.position.is_set()) {
    static const char* const kPos[] = {"bestFit", "ctr", "inBase", "inEnd", "outEnd", "t", "b", "l", "r"};
    const int i = static_cast<int>(p.position.value);
    if (i < 0 || i > 8) {
      return bad_arg("chart: unknown label position");
    }
    put_val(d, "c:dLblPos", kLabelsOrder, kPos[i]);
  }
  if (p.number_format.is_clear()) {
    drop(d, "numFmt");
  } else if (p.number_format.is_set()) {
    xml_node fmt = ensure(d, "c:numFmt", kLabelsOrder);
    set_attr(fmt, "formatCode", p.number_format.value);
    set_attr(fmt, "sourceLinked", "0");
  }
  if (p.font_color.is_keep() && p.font_size_pt.is_keep()) {
    return ok();
  }
  xml_node tx = child_local(d, "txPr");
  if (!tx) {
    tx = copy_at(d, child_local(defaults, "txPr"), kLabelsOrder);
  }
  xml_node para = ensure(tx, "a:p", kTextOrder);
  xml_node ppr = child_local(para, "pPr") ? child_local(para, "pPr") : para.prepend_child("a:pPr");
  xml_node rpr = child_local(ppr, "defRPr") ? child_local(ppr, "defRPr") : ppr.append_child("a:defRPr");
  if (p.font_size_pt.is_clear()) {
    rpr.remove_attribute("sz");
  } else if (p.font_size_pt.is_set()) {
    if (!in_range(p.font_size_pt.value, 1.0, 4000.0)) {
      return bad_arg("chart: label font size must be 1..4000 pt");
    }
    set_attr(rpr, "sz", rounded(p.font_size_pt.value * 100.0));
  }
  return apply_paint(rpr, kRunOrder, p.font_color, ChartPatchField<double>{});
}

// ---------------------------------------------------------------------------
// References and filters
// ---------------------------------------------------------------------------

/// Replaces `ser`'s `tag` input with a reference to `formula` and an empty cache.
void put_ref(xml_node ser, const char* tag, const std::string& formula, bool multi, bool filtered) {
  drop(ser, local_of(tag));
  const bool name = local_of(tag) == "tx";
  const char* ref = name ? "strRef" : multi ? "multiLvlStrRef" : "numRef";
  std::string x = std::string("<") + tag + "><c:" + ref + ">";
  std::string esc;
  io::AppendXmlEscaped(esc, formula);
  x += filtered ? std::string("<c:extLst><c:ext uri=\"") + kUriC15 + "\"><c15:formulaRef><c15:sqref>" + esc +
                      "</c15:sqref></c15:formulaRef></c:ext></c:extLst>"
                : "<c:f>" + esc + "</c:f>";
  x += name    ? "<c:strCache><c:ptCount val=\"0\"/></c:strCache>"
       : multi ? "<c:multiLvlStrCache><c:ptCount val=\"0\"/></c:multiLvlStrCache>"
               : "<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"0\"/></c:numCache>";
  x += std::string("</c:") + ref + "></" + tag + ">";
  pugi::xml_document frag;
  copy_at(ser, fragment(frag, x), kSerOrder);
}

/// Moves a ref's `c:f` into `c15:formulaRef` (a series entering the filtered group).
void hide_ref(xml_node ref) {
  const xml_node f = child_local(ref, "f");
  if (!f) {
    return;
  }
  const std::string text = f.child_value();
  ref.remove_child(f);
  c15_ext(ref, true, false).append_child("c15:formulaRef").append_child("c15:sqref").text().set(text.c_str());
}

void show_ref(xml_node ref) {
  xml_node ext = c15_ext(ref, false, false);
  const xml_node fr = child_local(ext, "formulaRef");
  if (!fr) {
    return;
  }
  const std::string text = child_local(fr, "sqref").child_value();
  ext.remove_child(fr);
  prune_ext(ref);
  const xml_node list = child_local(ref, "extLst");
  (list ? ref.insert_child_after("c:f", list) : ref.prepend_child("c:f")).text().set(text.c_str());
}

constexpr const char* kRoles[] = {"tx", "cat", "val", "xVal", "yVal", "bubbleSize"};

void each_ref(const xml_node& ser, void (*fn)(xml_node)) {
  for (const char* role : kRoles) {
    if (const xml_node ref = ref_of(child_local(ser, role))) {
      fn(ref);
    }
  }
}

std::string filtered_group_name(std::string_view plot) {
  static const char* const kKinds[][2] = {{"bar", "Bar"},     {"line", "Line"},       {"stock", "Line"},
                                          {"area", "Area"},   {"pie", "Pie"},         {"doughnut", "Pie"},
                                          {"ofPie", "Pie"},   {"scatter", "Scatter"}, {"bubble", "Bubble"},
                                          {"radar", "Radar"}, {"surface", "Surface"}};
  for (const auto& k : kKinds) {
    if (plot.substr(0, std::strlen(k[0])) == k[0]) {
      return std::string("c15:filtered") + k[1] + "Series";
    }
  }
  return "c15:filteredBarSeries";
}

void drop_empty_group(xml_node plot) {
  xml_node group = filter_group(plot);
  if (!group || first_element(group)) {
    return;
  }
  xml_node ext = group.parent();
  ext.remove_child(group);
  if (!first_element(ext)) {
    xml_node list = ext.parent();
    list.remove_child(ext);
    if (!first_element(list)) {
      plot.remove_child(list);
    }
  }
}

/// Moves series between the plot and its `c15:filtered*Series` group, in the form Excel saves.
void patch_series_filter(const Chart& c, const std::set<std::uint32_t>& hidden) {
  for (xml_node plot : plots_of(c)) {
    for (xml_node ser : kids(plot, "ser")) {
      if (hidden.count(val_u(ser, "idx")) == 0) {
        continue;
      }
      xml_node group = filter_group(plot);
      if (!group) {
        xml_node ext = ensure(plot, "c:extLst", kPlotOrder).append_child("c:ext");
        ext.append_attribute("uri").set_value(kUriC15);
        ext.append_attribute("xmlns:c15").set_value(kNsC15);
        group = ext.append_child(filtered_group_name(local_name(plot)).c_str());
      }
      each_ref(ser, hide_ref);
      group.append_move(ser).set_name("c15:ser");
    }
    for (xml_node ser : kids(filter_group(plot), "ser")) {
      if (hidden.count(val_u(ser, "idx")) != 0) {
        continue;
      }
      each_ref(ser, show_ref);
      const std::uint32_t order = val_u(ser, "order");
      xml_node before;
      for (const xml_node& s : kids(plot, "ser")) {
        if (val_u(s, "order") > order) {
          before = s;
          break;
        }
      }
      if (!before) {
        before = slot(plot, "ser", kPlotOrder);
      }
      (before ? plot.insert_move_before(ser, before) : plot.append_move(ser)).set_name("c:ser");
    }
    drop_empty_group(plot);
  }
}

/// Points each series' category / value refs at the visible cells, keeping
/// the full range in `c15:fullRef`, as Excel saves it. Positions index the full range.
Expected<void, Error> patch_category_filter(const Workbook& wb, const Chart& c, const std::set<std::uint32_t>& hidden) {
  struct Cell {
    std::size_t area;
    std::size_t sheet;
    std::uint32_t row;
    std::uint32_t col;
  };
  for (xml_node plot : plots_of(c)) {
    for (xml_node ser : plot_series(plot)) {
      for (const char* role : {"cat", "val", "xVal", "yVal"}) {
        xml_node ref = ref_of(child_local(ser, role));
        xml_node text = formula_node(ref);
        if (!text) {
          continue;
        }
        xml_node ext = c15_ext(ref, false, false);
        xml_node full = child_local(child_local(ext, "fullRef"), "sqref");
        const std::string full_text = full ? full.child_value() : text.child_value();
        if (hidden.empty()) {
          if (full) {
            text.text().set(full_text.c_str());
            ext.remove_child(full.parent());
            prune_ext(ref);
          }
          continue;
        }
        ASSIGN_OR_RETURN(ResolvedSource src, resolve_chart_source(wb, c.host, full_text, eval::default_registry()));
        std::vector<Cell> cells;
        for (std::size_t a = 0; a < src.areas.size(); ++a) {
          const SheetRect& r = src.areas[a];
          if (r.top != r.bottom && r.left != r.right) {
            return unsupported("chart: a two-dimensional reference cannot be filtered");
          }
          for (std::uint32_t col = r.left; col <= r.right; ++col) {
            for (std::uint32_t row = r.top; row <= r.bottom; ++row) {
              cells.push_back(Cell{a, r.sheet, row, col});
            }
          }
        }
        if (*hidden.rbegin() >= cells.size() || hidden.size() >= cells.size()) {
          return bad_arg("chart: hidden categories must leave one of the chart's categories");
        }
        std::vector<std::string> runs;
        for (std::size_t i = 0; i < cells.size();) {
          if (hidden.count(static_cast<std::uint32_t>(i)) != 0) {
            ++i;
            continue;
          }
          std::size_t j = i;
          while (j + 1 < cells.size() && hidden.count(static_cast<std::uint32_t>(j + 1)) == 0 &&
                 cells[j + 1].area == cells[i].area) {
            ++j;
          }
          runs.push_back(
              chart_rect_ref(wb.sheet(cells[i].sheet).name(), cells[i].row, cells[i].col, cells[j].row, cells[j].col));
          i = j + 1;
        }
        std::string joined;
        for (const std::string& run : runs) {
          joined += (joined.empty() ? "" : ",") + run;
        }
        text.text().set((runs.size() > 1 ? "(" + joined + ")" : joined).c_str());
        if (!full) {
          ext = c15_ext(ref, true, true);
          full = ext.prepend_child("c15:fullRef").append_child("c15:sqref");
          full.text().set(full_text.c_str());
        }
      }
    }
  }
  return ok();
}

std::set<std::uint32_t> filtered_series(const Chart& c) {
  std::set<std::uint32_t> out;
  for (const xml_node& plot : plots_of(c)) {
    for (const xml_node& s : kids(filter_group(plot), "ser")) {
      out.insert(val_u(s, "idx"));
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Patches (shared by the edits and `preview_chart_data`)
// ---------------------------------------------------------------------------

Expected<void, Error> patch_series(const Workbook& wb, const Chart& c, xml_node plot, xml_node ser,
                                   const ChartSeriesPatch& p, bool* refs) {
  const std::string_view pl = local_name(plot);
  const bool xy = pl == "scatterChart" || pl == "bubbleChart";
  struct Field {
    const std::optional<std::string>* text;
    const char* tag;
    bool required;
    bool allowed;
  };
  const Field fields[] = {{&p.name, "c:tx", false, true}, {&p.cat, "c:cat", false, !xy},
                          {&p.val, "c:val", true, !xy},   {&p.x, "c:xVal", false, xy},
                          {&p.y, "c:yVal", true, xy},     {&p.size, "c:bubbleSize", true, pl == "bubbleChart"}};
  for (const Field& f : fields) {
    if (!f.text->has_value()) {
      continue;
    }
    if (c.pivot()) {
      return unsupported("chart: a pivot chart's series follow its pivot table");
    }
    if (!f.allowed) {
      return bad_arg("chart: the plot has no such series input");
    }
    *refs = true;
    const std::string& text = **f.text;
    if (text.empty()) {
      if (f.required) {
        return bad_arg("chart: a series needs its values");
      }
      drop(ser, local_of(f.tag));
      continue;
    }
    ASSIGN_OR_RETURN(ResolvedSource src, resolve_chart_source(wb, c.host, text, eval::default_registry()));
    const bool multi = (f.tag == fields[1].tag || f.tag == fields[3].tag) && src.areas.size() == 1 &&
                       src.areas[0].top != src.areas[0].bottom && src.areas[0].left != src.areas[0].right;
    put_ref(ser, f.tag, text, multi, ser.parent() != plot);
  }
  if ((!p.marker.is_keep() || p.smooth.has_value()) && !has_markers(plot)) {
    return bad_arg("chart: markers and smoothing need a line, scatter or radar plot");
  }
  RETURN_IF_ERROR(apply_shape(ser, kSerOrder, marker_filled(plot) ? ChartFillPatch{} : p.fill, p.line));
  RETURN_IF_ERROR(apply_marker(ser, kSerOrder, p.marker));
  if (p.marker.is_set()) {
    settle_marker(ser);
  }
  if (p.smooth.has_value()) {
    put_val(ser, "c:smooth", kSerOrder, *p.smooth ? "1" : "0");
  }
  return apply_labels(ser, is_pie(plot), c.lang(), p.labels);
}

std::string accent_sppr(std::uint32_t i) {
  return "<c:spPr><a:solidFill><a:schemeClr val=\"accent" + std::to_string(i % 6U + 1U) +
         "\"/></a:solidFill><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>";
}

/// Inserts a copy of `src` among `ser`'s `c:dPt` in idx order.
xml_node insert_point(xml_node ser, const xml_node& src, std::uint32_t idx) {
  for (const xml_node& d : kids(ser, "dPt")) {
    if (val_u(d, "idx") > idx) {
      return ser.insert_copy_before(src, d);
    }
  }
  return copy_at(ser, src, kSerOrder);
}

xml_node find_point(const xml_node& ser, std::uint32_t idx) {
  for (const xml_node& d : kids(ser, "dPt")) {
    if (val_u(d, "idx") == idx) {
      return d;
    }
  }
  return {};
}

/// Excel's column vary-colours form: per-point accent fills replace the series fill.
void vary_bar(xml_node plot, bool on) {
  for (xml_node ser : plot_series(plot)) {
    if (on) {
      drop(ser, "spPr");
      const std::uint32_t n = point_count(ser);
      for (std::uint32_t p = 0; p < n; ++p) {
        if (find_point(ser, p)) {
          continue;
        }
        pugi::xml_document frag;
        insert_point(ser,
                     fragment(frag, "<c:dPt><c:idx val=\"" + std::to_string(p) +
                                        "\"/><c:invertIfNegative val=\"0\"/><c:bubble3D val=\"0\"/>" + accent_sppr(p) +
                                        "</c:dPt>"),
                     p);
      }
      continue;
    }
    for (const xml_node& d : kids(ser, "dPt")) {
      const std::string_view clr =
          child_local(child_local(child_local(d, "spPr"), "solidFill"), "schemeClr").attribute("val").value();
      if (clr.substr(0, 6) == "accent") {
        ser.remove_child(d);
      }
    }
    if (!child_local(ser, "spPr")) {
      pugi::xml_document frag;
      copy_at(ser, fragment(frag, accent_sppr(val_u(ser, "idx"))), kSerOrder);
    }
  }
}

void patch_display(const Chart& c, std::optional<ChartBlanksAs> blanks, std::optional<Tri> vis) {
  xml_node chart = c.chart();
  if (blanks.has_value()) {
    static const char* const kBlanks[] = {"gap", "zero", "span"};
    const int b = static_cast<int>(*blanks);
    if (b >= 1 && b <= 3) {
      put_val(chart, "c:dispBlanksAs", kChartOrder, kBlanks[b - 1]);
    } else {
      drop(chart, "dispBlanksAs");
    }
  }
  if (vis.has_value()) {
    if (*vis == Tri::kAbsent) {
      drop(chart, "plotVisOnly");
    } else {
      put_val(chart, "c:plotVisOnly", kChartOrder, *vis == Tri::kTrue ? "1" : "0");
    }
  }
}

/// Rebuilds every series from `source`; series keep their elements by idx.
Expected<void, Error> patch_source(const Workbook& wb, Chart& c, std::string_view source,
                                   ChartSeriesDirection direction, ChartHeaderMode header_row,
                                   ChartHeaderMode header_column) {
  const std::vector<xml_node> plots = plots_of(c);
  const int kind = plots.size() == 1 ? kind_of(plots[0]) : -1;
  if (kind < 0) {
    return unsupported(
        "chart: only a single column, bar, line, area, pie, doughnut or scatter plot takes a data range");
  }
  xml_node plot = plots[0];
  ASSIGN_OR_RETURN(ChartSplit split, split_chart_source(wb, c.host, source, static_cast<ChartInsertKind>(kind),
                                                        direction, header_row, header_column));
  pugi::xml_document t;
  RETURN_IF_ERROR(template_doc(wb, static_cast<ChartInsertKind>(kind), grouping_of(plot), std::nullopt,
                               ChartLegendPosition::kRight, split, t));
  patch_series_filter(c, {});
  RETURN_IF_ERROR(patch_category_filter(wb, c, {}));
  const std::vector<xml_node> fresh = kids(template_plot(t), "ser");
  for (const xml_node& ser : kids(plot, "ser")) {
    if (val_u(ser, "idx") >= fresh.size()) {
      plot.remove_child(ser);
    }
  }
  for (std::uint32_t i = 0; i < fresh.size(); ++i) {
    xml_node ser;
    for (const xml_node& s : kids(plot, "ser")) {
      if (val_u(s, "idx") == i) {
        ser = s;
      }
    }
    if (!ser) {
      copy_at(plot, fresh[i], kPlotOrder);
      continue;
    }
    for (const char* role : kRoles) {
      drop(ser, role);
      if (const xml_node from = child_local(fresh[i], role)) {
        copy_at(ser, from, kSerOrder);
      }
    }
    put_val(ser, "c:order", kSerOrder, std::to_string(i));
    for (const xml_node& d : kids(fresh[i], "dPt")) {
      if (!find_point(ser, val_u(d, "idx"))) {
        insert_point(ser, d, val_u(d, "idx"));
      }
    }
  }
  return ok();
}

enum class Family : std::uint8_t { kNone, kFill, kLine, kMarker };

Family family_of(int kind) {
  switch (static_cast<ChartInsertKind>(kind)) {
    case ChartInsertKind::kLine:
      return Family::kLine;
    case ChartInsertKind::kScatter:
      return Family::kMarker;
    case ChartInsertKind::kPie:
    case ChartInsertKind::kDoughnut:
      return Family::kNone;
    default:
      return Family::kFill;
  }
}

/// The explicit RGB colour a series is drawn in, by family.
std::optional<std::uint32_t> main_color(const xml_node& ser, Family f) {
  const xml_node sp = f == Family::kMarker ? child_local(child_local(ser, "marker"), "spPr") : child_local(ser, "spPr");
  const xml_node fill =
      f == Family::kLine ? child_local(child_local(sp, "ln"), "solidFill") : child_local(sp, "solidFill");
  const xml_node clr = child_local(fill, "srgbClr");
  if (f == Family::kNone || !clr) {
    return std::nullopt;
  }
  return io::parse_rgb_hex(clr.attribute("val").value(), 0U) | 0xFF000000U;
}

void set_main_color(xml_node ser, Family f, std::uint32_t argb) {
  ChartFillPatch fill;
  ChartLinePatch line;
  fill.color = ChartPatchField<std::uint32_t>::Set(argb);
  line.color = fill.color;
  if (f == Family::kFill) {
    static_cast<void>(apply_shape(ser, kSerOrder, fill, ChartLinePatch{}));
  } else if (f == Family::kLine) {
    static_cast<void>(apply_shape(ser, kSerOrder, ChartFillPatch{}, line));
  } else if (f == Family::kMarker) {
    ChartMarkerPatch m;
    m.fill = fill;
    m.line = line;
    static_cast<void>(apply_marker(ser, kSerOrder, ChartPatchField<ChartMarkerPatch>::Set(m)));
  }
}

/// Replaces the `a:ln` of shape properties `sp` with a copy of `ln`.
void replace_ln(xml_node sp, const xml_node& ln) {
  if (!sp || !ln) {
    return;
  }
  drop(sp, "ln");
  copy_at(sp, ln, kShapeOrder);
}

std::vector<xml_node> axes_of(const Chart& c, const std::vector<std::uint32_t>& ids) {
  std::vector<xml_node> out;
  for (const std::uint32_t id : ids) {
    for (xml_node n = c.area().first_child(); n; n = n.next_sibling()) {
      const std::string_view name = local_name(n);
      if (name.size() > 2 && name.substr(name.size() - 2) == "Ax" && val_u(n, "axId") == id) {
        out.push_back(n);
        break;
      }
    }
  }
  return out;
}

/// Puts `ax` on the side Excel derives on save: the far side (t / r) when it crosses at the maximum of its
/// crossing axis, with the sides swapped while that axis is reversed.
void place_axis(const Chart& c, xml_node ax) {
  const std::string_view pos = child_local(ax, "axPos").attribute("val").value();
  const bool horizontal = pos == "b" || pos == "t";
  if (!horizontal && pos != "l" && pos != "r") {
    return;
  }
  const std::vector<xml_node> cross = axes_of(c, {val_u(ax, "crossAx")});
  const bool reversed =
      !cross.empty() &&
      std::string_view(child_local(child_local(cross[0], "scaling"), "orientation").attribute("val").value()) ==
          "maxMin";
  const bool far = (std::string_view(child_local(ax, "crosses").attribute("val").value()) == "max") != reversed;
  put_val(ax, "c:axPos", kAxisOrder, horizontal ? (far ? "t" : "b") : (far ? "r" : "l"));
}

Expected<void, Error> patch_type(const Workbook& wb, Chart& c, std::size_t plot_index, ChartInsertKind kind,
                                 ChartGrouping grouping) {
  const int k = static_cast<int>(kind);
  if (k < 0 || k > 6 || static_cast<int>(grouping) < 0 || static_cast<int>(grouping) > 3) {
    return bad_arg("chart: unknown chart kind or grouping");
  }
  ASSIGN_OR_RETURN(xml_node plot, plot_at(c, plot_index));
  const int old = kind_of(plot);
  if (old < 0) {
    return unsupported("chart: only column, bar, line, area, pie, doughnut and scatter plots change type");
  }
  const bool scatter = kind == ChartInsertKind::kScatter;
  const auto ring = [](int x) {
    return x == static_cast<int>(ChartInsertKind::kPie) || x == static_cast<int>(ChartInsertKind::kDoughnut);
  };
  const bool old_ring = ring(old);
  const bool new_ring = ring(k);
  if (c.pivot() && (scatter || old == static_cast<int>(ChartInsertKind::kScatter))) {
    return unsupported("chart: a pivot chart cannot be a scatter chart");
  }
  std::vector<std::uint32_t> ids;
  for (const xml_node& a : kids(plot, "axId")) {
    ids.push_back(a.attribute("val").as_uint());
  }
  for (const xml_node& other : plots_of(c)) {
    for (const xml_node& a : kids(other, "axId")) {
      if (other != plot && std::count(ids.begin(), ids.end(), a.attribute("val").as_uint()) != 0) {
        return unsupported("chart: the plot shares its axes with another plot");
      }
    }
  }
  const auto norm = [](ChartGrouping g) { return g == ChartGrouping::kStandard ? ChartGrouping::kClustered : g; };
  if (old == k && (k > static_cast<int>(ChartInsertKind::kArea) || norm(grouping_of(plot)) == norm(grouping))) {
    return ok();
  }

  const std::set<std::uint32_t> hidden = filtered_series(c);
  patch_series_filter(c, {});
  const std::vector<xml_node> sers = kids(plot, "ser");
  ChartSplit split;
  for (const xml_node& s : sers) {
    const xml_node cat = child_local(s, "cat") ? child_local(s, "cat") : child_local(s, "xVal");
    const xml_node val = child_local(s, "val") ? child_local(s, "val") : child_local(s, "yVal");
    ChartSeriesRefs refs;
    refs.name = formula_of(child_local(s, "tx"));
    (scatter ? refs.x : refs.cat) = formula_of(cat);
    (scatter ? refs.y : refs.val) = formula_of(val);
    if (formula_of(val).empty()) {
      return unsupported("chart: a series without a value reference cannot change type");
    }
    split.series.push_back(std::move(refs));
  }
  pugi::xml_document t;
  RETURN_IF_ERROR(template_doc(wb, kind, grouping, std::nullopt, ChartLegendPosition::kRight, split, t));
  xml_node tplot = template_plot(t);
  const std::vector<xml_node> fresh = kids(tplot, "ser");
  for (std::size_t i = 0; i < sers.size() && i < fresh.size(); ++i) {
    xml_node to = fresh[i];
    const xml_node from = sers[i];
    put_val(to, "c:idx", kSerOrder, std::to_string(val_u(from, "idx")));
    put_val(to, "c:order", kSerOrder, std::to_string(val_u(from, "order")));
    for (const char* role : kRoles) {
      drop(to, role);
    }
    drop(to, "extLst");
    if (const xml_node tx = child_local(from, "tx")) {
      copy_at(to, tx, kSerOrder);
    }
    const xml_node cat = child_local(from, "cat") ? child_local(from, "cat") : child_local(from, "xVal");
    const xml_node val = child_local(from, "val") ? child_local(from, "val") : child_local(from, "yVal");
    if (cat) {
      copy_at(to, cat, kSerOrder, scatter ? "c:xVal" : "c:cat");
    }
    copy_at(to, val, kSerOrder, scatter ? "c:yVal" : "c:val");
    if (const xml_node ext = child_local(from, "extLst")) {
      copy_at(to, ext, kSerOrder);
    }
    if (const auto color = main_color(from, family_of(old))) {
      set_main_color(to, family_of(k), *color);
    }
    // As Excel changes type: line formatting follows the colours between series and points; pie and doughnut swap
    // points as-is.
    if (old_ring && new_ring) {
      drop(to, "dPt");
      for (const xml_node& d : kids(from, "dPt")) {
        copy_at(to, d, kSerOrder);
      }
    } else if (old_ring && family_of(k) == Family::kFill) {
      replace_ln(child_local(to, "spPr"), child_local(child_local(child_local(from, "dPt"), "spPr"), "ln"));
    } else if (family_of(old) == Family::kFill && new_ring) {
      for (const xml_node& d : kids(to, "dPt")) {
        replace_ln(child_local(d, "spPr"), child_local(child_local(from, "spPr"), "ln"));
      }
    }
  }

  // Axes: kept when the axis kinds match, else replaced by the template's.
  const std::vector<xml_node> old_axes = axes_of(c, ids);
  std::vector<xml_node> new_axes;
  for (xml_node n = child_local(template_chart(t), "plotArea").first_child(); n; n = n.next_sibling()) {
    if (local_name(n) == "catAx" || local_name(n) == "valAx") {
      new_axes.push_back(n);
    }
  }
  bool same = old_axes.size() == new_axes.size() && !old_axes.empty();
  for (std::size_t i = 0; same && i < old_axes.size(); ++i) {
    same = local_name(old_axes[i]) == local_name(new_axes[i]);
  }
  xml_node area = c.area();
  std::vector<std::uint32_t> new_ids;
  if (same) {
    for (std::size_t i = 0; i < old_axes.size(); ++i) {
      put_val(old_axes[i], "c:axPos", kAxisOrder, child_local(new_axes[i], "axPos").attribute("val").value());
      if (const xml_node between = child_local(new_axes[i], "crossBetween")) {
        put_val(old_axes[i], "c:crossBetween", kAxisOrder, between.attribute("val").value());
      }
      // A linked value-axis format follows the grouping, as Excel re-derives it: 0% for percent-stacked.
      xml_node fmt = child_local(old_axes[i], "numFmt");
      if (local_name(old_axes[i]) == "valAx" && fmt && flag_attr(fmt, "sourceLinked") &&
          (grouping_of(plot) == ChartGrouping::kPercentStacked) != (grouping == ChartGrouping::kPercentStacked)) {
        set_attr(fmt, "formatCode", child_local(new_axes[i], "numFmt").attribute("formatCode").value());
      }
      new_ids.push_back(val_u(old_axes[i], "axId"));
    }
    for (const xml_node& a : old_axes) {
      place_axis(c, a);
    }
  } else {
    for (const xml_node& a : old_axes) {
      area.remove_child(a);
    }
    std::set<std::uint32_t> used;
    for (xml_node n = area.first_child(); n; n = n.next_sibling()) {
      if (child_local(n, "axId") && !is_plot(n)) {
        used.insert(val_u(n, "axId"));
      }
    }
    for (const xml_node& a : new_axes) {
      std::uint32_t id = val_u(a, "axId");
      while (used.count(id) != 0) {
        ++id;
      }
      used.insert(id);
      new_ids.push_back(id);
    }
    xml_node before;
    for (xml_node n = area.first_child(); n && !before; n = n.next_sibling()) {
      const std::string_view name = local_name(n);
      before = name == "dTable" || name == "spPr" || name == "extLst" ? n : xml_node();
    }
    for (std::size_t i = 0; i < new_axes.size(); ++i) {
      xml_node a = before ? area.insert_copy_before(new_axes[i], before) : area.append_copy(new_axes[i]);
      put_val(a, "c:axId", kAxisOrder, std::to_string(new_ids[i]));
      put_val(a, "c:crossAx", kAxisOrder, std::to_string(new_ids[1 - i]));
    }
  }
  const std::vector<xml_node> axis_refs = kids(tplot, "axId");
  for (std::size_t i = 0; i < axis_refs.size() && i < new_ids.size(); ++i) {
    set_attr(axis_refs[i], "val", std::to_string(new_ids[i]));
  }
  // As Excel changes type: a new bar plot takes the schema default gap; a bar plot keeps its own.
  if (is_bar(tplot)) {
    const xml_node gap = child_local(plot, "gapWidth");
    put_val(tplot, "c:gapWidth", kPlotOrder, is_bar(plot) && gap ? gap.attribute("val").value() : "150");
  }
  if (old_ring && new_ring && child_local(plot, "firstSliceAng")) {
    put_val(tplot, "c:firstSliceAng", kPlotOrder, child_local(plot, "firstSliceAng").attribute("val").value());
  }
  area.insert_copy_before(tplot, plot);
  area.remove_child(plot);
  patch_series_filter(c, hidden);
  // As Excel changes type: a pie with no title that was not deleted takes the automatic title.
  xml_node chart = c.chart();
  const xml_node deleted = child_local(chart, "autoTitleDeleted");
  if (kind == ChartInsertKind::kPie && !child_local(chart, "title") && deleted && !flag(deleted)) {
    copy_at(chart, child_local(template_chart(t), "title"), kChartOrder);
  }
  return ok();
}

Expected<void, Error> patch_title(const Workbook& wb, const Chart& c, const ChartTitle& title) {
  xml_node chart = c.chart();
  if (title.kind == ChartTitleKind::kNone) {
    drop(chart, "title");
    put_val(chart, "c:autoTitleDeleted", kChartOrder, "1");
    return ok();
  }
  if (title.kind == ChartTitleKind::kRef) {
    if (title.ref_formula.empty()) {
      return bad_arg("chart: a reference title needs a formula");
    }
    ASSIGN_OR_RETURN(ResolvedSource unused,
                     resolve_chart_source(wb, c.host, title.ref_formula, eval::default_registry()));
    static_cast<void>(unused);
  } else if (title.kind == ChartTitleKind::kText && title.text.empty()) {
    return bad_arg("chart: a text title needs text");
  } else if (title.kind != ChartTitleKind::kText && title.kind != ChartTitleKind::kAuto) {
    return bad_arg("chart: unknown title kind");
  }
  const bool text = title.kind == ChartTitleKind::kText;
  pugi::xml_document t;
  RETURN_IF_ERROR(template_doc(wb, ChartInsertKind::kColumn, ChartGrouping::kClustered,
                               text ? std::optional<std::string>(title.text) : std::nullopt,
                               ChartLegendPosition::kRight, {}, t));
  const xml_node from = child_local(template_chart(t), "title");
  xml_node node = child_local(chart, "title");
  if (!node) {
    node = copy_at(chart, from, kChartOrder);
  }
  drop(node, "tx");
  if (text) {
    copy_at(node, child_local(from, "tx"), kTitleOrder);
    drop(node, "txPr");
  } else if (!child_local(node, "txPr")) {
    copy_at(node, child_local(from, "txPr"), kTitleOrder);
  }
  if (title.kind == ChartTitleKind::kRef) {
    std::string x = "<c:tx><c:strRef><c:f>";
    io::AppendXmlEscaped(x, title.ref_formula);
    x += "</c:f></c:strRef></c:tx>";
    pugi::xml_document frag;
    copy_at(node, fragment(frag, x), kTitleOrder);
  }
  put_val(chart, "c:autoTitleDeleted", kChartOrder, "0");
  return ok();
}

Expected<ChartModel, Error> model_of(const Workbook& wb, std::size_t sheet, std::uint32_t object_id, const Chart& c,
                                     bool derive) {
  ASSIGN_OR_RETURN(ChartModel model, read_chart_model(wb, c.doc));
  model.object_id = object_id;
  model.on_chartsheet = c.host == kChartNoHostSheet;
  ASSIGN_OR_RETURN(std::vector<DrawingObject> objects, list_drawing_objects(wb, sheet));
  for (const DrawingObject& o : objects) {
    if (o.object_id == object_id) {
      model.name = o.name;
      model.descr = o.descr;
      model.anchor_kind = o.anchor_kind;
      model.edit_as = o.edit_as;
      model.from = o.from;
      model.to = o.to;
      model.cx = o.cx;
      model.cy = o.cy;
      break;
    }
  }
  if (derive) {
    model.source = derive_chart_source(wb, c.host, model);
  }
  return model;
}

Expected<ChartData, Error> data_of(const Workbook& wb, std::size_t sheet, std::uint32_t object_id, const Chart& c) {
  ASSIGN_OR_RETURN(ChartModel model, model_of(wb, sheet, object_id, c, false));
  return resolve_chart_data(wb, c.host, model);
}

}  // namespace

Expected<ChartModel, Error> get_chart(const Workbook& wb, std::size_t sheet, std::uint32_t object_id) {
  Chart c;
  RETURN_IF_ERROR(open_chart(wb, sheet, object_id, c));
  return model_of(wb, sheet, object_id, c, true);
}

Expected<ChartData, Error> get_chart_data(const Workbook& wb, std::size_t sheet, std::uint32_t object_id) {
  Chart c;
  RETURN_IF_ERROR(open_chart(wb, sheet, object_id, c));
  return data_of(wb, sheet, object_id, c);
}

Expected<ChartData, Error> preview_chart_data(const Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                              const ChartPreviewOverrides& overrides) {
  Chart c;
  RETURN_IF_ERROR(open_chart(wb, sheet, object_id, c));
  if (c.chartex) {
    return unsupported("chart: chartex charts are not editable");
  }
  const ChartPreviewOverrides& o = overrides;
  const bool resplit = o.source || o.direction || o.header_row || o.header_column;
  if (resplit) {
    if (c.pivot()) {
      return unsupported("chart: a pivot chart's series follow its pivot table");
    }
    std::optional<ChartSource> base;
    if (!o.source) {
      ASSIGN_OR_RETURN(ChartModel model, model_of(wb, sheet, object_id, c, true));
      base = model.source;
      if (!base) {
        return unsupported("chart: the series do not share one data range");
      }
    }
    const auto mode = [&](const std::optional<bool>& v, bool derived) {
      return v      ? (*v ? ChartHeaderMode::kYes : ChartHeaderMode::kNo)
             : base ? (derived ? ChartHeaderMode::kYes : ChartHeaderMode::kNo)
                    : ChartHeaderMode::kAuto;
    };
    RETURN_IF_ERROR(patch_source(wb, c, o.source ? *o.source : base->range,
                                 o.direction ? *o.direction
                                 : base      ? base->direction
                                             : ChartSeriesDirection::kAuto,
                                 mode(o.header_row, base && base->header_row),
                                 mode(o.header_column, base && base->header_column)));
  }
  std::optional<Tri> vis;
  if (o.plot_vis_only) {
    vis = *o.plot_vis_only ? Tri::kTrue : Tri::kFalse;
  }
  patch_display(c, o.blanks_as, vis);
  if (resplit) {
    refresh_chart_caches(wb, c.host, c.doc);
  }
  return data_of(wb, sheet, object_id, c);
}

Expected<std::uint32_t, Error> insert_chart(Workbook& wb, std::size_t sheet, const ChartInsertOptions& options) {
  if (sheet >= wb.sheet_count()) {
    return bad_arg("chart: sheet index out of range");
  }
  if (static_cast<int>(options.kind) < 0 || static_cast<int>(options.kind) > 6 ||
      static_cast<int>(options.grouping) < 0 || static_cast<int>(options.grouping) > 3) {
    return bad_arg("chart: unknown chart kind or grouping");
  }
  ASSIGN_OR_RETURN(ChartSplit split, split_chart_source(wb, sheet, options.source, options.kind, options.direction,
                                                        options.header_row, options.header_column));
  ASSIGN_OR_RETURN(std::string xml, build_chart_xml(wb, options, split));
  pugi::xml_document doc;
  if (!load_xml(doc, xml)) {
    return unparseable("chart: template does not parse");
  }
  refresh_chart_caches(wb, sheet, doc);
  std::string path;
  for (std::uint32_t n = 1;; ++n) {
    path = "xl/charts/chart" + std::to_string(n) + ".xml";
    if (io::ooxml::find_passthrough_part(wb, path) == nullptr) {
      break;
    }
  }
  ImageInsertOptions placement = options.placement;
  placement.name = options.name;
  placement.descr = options.descr;
  placement.width_emu = placement.width_emu == 0 ? kFrameCx : placement.width_emu;
  placement.height_emu = placement.height_emu == 0 ? kFrameCy : placement.height_emu;
  const std::string data = std::string("<a:graphicData uri=\"") + kNsC + "\"><c:chart xmlns:c=\"" + kNsC +
                           "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
                           "r:id=\"{rId}\"/></a:graphicData>";
  ASSIGN_OR_RETURN(std::uint32_t id, insert_graphic_frame(wb, sheet, placement, data, path, kRelChart));
  std::string bytes(io::kXmlDecl);
  io::append_raw_xml(bytes, doc.document_element());
  RETURN_IF_ERROR(
      wb.add_passthrough_part(PassthroughPart(path, kCtChart, std::vector<std::uint8_t>(bytes.begin(), bytes.end()))));
  return id;
}

Expected<void, Error> set_chart_type(Workbook& wb, std::size_t sheet, std::uint32_t object_id, std::size_t plot,
                                     ChartInsertKind kind, ChartGrouping grouping) {
  Chart c;
  RETURN_IF_ERROR(open_edit(wb, sheet, object_id, c, false));
  RETURN_IF_ERROR(patch_type(wb, c, plot, kind, grouping));
  return commit(wb, c, true);
}

Expected<void, Error> set_chart_title(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                      const ChartTitle& title) {
  Chart c;
  RETURN_IF_ERROR(open_edit(wb, sheet, object_id, c, false));
  RETURN_IF_ERROR(patch_title(wb, c, title));
  return commit(wb, c, false);
}

Expected<void, Error> set_chart_legend(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                       const ChartLegendPatch& legend) {
  Chart c;
  RETURN_IF_ERROR(open_edit(wb, sheet, object_id, c, false));
  xml_node chart = c.chart();
  const int pos = static_cast<int>(legend.position);
  if (pos < 0 || pos > 5) {
    return bad_arg("chart: unknown legend position");
  }
  if (legend.position == ChartLegendPosition::kNone) {
    drop(chart, "legend");
    return commit(wb, c, false);
  }
  xml_node node = child_local(chart, "legend");
  if (!node) {
    pugi::xml_document t;
    RETURN_IF_ERROR(
        template_doc(wb, ChartInsertKind::kColumn, ChartGrouping::kClustered, std::nullopt, legend.position, {}, t));
    node = copy_at(chart, child_local(template_chart(t), "legend"), kChartOrder);
  }
  static const char* const kPos[] = {"", "b", "t", "l", "r", "tr"};
  put_val(node, "c:legendPos", kLegendOrder, kPos[pos]);
  put_val(node, "c:overlay", kLegendOrder, legend.overlay ? "1" : "0");
  if (legend.hidden_series_idx.has_value()) {
    const std::set<std::uint32_t> hidden(legend.hidden_series_idx->begin(), legend.hidden_series_idx->end());
    for (const xml_node& e : kids(node, "legendEntry")) {
      if (flag(child_local(e, "delete")) || hidden.count(val_u(e, "idx")) != 0) {
        node.remove_child(e);
      }
    }
    for (const std::uint32_t idx : hidden) {
      pugi::xml_document frag;
      copy_at(node,
              fragment(frag, "<c:legendEntry><c:idx val=\"" + std::to_string(idx) +
                                 "\"/><c:delete val=\"1\"/></c:legendEntry>"),
              kLegendOrder);
    }
  }
  return commit(wb, c, false);
}

Expected<void, Error> set_chart_axis(Workbook& wb, std::size_t sheet, std::uint32_t object_id, std::uint32_t axis_id,
                                     const ChartAxisPatch& patch) {
  Chart c;
  RETURN_IF_ERROR(open_edit(wb, sheet, object_id, c, false));
  const std::vector<xml_node> found = axes_of(c, {axis_id});
  if (found.empty()) {
    return bad_arg("chart: no axis with that id");
  }
  xml_node ax = found[0];
  const std::string_view kind = local_name(ax);
  if (patch.visible.has_value()) {
    put_val(ax, "c:delete", kAxisOrder, *patch.visible ? "0" : "1");
  }
  if (patch.title.is_clear()) {
    drop(ax, "title");
  } else if (patch.title.is_set()) {
    if (patch.title.value.empty()) {
      return bad_arg("chart: an axis title needs text");
    }
    pugi::xml_document t;
    RETURN_IF_ERROR(template_doc(wb, ChartInsertKind::kColumn, ChartGrouping::kClustered, patch.title.value,
                                 ChartLegendPosition::kRight, {}, t));
    drop(ax, "title");
    xml_node title = copy_at(ax, child_local(template_chart(t), "title"), kAxisOrder);
    const xml_node rich = child_local(child_local(title, "tx"), "rich");
    for (const xml_node& p : kids(rich, "p")) {
      set_attr(child_local(child_local(p, "pPr"), "defRPr"), "sz", "1000");
    }
    const std::string_view pos = child_local(ax, "axPos").attribute("val").value();
    if (pos == "l" || pos == "r") {
      set_attr(child_local(rich, "bodyPr"), "rot", "-5400000");
    }
  }
  const auto scaling = [&] { return ensure(ax, "c:scaling", kAxisOrder); };
  const ChartPatchField<double>* bounds[] = {&patch.min, &patch.max};
  const char* bound_names[] = {"c:min", "c:max"};
  for (int i = 0; i < 2; ++i) {
    if (bounds[i]->is_clear()) {
      drop(scaling(), local_of(bound_names[i]));
    } else if (bounds[i]->is_set()) {
      if (!std::isfinite(bounds[i]->value)) {
        return bad_arg("chart: axis bounds must be finite");
      }
      put_val(scaling(), bound_names[i], kScalingOrder, num(bounds[i]->value));
    }
  }
  if (patch.reversed.has_value()) {
    put_val(scaling(), "c:orientation", kScalingOrder, *patch.reversed ? "maxMin" : "minMax");
  }
  if (patch.log_base.is_clear()) {
    drop(scaling(), "logBase");
  } else if (patch.log_base.is_set()) {
    if (!in_range(patch.log_base.value, 2.0, 1000.0)) {
      return bad_arg("chart: log base must be 2..1000");
    }
    put_val(scaling(), "c:logBase", kScalingOrder, num(patch.log_base.value));
  }
  if (!patch.major_unit.is_keep() && (kind == "catAx" || kind == "serAx")) {
    return bad_arg("chart: a category axis has no major unit");
  }
  if (patch.major_unit.is_clear()) {
    drop(ax, "majorUnit");
  } else if (patch.major_unit.is_set()) {
    if (!(std::isfinite(patch.major_unit.value) && patch.major_unit.value > 0.0)) {
      return bad_arg("chart: major unit must be positive");
    }
    put_val(ax, "c:majorUnit", kAxisOrder, num(patch.major_unit.value));
  }
  if (patch.number_format.is_clear()) {
    drop(ax, "numFmt");
  } else if (patch.number_format.is_set()) {
    xml_node fmt = ensure(ax, "c:numFmt", kAxisOrder);
    set_attr(fmt, "formatCode", patch.number_format.value);
    set_attr(fmt, "sourceLinked", "0");
  }
  if (patch.major_gridlines.has_value()) {
    if (!*patch.major_gridlines) {
      drop(ax, "majorGridlines");
    } else if (!child_local(ax, "majorGridlines")) {
      pugi::xml_document frag;
      copy_at(ax, fragment(frag, kGridlines), kAxisOrder);
    }
  }
  if (patch.crosses.has_value()) {
    static const char* const kCrosses[] = {"", "autoZero", "min", "max"};
    const ChartAxisCrosses cr = *patch.crosses;
    drop(ax, "crosses");
    drop(ax, "crossesAt");
    if (cr == ChartAxisCrosses::kAt) {
      if (!std::isfinite(patch.crosses_at)) {
        return bad_arg("chart: crossing point must be finite");
      }
      put_val(ax, "c:crossesAt", kAxisOrder, num(patch.crosses_at));
    } else if (cr != ChartAxisCrosses::kAbsent) {
      const int i = static_cast<int>(cr);
      if (i < 1 || i > 3) {
        return bad_arg("chart: unknown axis crossing");
      }
      put_val(ax, "c:crosses", kAxisOrder, kCrosses[i]);
    }
  }
  if (!patch.tick_label_pos.is_keep()) {
    const std::string pos = patch.tick_label_pos.is_set() ? patch.tick_label_pos.value : "nextTo";
    if (pos != "nextTo" && pos != "low" && pos != "high" && pos != "none") {
      return bad_arg("chart: tick label position must be nextTo, low, high or none");
    }
    put_val(ax, "c:tickLblPos", kAxisOrder, pos);
  }
  if (patch.reversed.has_value() || patch.crosses.has_value()) {
    place_axis(c, ax);
    for (xml_node n = c.area().first_child(); n; n = n.next_sibling()) {
      if (n != ax && child_local(n, "crossAx") && val_u(n, "crossAx") == axis_id) {
        place_axis(c, n);
      }
    }
  }
  return commit(wb, c, false);
}

Expected<void, Error> set_chart_display(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                        ChartBlanksAs blanks_as, Tri plot_vis_only) {
  Chart c;
  RETURN_IF_ERROR(open_edit(wb, sheet, object_id, c, false));
  if (static_cast<int>(blanks_as) < 0 || static_cast<int>(blanks_as) > 3 || static_cast<int>(plot_vis_only) > 2) {
    return bad_arg("chart: unknown display option");
  }
  patch_display(c, blanks_as, plot_vis_only);
  return commit(wb, c, false);
}

Expected<void, Error> set_chart_series(Workbook& wb, std::size_t sheet, std::uint32_t object_id, std::size_t plot,
                                       std::uint32_t series_idx, const ChartSeriesPatch& patch) {
  Chart c;
  RETURN_IF_ERROR(open_edit(wb, sheet, object_id, c, false));
  ASSIGN_OR_RETURN(xml_node node, plot_at(c, plot));
  ASSIGN_OR_RETURN(xml_node ser, series_at(node, series_idx));
  bool refs = false;
  RETURN_IF_ERROR(patch_series(wb, c, node, ser, patch, &refs));
  return commit(wb, c, refs);
}

Expected<void, Error> set_chart_point(Workbook& wb, std::size_t sheet, std::uint32_t object_id, std::size_t plot,
                                      std::uint32_t series_idx, std::uint32_t point_idx, const ChartPointPatch& patch) {
  Chart c;
  RETURN_IF_ERROR(open_edit(wb, sheet, object_id, c, false));
  ASSIGN_OR_RETURN(xml_node node, plot_at(c, plot));
  ASSIGN_OR_RETURN(xml_node ser, series_at(node, series_idx));
  xml_node point = find_point(ser, point_idx);
  if (fill_empty(patch.fill) && line_empty(patch.line) && patch.marker.is_keep()) {
    if (point) {
      ser.remove_child(point);
    }
    return commit(wb, c, false);
  }
  if (!patch.marker.is_keep() && !has_markers(node)) {
    return bad_arg("chart: markers need a line, scatter or radar plot");
  }
  const std::string_view name = local_name(node);
  if (!point && !is_pie(node)) {
    pugi::xml_document frag;
    const std::string x = "<c:dPt><c:idx val=\"" + std::to_string(point_idx) + "\"/>" +
                          (is_bar(node) ? "<c:invertIfNegative val=\"0\"/>" : "") +
                          "<c:bubble3D val=\"0\"/><c:spPr><a:effectLst/></c:spPr></c:dPt>";
    point = insert_point(ser, fragment(frag, x), point_idx);
    // As Excel saves a new point: a bar or area point repeats the series fill and line, a line point the series
    // line and marker, and a scatter point those too once it has a line or marker of its own.
    const xml_node from = child_local(ser, "spPr");
    const bool filled = is_bar(node) || name == "areaChart";
    const bool stroked =
        name == "lineChart" || (name == "scatterChart" && (!line_empty(patch.line) || !patch.marker.is_keep()));
    for (xml_node e = from.first_child(); e; e = e.next_sibling()) {
      if (local_name(e) == "ln" ? filled || stroked : filled && rank(kFills, local_name(e)) >= 0) {
        copy_at(child_local(point, "spPr"), e, kShapeOrder);
      }
    }
    if (const xml_node marker = child_local(ser, "marker"); marker && stroked) {
      copy_at(point, marker, kPointOrder);
    }
  } else if (!point) {
    pugi::xml_document frag;
    point = insert_point(
        ser, fragment(frag, "<c:dPt><c:idx val=\"" + std::to_string(point_idx) + "\"/><c:bubble3D val=\"0\"/></c:dPt>"),
        point_idx);
  }
  RETURN_IF_ERROR(apply_shape(point, kPointOrder, marker_filled(node) ? ChartFillPatch{} : patch.fill, patch.line));
  RETURN_IF_ERROR(apply_marker(point, kPointOrder, patch.marker));
  return commit(wb, c, false);
}

Expected<void, Error> set_chart_plot(Workbook& wb, std::size_t sheet, std::uint32_t object_id, std::size_t plot,
                                     const ChartPlotPatch& patch) {
  Chart c;
  RETURN_IF_ERROR(open_edit(wb, sheet, object_id, c, false));
  ASSIGN_OR_RETURN(xml_node node, plot_at(c, plot));
  const std::string_view name = local_name(node);
  if (patch.gap_width) {
    if (!(is_bar(node) || name == "ofPieChart") || *patch.gap_width > 500U) {
      return bad_arg("chart: gap width needs a bar plot and 0..500");
    }
    put_val(node, "c:gapWidth", kPlotOrder, std::to_string(*patch.gap_width));
  }
  if (patch.overlap) {
    if (name != "barChart" || *patch.overlap < -100 || *patch.overlap > 100) {
      return bad_arg("chart: overlap needs a 2-D bar plot and -100..100");
    }
    put_val(node, "c:overlap", kPlotOrder, std::to_string(*patch.overlap));
  }
  if (patch.first_slice_angle) {
    if ((name != "pieChart" && name != "doughnutChart") || *patch.first_slice_angle > 360U) {
      return bad_arg("chart: first slice angle needs a pie plot and 0..360");
    }
    put_val(node, "c:firstSliceAng", kPlotOrder, std::to_string(*patch.first_slice_angle));
  }
  if (patch.hole_size) {
    if (name != "doughnutChart" || *patch.hole_size < 10U || *patch.hole_size > 90U) {
      return bad_arg("chart: hole size needs a doughnut plot and 10..90");
    }
    put_val(node, "c:holeSize", kPlotOrder, std::to_string(*patch.hole_size));
  }
  if (patch.vary_colors) {
    const bool was = flag(child_local(node, "varyColors"));
    put_val(node, "c:varyColors", kPlotOrder, *patch.vary_colors ? "1" : "0");
    if (name == "barChart" && was != *patch.vary_colors) {
      vary_bar(node, *patch.vary_colors);
    }
  }
  return commit(wb, c, false);
}

Expected<void, Error> set_chart_filter(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                       const std::vector<std::uint32_t>& hidden_series_idx,
                                       const std::vector<std::uint32_t>& hidden_category_idx) {
  Chart c;
  RETURN_IF_ERROR(open_edit(wb, sheet, object_id, c, true));
  std::set<std::uint32_t> all;
  for (const xml_node& plot : plots_of(c)) {
    for (const xml_node& s : plot_series(plot)) {
      all.insert(val_u(s, "idx"));
    }
  }
  const std::set<std::uint32_t> hidden(hidden_series_idx.begin(), hidden_series_idx.end());
  for (const std::uint32_t idx : hidden) {
    if (all.count(idx) == 0) {
      return bad_arg("chart: no series with that idx");
    }
  }
  patch_series_filter(c, hidden);
  RETURN_IF_ERROR(
      patch_category_filter(wb, c, std::set<std::uint32_t>(hidden_category_idx.begin(), hidden_category_idx.end())));
  return commit(wb, c, true);
}

Expected<std::uint32_t, Error> add_chart_series(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                                std::size_t plot, const ChartSeriesPatch& patch) {
  Chart c;
  RETURN_IF_ERROR(open_edit(wb, sheet, object_id, c, true));
  ASSIGN_OR_RETURN(xml_node node, plot_at(c, plot));
  const int kind = kind_of(node);
  if (kind < 0) {
    return unsupported("chart: series are added to column, bar, line, area, pie, doughnut and scatter plots");
  }
  const bool scatter = kind == static_cast<int>(ChartInsertKind::kScatter);
  if (!(scatter ? patch.y : patch.val) || (scatter ? patch.y : patch.val)->empty()) {
    return bad_arg("chart: a series needs its values");
  }
  std::uint32_t idx = 0;
  std::uint32_t order = 0;
  for (const xml_node& p : plots_of(c)) {
    for (const xml_node& s : plot_series(p)) {
      idx = std::max(idx, val_u(s, "idx") + 1U);
      order = std::max(order, val_u(s, "order") + 1U);
    }
  }
  // The template numbers and colours series by position, so the new one is placed at `idx`.
  ChartSeriesRefs refs;
  refs.val = refs.y = scatter ? *patch.y : *patch.val;
  ChartSplit split;
  split.series.assign(idx + 1U, refs);
  pugi::xml_document t;
  RETURN_IF_ERROR(template_doc(wb, static_cast<ChartInsertKind>(kind), grouping_of(node), std::nullopt,
                               ChartLegendPosition::kRight, split, t));
  xml_node ser = copy_at(node, kids(template_plot(t), "ser").back(), kPlotOrder);
  put_val(ser, "c:order", kSerOrder, std::to_string(order));
  bool refs_changed = false;
  RETURN_IF_ERROR(patch_series(wb, c, node, ser, patch, &refs_changed));
  RETURN_IF_ERROR(commit(wb, c, true));
  return idx;
}

Expected<void, Error> remove_chart_series(Workbook& wb, std::size_t sheet, std::uint32_t object_id, std::size_t plot,
                                          std::uint32_t series_idx) {
  Chart c;
  RETURN_IF_ERROR(open_edit(wb, sheet, object_id, c, true));
  ASSIGN_OR_RETURN(xml_node node, plot_at(c, plot));
  ASSIGN_OR_RETURN(xml_node ser, series_at(node, series_idx));
  ser.parent().remove_child(ser);
  drop_empty_group(node);
  return commit(wb, c, false);
}

Expected<void, Error> set_chart_source(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                       const std::string& source, ChartSeriesDirection direction,
                                       ChartHeaderMode header_row, ChartHeaderMode header_column) {
  Chart c;
  RETURN_IF_ERROR(open_edit(wb, sheet, object_id, c, true));
  RETURN_IF_ERROR(patch_source(wb, c, source, direction, header_row, header_column));
  return commit(wb, c, true);
}

}  // namespace formulon
