#include "chart/chart_edit.h"

#include <algorithm>
#include <cmath>
#include <iterator>
#include <optional>
#include <set>
#include <string>
#include <string_view>
#include <vector>

#include "chart/chart_cache.h"
#include "chart/chart_data.h"
#include "chart/chart_edit_internal.h"
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
#include "utils/status_macros.h"
#include "workbook.h"

namespace formulon {

using namespace chart_edit;

namespace {

// Excel's default chart frame, 360 x 216 pt.
constexpr std::int64_t kFrameCx = 4572000;
constexpr std::int64_t kFrameCy = 2743200;

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
  RETURN_IF_ERROR(template_doc(wb, c, static_cast<ChartInsertKind>(kind), grouping_of(plot), std::nullopt,
                               ChartLegendPosition::kRight, split, t));
  patch_series_filter(c, {});
  RETURN_IF_ERROR(patch_category_filter(wb, c, {}));
  const std::vector<xml_node> fresh = children_local(template_plot(t), "ser");
  for (const xml_node& ser : children_local(plot, "ser")) {
    if (val_u(ser, "idx") >= fresh.size()) {
      plot.remove_child(ser);
    }
  }
  for (std::uint32_t i = 0; i < fresh.size(); ++i) {
    xml_node ser;
    for (const xml_node& s : children_local(plot, "ser")) {
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
    for (const xml_node& d : children_local(fresh[i], "dPt")) {
      if (!find_point(ser, val_u(d, "idx"))) {
        insert_point(ser, d, val_u(d, "idx"));
      }
    }
  }
  return ok();
}

/// True when `patch` clears every member of a point's override.
bool clears_all(const ChartPointPatch& patch) {
  const ChartFillPatch& f = patch.fill;
  const ChartLinePatch& l = patch.line;
  return f.color.is_clear() && f.transparency_pct.is_clear() && l.color.is_clear() && l.transparency_pct.is_clear() &&
         l.width_pt.is_clear() && l.dash.is_clear() && !l.none.value_or(false) && patch.marker.is_clear();
}

/// True when the `c:dPt` `point` holds no fill, line colour, width, dash or marker of its own.
bool line_patch_sets(const ChartLinePatch& p) {
  return p.color.is_set() || p.transparency_pct.is_set() || p.width_pt.is_set() || p.dash.is_set() ||
         p.none.value_or(false);
}

bool point_patch_sets(const ChartPointPatch& p) {
  return p.fill.color.is_set() || p.fill.transparency_pct.is_set() || line_patch_sets(p.line) || p.marker.is_set();
}

bool point_override_empty(const xml_node& point) {
  for (xml_node n = point.first_child(); n; n = n.next_sibling()) {
    const std::string_view name = local_name(n);
    if (name == "idx" || name == "invertIfNegative" || name == "bubble3D") {
      continue;
    }
    if (name != "spPr") {
      return false;
    }
    for (xml_node sp = n.first_child(); sp; sp = sp.next_sibling()) {
      const std::string_view part = local_name(sp);
      if (part == "effectLst") {
        continue;
      }
      if (part != "ln" || sp.attribute("w")) {
        return false;
      }
      for (xml_node ln = sp.first_child(); ln; ln = ln.next_sibling()) {
        const std::string_view piece = local_name(ln);
        if (piece != "round" && piece != "bevel" && piece != "miter" && piece != "headEnd" && piece != "tailEnd") {
          return false;
        }
      }
    }
  }
  return true;
}

/// Checks a title that is not kNone and returns the template text it needs: the text for kText, else unset.
Expected<std::optional<std::string>, Error> title_text(const Workbook& wb, const Chart& c, const ChartTitle& title) {
  if (title.kind == ChartTitleKind::kRef) {
    if (title.ref_formula.empty()) {
      return bad_arg("chart: a reference title needs a formula");
    }
    RETURN_IF_ERROR(resolve_chart_source(wb, c.host, title.ref_formula, eval::default_registry()));
  } else if (title.kind == ChartTitleKind::kText && title.text.empty()) {
    return bad_arg("chart: a text title needs text");
  } else if (title.kind != ChartTitleKind::kText && title.kind != ChartTitleKind::kAuto) {
    return bad_arg("chart: unknown title kind");
  }
  return title.kind == ChartTitleKind::kText ? std::optional<std::string>(title.text) : std::nullopt;
}

/// Writes `title` into the `c:title` of `host`, creating it from template title `from`. An existing title keeps
/// its own formatting; `c:txPr` follows the template.
void put_title(xml_node host, Order order, const xml_node& from, const ChartTitle& title) {
  xml_node node = child_local(host, "title");
  if (!node) {
    node = copy_at(host, from, order);
  }
  drop(node, "tx");
  if (title.kind == ChartTitleKind::kText) {
    copy_at(node, child_local(from, "tx"), kTitleOrder);
  }
  if (!child_local(from, "txPr")) {
    drop(node, "txPr");
  } else if (!child_local(node, "txPr")) {
    copy_at(node, child_local(from, "txPr"), kTitleOrder);
  }
  if (title.kind == ChartTitleKind::kRef) {
    std::string x = "<c:tx><c:strRef><c:f>";
    io::AppendXmlRawTextEscaped(x, title.ref_formula);
    x += "</c:f></c:strRef></c:tx>";
    pugi::xml_document frag;
    copy_at(node, fragment(frag, x), kTitleOrder);
  }
}

Expected<void, Error> patch_title(const Workbook& wb, const Chart& c, const ChartTitle& title) {
  xml_node chart = c.chart();
  if (title.kind == ChartTitleKind::kNone) {
    drop(chart, "title");
    put_val(chart, "c:autoTitleDeleted", kChartOrder, "1");
    return ok();
  }
  ASSIGN_OR_RETURN(std::optional<std::string> text, title_text(wb, c, title));
  pugi::xml_document t;
  RETURN_IF_ERROR(template_doc(wb, c, ChartInsertKind::kColumn, ChartGrouping::kClustered, text,
                               ChartLegendPosition::kRight, {}, t));
  put_title(chart, kChartOrder, child_local(template_chart(t), "title"), title);
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
  ChartDisplayPatch display;
  if (o.blanks_as) {
    display.blanks_as = ChartPatchField<ChartBlanksAs>::Set(*o.blanks_as);
  }
  if (o.plot_vis_only) {
    display.plot_vis_only = ChartPatchField<bool>::Set(*o.plot_vis_only);
  }
  if (o.na_as_blank) {
    display.na_as_blank = ChartPatchField<bool>::Set(*o.na_as_blank);
  }
  patch_display(c, display);
  if (resplit) {
    refresh_chart_caches(wb, c.host, c.doc);
  }
  return data_of(wb, sheet, object_id, c);
}

Expected<std::uint32_t, Error> insert_chart(Workbook& wb, std::size_t sheet, const ChartInsertOptions& options) {
  if (sheet >= wb.sheet_count()) {
    return bad_arg("chart: sheet index out of range");
  }
  if (!valid_kind_grouping(options.kind, options.grouping)) {
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
                           "\" xmlns:r=\"" + std::string(io::kNsOfficeRelationships) +
                           "\" r:id=\"{rId}\"/></a:graphicData>";
  ASSIGN_OR_RETURN(std::uint32_t id, insert_graphic_frame(wb, sheet, placement, data, path, io::kRelChart));
  std::string bytes(io::kXmlDecl);
  io::append_raw_xml(bytes, doc.document_element());
  RETURN_IF_ERROR(wb.add_passthrough_part(
      PassthroughPart(path, std::string(io::kCtChart), std::vector<std::uint8_t>(bytes.begin(), bytes.end()))));
  return id;
}

Expected<ChartInsertKind, Error> chart_insert_kind(ChartPlotType type, ChartBarDirection bar_direction) {
  if (type == ChartPlotType::kBar) {
    switch (bar_direction) {
      case ChartBarDirection::kAbsent:
      case ChartBarDirection::kCol:
        return ChartInsertKind::kColumn;
      case ChartBarDirection::kBar:
        return ChartInsertKind::kBar;
    }
    return bad_arg("chart: unknown bar direction");
  }
  if (bar_direction != ChartBarDirection::kAbsent) {
    return bad_arg("chart: a bar direction needs a bar plot");
  }
  switch (type) {
    case ChartPlotType::kLine:
      return ChartInsertKind::kLine;
    case ChartPlotType::kArea:
      return ChartInsertKind::kArea;
    case ChartPlotType::kPie:
      return ChartInsertKind::kPie;
    case ChartPlotType::kDoughnut:
      return ChartInsertKind::kDoughnut;
    case ChartPlotType::kScatter:
      return ChartInsertKind::kScatter;
    default:
      return bad_arg("chart: only bar, line, area, pie, doughnut and scatter plots are written");
  }
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
  if (pos < 0 || pos >= static_cast<int>(std::size(kChartLegendTokens))) {
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
        template_doc(wb, c, ChartInsertKind::kColumn, ChartGrouping::kClustered, std::nullopt, legend.position, {}, t));
    node = copy_at(chart, child_local(template_chart(t), "legend"), kChartOrder);
  }
  put_val(node, "c:legendPos", kLegendOrder, kChartLegendTokens[pos]);
  put_val(node, "c:overlay", kLegendOrder, legend.overlay ? "1" : "0");
  if (legend.hidden_series_idx.has_value()) {
    const std::set<std::uint32_t> hidden(legend.hidden_series_idx->begin(), legend.hidden_series_idx->end());
    for (const xml_node& e : children_local(node, "legendEntry")) {
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
  if (patch.title.is_clear() || (patch.title.is_set() && patch.title.value.kind == ChartTitleKind::kNone)) {
    drop(ax, "title");
  } else if (patch.title.is_set()) {
    ASSIGN_OR_RETURN(std::optional<std::string> text, title_text(wb, c, patch.title.value));
    const std::string_view pos = child_local(ax, "axPos").attribute("val").value();
    pugi::xml_document t;
    put_title(ax, kAxisOrder, fragment(t, build_axis_title_xml(wb, text, pos == "l" || pos == "r")), patch.title.value);
  }
  const auto scaling = [&] { return ensure(ax, "c:scaling", kAxisOrder); };
  const ChartPatchField<double>* bounds[] = {&patch.min, &patch.max};
  const char* bound_names[] = {"c:min", "c:max"};
  for (int i = 0; i < 2; ++i) {
    if (bounds[i]->is_clear()) {
      drop(scaling(), local_name(bound_names[i]));
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
  const ChartPatchField<double>* units[] = {&patch.major_unit, &patch.minor_unit};
  const char* unit_names[] = {"c:majorUnit", "c:minorUnit"};
  for (int i = 0; i < 2; ++i) {
    if (units[i]->is_set() && (kind == "catAx" || kind == "serAx")) {
      return bad_arg("chart: a category axis has no major or minor unit");
    }
    if (units[i]->is_clear()) {
      drop(ax, local_name(unit_names[i]));
    } else if (units[i]->is_set()) {
      if (!(std::isfinite(units[i]->value) && units[i]->value > 0.0)) {
        return bad_arg("chart: axis units must be positive");
      }
      put_val(ax, unit_names[i], kAxisOrder, num(units[i]->value));
    }
  }
  if (patch.number_format.is_clear()) {
    drop(ax, "numFmt");
  } else if (patch.number_format.is_set()) {
    xml_node fmt = ensure(ax, "c:numFmt", kAxisOrder);
    io::set_attr(fmt, "formatCode", patch.number_format.value);
    io::set_attr(fmt, "sourceLinked", "0");
  }
  link_number_format(ax, kAxisOrder, patch.source_linked);
  const std::optional<bool>* gridlines[] = {&patch.major_gridlines, &patch.minor_gridlines};
  for (int i = 0; i < 2; ++i) {
    const char* name = i == 0 ? "majorGridlines" : "minorGridlines";
    if (!gridlines[i]->has_value()) {
      continue;
    }
    if (!**gridlines[i]) {
      drop(ax, name);
    } else if (!child_local(ax, name)) {
      pugi::xml_document frag;
      copy_at(ax, fragment(frag, gridlines_xml(i == 1)), kAxisOrder);
    }
  }
  if (patch.position.has_value()) {
    const int p = static_cast<int>(*patch.position);
    if (p < 0 || p >= static_cast<int>(std::size(kChartAxisPositionTokens))) {
      return bad_arg("chart: unknown axis position");
    }
  }
  if (patch.crosses.has_value()) {
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
      if (i < 1 || i >= static_cast<int>(std::size(kChartCrossesTokens))) {
        return bad_arg("chart: unknown axis crossing");
      }
      put_val(ax, "c:crosses", kAxisOrder, kChartCrossesTokens[i]);
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
  if (patch.position.has_value()) {
    put_val(ax, "c:axPos", kAxisOrder, kChartAxisPositionTokens[static_cast<int>(*patch.position)]);
  }
  return commit(wb, c, false);
}

Expected<void, Error> set_chart_display(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                        const ChartDisplayPatch& patch) {
  Chart c;
  RETURN_IF_ERROR(open_edit(wb, sheet, object_id, c, false));
  if (patch.blanks_as.is_set() &&
      (static_cast<int>(patch.blanks_as.value) < 0 || static_cast<int>(patch.blanks_as.value) > 3)) {
    return bad_arg("chart: unknown display option");
  }
  if (patch.style.is_set() && (patch.style.value < 1 || patch.style.value > 48)) {
    return bad_arg("chart: style must be 1..48");
  }
  patch_display(c, patch);
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
    return ok();
  }
  if (clears_all(patch)) {
    if (point) {
      ser.remove_child(point);
    }
    return commit(wb, c, false);
  }
  // A plot without markers has none to set; clearing one (a read record passed back) changes nothing.
  if (patch.marker.is_set() && !has_markers(node)) {
    return bad_arg("chart: markers need a line, scatter or radar plot");
  }
  const std::string_view name = local_name(node);
  if (!point && !is_pie(node)) {
    pugi::xml_document frag;
    const std::string x = data_point_open(point_idx) + (is_bar(node) ? "<c:invertIfNegative val=\"0\"/>" : "") +
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
    point =
        insert_point(ser, fragment(frag, data_point_open(point_idx) + "<c:bubble3D val=\"0\"/></c:dPt>"), point_idx);
  }
  RETURN_IF_ERROR(apply_shape(point, kPointOrder, marker_filled(node) ? ChartFillPatch{} : patch.fill, patch.line));
  RETURN_IF_ERROR(apply_marker(point, kPointOrder, patch.marker));
  if (patch.marker.is_set()) {
    Palette palette;
    load_palette(wb, c.path, palette);
    settle_marker(ser, point, palette);
  }
  // An override that a clearing patch empties disappears; a patch that sets a member keeps its point.
  if (!point_patch_sets(patch) && point_override_empty(point)) {
    ser.remove_child(point);
  }
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
      Palette palette;
      load_palette(wb, c.path, palette);
      vary_bar(node, *patch.vary_colors, palette);
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
  const std::optional<ChartSeriesSource>& values = scatter ? patch.y : patch.val;
  if (!values || values->kind == ChartDataSource::Kind::kNone) {
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
  // The template numbers and colours series by position, so the new one is placed at `idx`. A literal input
  // replaces the placeholder reference below.
  ChartSeriesRefs refs;
  refs.val = refs.y = values->kind == ChartDataSource::Kind::kRef ? values->formula : std::string("A1");
  ChartSplit split;
  split.series.assign(idx + 1U, refs);
  pugi::xml_document t;
  RETURN_IF_ERROR(template_doc(wb, c, static_cast<ChartInsertKind>(kind), grouping_of(node), std::nullopt,
                               ChartLegendPosition::kRight, split, t));
  xml_node ser = copy_at(node, children_local(template_plot(t), "ser").back(), kPlotOrder);
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
