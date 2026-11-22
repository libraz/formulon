#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "chart/chart_xml_internal.h"
#include "drawing/drawing_xml.h"

namespace formulon {
namespace chart_xml_detail {

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

void read_chartex(const Workbook& wb, const xml_node& root, ChartModel& model) {
  model.is_chartex = true;
  model.date1904 = false;
  const ColorEnv env{wb, kid(root, "clrMapOvr")};
  read_shape(env, kid(root, "spPr"), model.chart_area);

  struct Dim {
    std::string type;
    ChartDataSource source;
  };
  struct Data {
    std::string id;
    std::vector<Dim> dims;
  };
  std::vector<Data> data_sets;
  for (const xml_node& data : children_local(kid(root, "chartData"), "data")) {
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
  read_shape(env, kid(plot_area, "spPr"), model.plot_area);
  ChartPlot plot;
  plot.type = ChartPlotType::kChartex;
  for (const xml_node& ser : children_local(kid(plot_area, "plotAreaRegion"), "series")) {
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
    read_axis_ids(ser, "axisId", plot);
    plot.series.push_back(std::move(s));
  }
  model.plots.push_back(std::move(plot));

  model.title = read_cx_title(kid(chart, "title"));
  if (const xml_node legend = kid(chart, "legend")) {
    model.legend.present = true;
    model.legend.position = legend_position_of(legend.attribute("pos").value(), ChartLegendPosition::kRight);
    model.legend.overlay = attr_flag(legend, "overlay");
  }
  for (const xml_node& axis : children_local(plot_area, "axis")) {
    model.axes.push_back(read_cx_axis(axis));
  }
}

}  // namespace chart_xml_detail
}  // namespace formulon
