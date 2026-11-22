// Readers shared by the DrawingML and chartex halves of the chart part reader (chart_xml.cpp,
// chart_xml_cx.cpp). Internal to those two translation units.

#ifndef FORMULON_CHART_CHART_XML_INTERNAL_H_
#define FORMULON_CHART_CHART_XML_INTERNAL_H_

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "chart/chart_model.h"
#include "pugixml.hpp"

namespace formulon {

class Workbook;

namespace chart_xml_detail {

using pugi::xml_node;

/// Workbook and the part's colour map: a remapped `bg1` / `tx1` / `bg2` / `tx2` reads as Rgb.
struct ColorEnv {
  const Workbook& wb;
  xml_node clr_map_ovr;  ///< The chart's `c:clrMapOvr`; empty when it has none.
};

xml_node kid(const xml_node& parent, std::string_view name);
std::string_view val_of(const xml_node& node);
bool attr_flag(const xml_node& node, const char* name);
double double_text(std::string_view text);
ChartFill read_fill(const ColorEnv& env, const xml_node& sp_pr);
ChartLine read_line(const ColorEnv& env, const xml_node& sp_pr);

/// Reads the fill and the line of one shape-properties element into `owner`.
template <typename Shaped>
void read_shape(const ColorEnv& env, const xml_node& sp_pr, Shaped& owner) {
  owner.fill = read_fill(env, sp_pr);
  owner.line = read_line(env, sp_pr);
}

std::optional<ChartLabelPosition> label_position_of(std::string_view name);
ChartNumFmt read_num_fmt(const xml_node& node);
ChartValue make_value(std::string_view text, bool numeric);
void read_points(const xml_node& cache, bool numeric, bool direct_text, std::vector<ChartValue>& out,
                 std::uint32_t& point_count);
std::string flatten_rich(const xml_node& rich);
ChartLegendPosition legend_position_of(std::string_view name, ChartLegendPosition fallback);
void read_axis_ids(const xml_node& parent, std::string_view name, ChartPlot& plot);

/// Reads a chartex (`cx:`) chart space into `model`.
void read_chartex(const Workbook& wb, const xml_node& root, ChartModel& model);

}  // namespace chart_xml_detail
}  // namespace formulon

#endif  // FORMULON_CHART_CHART_XML_INTERNAL_H_
