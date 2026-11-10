#include <algorithm>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
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
#include "io/ooxml/package_validator.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml_defs.h"
#include "io/xml_utils.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "unknown_relationship.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "workbook.h"

namespace formulon {
namespace chart_edit {

// ---------------------------------------------------------------------------
// DOM helpers
// ---------------------------------------------------------------------------

int rank(Order order, std::string_view name) {
  for (int i = 0; order[i] != nullptr; ++i) {
    if (name == order[i]) {
      return i;
    }
  }
  return -1;
}

xml_node slot(const xml_node& parent, std::string_view name, Order order) {
  const int r = rank(order, name);
  for (xml_node c = parent.first_child(); c; c = c.next_sibling()) {
    if (c.type() == pugi::node_element && rank(order, local_name(c)) > r) {
      return c;
    }
  }
  return {};
}

xml_node copy_at(xml_node parent, const xml_node& src, Order order, const char* qname) {
  xml_node node = parent.append_copy(src);
  if (qname != nullptr) {
    node.set_name(qname);
  }
  const xml_node before = slot(parent, local_name(node), order);
  return before ? parent.insert_move_before(node, before) : node;
}

// ---------------------------------------------------------------------------
// The chart being read or edited
// ---------------------------------------------------------------------------

Expected<void, Error> open_chart(const Workbook& wb, std::size_t sheet, std::uint32_t object_id, Chart& c) {
  ASSIGN_OR_RETURN(std::vector<ChartPartRef> parts, list_chart_parts(wb, sheet));
  const auto it =
      std::find_if(parts.begin(), parts.end(), [&](const ChartPartRef& p) { return p.object_id == object_id; });
  if (it == parts.end()) {
    return no_such("No chart with that id");
  }
  const PassthroughPart* part = it->part_path.empty() ? nullptr : io::ooxml::find_passthrough_part(wb, it->part_path);
  if (part == nullptr) {
    return unparseable("Chart part is missing");
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
    return make_error(FormulonErrorCode::kIoDrawingUnparseable, "A chart retained from xlsb is not editable");
  }
  if (c.chartex) {
    return unsupported("Chartex charts are not editable");
  }
  if (series_op && c.pivot()) {
    return unsupported("A pivot chart's series follow its pivot table");
  }
  c.before = io::raw_xml(c.root());
  return ok();
}

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

std::vector<xml_node> plots_of(const Chart& c) {
  std::vector<xml_node> out;
  for (xml_node n = c.area().first_child(); n; n = n.next_sibling()) {
    if (is_plot(n)) {
      out.push_back(n);
    }
  }
  return out;
}

int kind_of(const xml_node& plot) {
  const std::string_view name = local_name(plot);
  if (name == "barChart") {
    return std::string_view(child_local(plot, "barDir").attribute("val").value()) == "bar"
               ? static_cast<int>(ChartInsertKind::kBar)
               : static_cast<int>(ChartInsertKind::kColumn);
  }
  static const char* const kNames[] = {"lineChart", "areaChart", "pieChart", "scatterChart", "doughnutChart"};
  for (int i = 0; i < static_cast<int>(std::size(kNames)); ++i) {
    if (name == kNames[i]) {
      return static_cast<int>(ChartInsertKind::kLine) + i;
    }
  }
  return -1;
}

xml_node filter_group(const xml_node& plot) {
  for (const xml_node& ext : children_local(child_local(plot, "extLst"), "ext")) {
    for (xml_node g = ext.first_child(); g; g = g.next_sibling()) {
      if (is_filtered_series_group(g)) {
        return g;
      }
    }
  }
  return {};
}

Expected<xml_node, Error> series_at(const xml_node& plot, std::uint32_t idx) {
  for (const xml_node& s : plot_series(plot)) {
    if (val_u(s, "idx") == idx) {
      return s;
    }
  }
  return no_such("No series with that idx");
}

xml_node find_ext(const xml_node& ext_list, std::string_view uri) {
  for (const xml_node& ext : children_local(ext_list, "ext")) {
    if (std::string_view(ext.attribute("uri").value()) == uri) {
      return ext;
    }
  }
  return {};
}

xml_node c15_ext(xml_node ref, bool create) {
  xml_node list = child_local(ref, "extLst");
  if (xml_node ext = find_ext(list, kUriC15)) {
    if (create) {
      bind_c15(ext);
    }
    return ext;
  }
  if (!create) {
    return {};
  }
  if (!list) {
    list = ref.prepend_child("c:extLst");
  }
  xml_node ext = list.prepend_child("c:ext");
  ext.append_attribute("uri").set_value(kUriC15);
  bind_c15(ext);
  return ext;
}

void prune_ext(xml_node ref) {
  xml_node list = child_local(ref, "extLst");
  for (const xml_node& ext : children_local(list, "ext")) {
    if (!io::first_element_child(ext)) {
      list.remove_child(ext);
    }
  }
  if (list && !io::first_element_child(list)) {
    ref.remove_child(list);
  }
}

std::uint32_t point_count(const xml_node& ser) {
  const xml_node val = value_input(ser);
  const xml_node ref = ref_of(val);
  for (xml_node c = ref.first_child(); c; c = c.next_sibling()) {
    if (child_local(c, "ptCount")) {
      return val_u(c, "ptCount");
    }
  }
  return 0;
}

void load_palette(const Workbook& wb, const std::string& chart_path, Palette& out) {
  const PassthroughPart* rels = io::ooxml::find_passthrough_part(wb, io::ooxml::rels_path_for_part(chart_path));
  pugi::xml_document doc;
  std::string colors;
  if (rels != nullptr && doc.load_buffer(rels->bytes.data(), rels->bytes.size())) {
    for (const xml_node& r : children_local(doc.document_element(), "Relationship")) {
      const std::string_view type = r.attribute("Type").value();
      const auto ends = [&](std::string_view tail) { return strings::ends_with(type, tail); };
      if (ends("/chartStyle")) {
        out.styled = true;
      } else if (ends("/chartColorStyle")) {
        auto path = io::ooxml::resolve_relative_path(io::ooxml::dir_of(chart_path), r.attribute("Target").value());
        const PassthroughPart* part = path ? io::ooxml::find_passthrough_part(wb, path.value()) : nullptr;
        if (part != nullptr) {
          colors.assign(part->bytes.begin(), part->bytes.end());
        }
      }
    }
  }
  if (colors.empty() || !load_xml(out.colors, colors) ||
      std::string_view(out.colors.document_element().attribute("meth").value()) != "cycle") {
    load_xml(out.colors, kDefaultChartColorStyle);
  }
}

Expected<void, Error> template_doc(const Workbook& wb, const Chart& c, ChartInsertKind kind, ChartGrouping grouping,
                                   const std::optional<std::string>& title, ChartLegendPosition legend,
                                   ChartSplit split, pugi::xml_document& out) {
  Palette palette;
  load_palette(wb, c.path, palette);
  return template_doc(wb, palette, kind, grouping, title, legend, std::move(split), out);
}

Expected<void, Error> template_doc(const Workbook& wb, const Palette& palette, ChartInsertKind kind,
                                   ChartGrouping grouping, const std::optional<std::string>& title,
                                   ChartLegendPosition legend, ChartSplit split, pugi::xml_document& out) {
  if (split.series.empty()) {
    split.series.push_back(ChartSeriesRefs{"", "", "A1", "A1", "A1"});
  }
  ChartInsertOptions options;
  options.kind = kind;
  options.grouping = grouping;
  options.title = title;
  options.legend = legend;
  ASSIGN_OR_RETURN(std::string xml, build_chart_xml(wb, options, split, palette.colors.document_element()));
  if (!load_xml(out, xml)) {
    return unparseable("Template does not parse");
  }
  return ok();
}

xml_node template_plot(const pugi::xml_document& t) {
  for (xml_node n = child_local(template_chart(t), "plotArea").first_child(); n; n = n.next_sibling()) {
    if (is_plot(n)) {
      return n;
    }
  }
  return {};
}

}  // namespace chart_edit
}  // namespace formulon
