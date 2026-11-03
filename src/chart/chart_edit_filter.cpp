#include <cstring>
#include <initializer_list>
#include <optional>
#include <set>
#include <string>
#include <string_view>
#include <unordered_set>
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
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "sheet.h"
#include "utils/status_macros.h"
#include "workbook.h"

namespace formulon {
namespace chart_edit {

namespace {

// ---------------------------------------------------------------------------
// References and filters
// ---------------------------------------------------------------------------

/// Replaces `ser`'s `tag` input with a reference to `formula` and an empty cache.
void put_ref(xml_node ser, const char* tag, const std::string& formula, bool multi, bool filtered) {
  drop(ser, local_name(tag));
  const bool name = local_name(tag) == "tx";
  const char* ref = name ? "strRef" : multi ? "multiLvlStrRef" : "numRef";
  std::string x = std::string("<") + tag + "><c:" + ref + ">";
  std::string esc;
  io::AppendXmlAttrEscaped(esc, formula);
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
  c15_ext(ref, true).append_child("c15:formulaRef").append_child("c15:sqref").text().set(text.c_str());
}

void show_ref(xml_node ref) {
  xml_node ext = c15_ext(ref, false);
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

}  // namespace

void drop_empty_group(xml_node plot) {
  xml_node group = filter_group(plot);
  if (!group || io::first_element_child(group)) {
    return;
  }
  xml_node ext = group.parent();
  ext.remove_child(group);
  if (!io::first_element_child(ext)) {
    xml_node list = ext.parent();
    list.remove_child(ext);
    if (!io::first_element_child(list)) {
      plot.remove_child(list);
    }
  }
}

void patch_series_filter(const Chart& c, const std::set<std::uint32_t>& hidden) {
  for (xml_node plot : plots_of(c)) {
    for (xml_node ser : children_local(plot, "ser")) {
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
      bind_c15(group);
      xml_node moved = group.append_move(ser);
      moved.set_name("c15:ser");
      each_ref(moved, hide_ref);
    }
    for (xml_node ser : children_local(filter_group(plot), "ser")) {
      if (hidden.count(val_u(ser, "idx")) != 0) {
        continue;
      }
      each_ref(ser, show_ref);
      const std::uint32_t order = val_u(ser, "order");
      xml_node before;
      for (const xml_node& s : children_local(plot, "ser")) {
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
        xml_node ext = c15_ext(ref, false);
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
          ext = c15_ext(ref, true);
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
    for (const xml_node& s : children_local(filter_group(plot), "ser")) {
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
      drop(ser, local_name(f.tag));
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
    Palette palette;
    load_palette(wb, c.path, palette);
    settle_marker(ser, ser, palette);
  }
  if (p.smooth.has_value()) {
    put_val(ser, "c:smooth", kSerOrder, *p.smooth ? "1" : "0");
  }
  return apply_labels(ser, is_pie(plot), c.lang(), p.labels);
}

namespace {

std::string accent_sppr(const Palette& p, std::uint32_t i) {
  return "<c:spPr><a:solidFill>" + chart_cycle_color(p.colors.document_element(), i) +
         "</a:solidFill><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>";
}

}  // namespace

xml_node insert_point(xml_node ser, const xml_node& src, std::uint32_t idx) {
  for (const xml_node& d : children_local(ser, "dPt")) {
    if (val_u(d, "idx") > idx) {
      return ser.insert_copy_before(src, d);
    }
  }
  return copy_at(ser, src, kSerOrder);
}

xml_node find_point(const xml_node& ser, std::uint32_t idx) {
  for (const xml_node& d : children_local(ser, "dPt")) {
    if (val_u(d, "idx") == idx) {
      return d;
    }
  }
  return {};
}

void vary_bar(xml_node plot, bool on, const Palette& palette) {
  for (xml_node ser : plot_series(plot)) {
    if (on) {
      drop(ser, "spPr");
      const std::uint32_t n = point_count(ser);
      const std::vector<xml_node> original_points = children_local(ser, "dPt");
      std::vector<std::uint32_t> original_indices;
      original_indices.reserve(original_points.size());
      std::unordered_set<std::uint32_t> existing_ids;
      existing_ids.reserve(original_points.size());
      for (const xml_node& point : original_points) {
        const std::uint32_t idx = val_u(point, "idx");
        original_indices.push_back(idx);
        existing_ids.insert(idx);
      }
      const xml_node later_schema_slot = slot(ser, "dPt", kSerOrder);
      std::size_t next_original = 0;
      for (std::uint32_t p = 0; p < n; ++p) {
        if (existing_ids.find(p) != existing_ids.end()) {
          continue;
        }
        while (next_original < original_points.size() && original_indices[next_original] <= p) {
          ++next_original;
        }
        pugi::xml_document frag;
        const xml_node point = fragment(frag, "<c:dPt><c:idx val=\"" + std::to_string(p) +
                                                  "\"/><c:invertIfNegative val=\"0\"/><c:bubble3D val=\"0\"/>" +
                                                  accent_sppr(palette, p) + "</c:dPt>");
        if (next_original < original_points.size()) {
          ser.insert_copy_before(point, original_points[next_original]);
        } else if (later_schema_slot) {
          ser.insert_copy_before(point, later_schema_slot);
        } else {
          ser.append_copy(point);
        }
      }
      continue;
    }
    for (const xml_node& d : children_local(ser, "dPt")) {
      const std::string_view clr =
          child_local(child_local(child_local(d, "spPr"), "solidFill"), "schemeClr").attribute("val").value();
      if (clr.substr(0, 6) == "accent") {
        ser.remove_child(d);
      }
    }
    if (!child_local(ser, "spPr")) {
      pugi::xml_document frag;
      copy_at(ser, fragment(frag, accent_sppr(palette, val_u(ser, "idx"))), kSerOrder);
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

}  // namespace chart_edit
}  // namespace formulon
