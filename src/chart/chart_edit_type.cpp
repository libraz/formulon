#include <algorithm>
#include <initializer_list>
#include <set>
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
#include "io/ooxml/part_dom.h"
#include "io/xml_utils.h"
#include "sheet.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "workbook.h"

namespace formulon {
namespace chart_edit {

namespace {

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

/// Gives bar / area series `to` the fill of bar / area series `from` as it is, as Excel changes between those types;
/// `to` keeps its automatic fill when `from` states none.
void carry_series_fill(const xml_node& from, xml_node to) {
  const auto fill_of = [](const xml_node& sp) {
    for (xml_node e = sp.first_child(); e; e = e.next_sibling()) {
      if (rank(kFills, local_name(e)) >= 0) {
        return e;
      }
    }
    return xml_node();
  };
  const xml_node src = fill_of(child_local(from, "spPr"));
  xml_node sp = child_local(to, "spPr");
  if (!src || !sp) {
    return;
  }
  while (const xml_node old = fill_of(sp)) {
    sp.remove_child(old);
  }
  copy_at(sp, src, kShapeOrder);
}

/// Replaces the `a:ln` of shape properties `sp` with a copy of `ln`.
void replace_ln(xml_node sp, const xml_node& ln) {
  if (!sp || !ln) {
    return;
  }
  drop(sp, "ln");
  copy_at(sp, ln, kShapeOrder);
}

/// The width of the line template's automatic series line.
constexpr std::string_view kAutoLineWidth = "28575";

/// Whether two subtrees match by local names, attribute values and element children, walked without recursion.
bool same_tree(const xml_node& a, const xml_node& b) {
  std::vector<std::pair<xml_node, xml_node>> todo{{a, b}};
  while (!todo.empty()) {
    const auto [x, y] = todo.back();
    todo.pop_back();
    if (local_name(x) != local_name(y)) {
      return false;
    }
    pugi::xml_attribute ax = x.first_attribute();
    pugi::xml_attribute ay = y.first_attribute();
    for (; ax && ay; ax = ax.next_attribute(), ay = ay.next_attribute()) {
      if (std::string_view(ax.name()) != ay.name() || std::string_view(ax.value()) != ay.value()) {
        return false;
      }
    }
    if (ax || ay) {
      return false;
    }
    const auto next = [](xml_node n) {
      do {
        n = n.next_sibling();
      } while (n && n.type() != pugi::node_element);
      return n;
    };
    xml_node cx = io::first_element_child(x);
    xml_node cy = io::first_element_child(y);
    for (; cx && cy; cx = next(cx), cy = next(cy)) {
      todo.emplace_back(cx, cy);
    }
    if (cx || cy) {
      return false;
    }
  }
  return true;
}

/// Replaces the automatic outline `ln` of bar / area shape properties `sp` with line `src` less what the automatic
/// line sets: its cap and join, its width, and its colour when that is the series' automatic one.
void outline_from_line(xml_node sp, xml_node ln, const xml_node& src) {
  xml_node out = sp.insert_copy_after(src, ln);
  sp.remove_child(ln);
  out.remove_attribute("cap");
  for (const char* join : {"round", "bevel", "miter"}) {
    drop(out, join);
  }
  if (std::string_view(out.attribute("w").value()) == kAutoLineWidth) {
    out.remove_attribute("w");
  }
  const xml_node colour = child_local(out, "solidFill");
  const xml_node automatic = child_local(sp, "solidFill");
  if (colour && automatic && same_tree(colour, automatic)) {
    out.insert_child_before("a:noFill", colour);
    out.remove_child(colour);
  }
}

/// Replaces the automatic line `ln` of a line series with visible outline `src`, drawn with the automatic line's
/// round cap and join and, when it states none, its width.
void line_from_outline(xml_node ln, const xml_node& src) {
  xml_node sp = ln.parent();
  xml_node out = sp.insert_copy_after(src, ln);
  sp.remove_child(ln);
  if (!out.attribute("w")) {
    out.prepend_attribute("w").set_value(std::string(kAutoLineWidth).c_str());
  }
  if (!out.attribute("cap")) {
    out.insert_attribute_after("cap", out.attribute("w"));
  }
  out.attribute("cap").set_value("rnd");
  for (const char* join : {"bevel", "miter"}) {
    drop(out, join);
  }
  ensure(out, "a:round", kLineOrder);
}

/// Sets `to`'s series line as Excel changes type. Between line and bar / area types an explicit line becomes the
/// outline and a visible outline the line; otherwise a bar or area series keeps the width of a hidden source line
/// and has none, and a line series from a scatter plot takes 1.5 pt.
void carry_series_line(const xml_node& from, xml_node to, int old_kind, int new_kind) {
  xml_node ln = child_local(child_local(to, "spPr"), "ln");
  if (!ln) {
    return;
  }
  const xml_node src = child_local(child_local(from, "spPr"), "ln");
  const Family before = family_of(old_kind);
  const Family after = family_of(new_kind);
  if (after == Family::kFill && before == Family::kLine) {
    if (src) {
      outline_from_line(child_local(to, "spPr"), ln, src);
    }
  } else if (after == Family::kFill) {
    ln.remove_attribute("w");
    if (src && child_local(src, "noFill") && src.attribute("w")) {
      ln.prepend_attribute("w").set_value(src.attribute("w").value());
    }
  } else if (after == Family::kLine && before == Family::kFill) {
    if (src && child_local(src, "solidFill")) {
      line_from_outline(ln, src);
    }
  } else if (new_kind == static_cast<int>(ChartInsertKind::kLine) &&
             old_kind == static_cast<int>(ChartInsertKind::kScatter)) {
    io::set_attr(ln, "w", "19050");
  }
}

/// Gives pie or doughnut series `to` and its points the outline Excel derives from bar, area, line or scatter series
/// `from` in a chart with style parts: `from`'s line less what the old type's style sets (a hidden outline; the
/// automatic line's width, cap, join and colour), on the series when anything is left, and on every point, hidden
/// when it states no fill.
void ring_outline(const xml_node& from, xml_node to, Family before, const Palette& palette) {
  const xml_node src = child_local(child_local(from, "spPr"), "ln");
  if (!src) {
    return;
  }
  pugi::xml_document doc;
  xml_node ln = doc.append_copy(src);
  if (before == Family::kFill) {
    drop(ln, "noFill");
  } else {
    const std::string_view width = before == Family::kLine ? kAutoLineWidth : std::string_view("19050");
    if (std::string_view(ln.attribute("w").value()) == width) {
      ln.remove_attribute("w");
    }
    if (std::string_view(ln.attribute("cap").value()) == "rnd") {
      ln.remove_attribute("cap");
    }
    drop(ln, "round");
    pugi::xml_document frag;
    const xml_node automatic =
        fragment(frag, "<a:solidFill>" + chart_cycle_color(palette.colors.document_element(), val_u(from, "idx")) +
                           "</a:solidFill>");
    const xml_node colour = child_local(ln, "solidFill");
    if (colour && automatic && same_tree(colour, automatic)) {
      ln.remove_child(colour);
    }
  }
  drop(to, "spPr");
  if (ln.first_attribute() || io::first_element_child(ln)) {
    ensure(to, "c:spPr", kSerOrder).append_copy(ln);
  }
  bool filled = false;
  for (const char* f : {"noFill", "solidFill", "gradFill", "pattFill"}) {
    filled = filled || child_local(ln, f);
  }
  if (!filled) {
    ln.prepend_child("a:noFill");
  }
  for (const xml_node& d : children_local(to, "dPt")) {
    replace_ln(child_local(d, "spPr"), ln);
  }
}

/// Carries `from`'s points to `to` as Excel changes type: a line or scatter point takes the series' marker and
/// line; a bar or area point keeps its shape properties, or from a line or scatter plot takes the series fill
/// and keeps its own line.
void convert_points(const xml_node& from, xml_node to, Family old_family, Family new_family, bool bar) {
  const xml_node series_sp = child_local(to, "spPr");
  for (const xml_node& d : children_local(from, "dPt")) {
    const std::uint32_t idx = val_u(d, "idx");
    pugi::xml_document frag;
    xml_node p = fragment(frag, data_point_open(idx) + (bar ? "<c:invertIfNegative val=\"0\"/>" : "") +
                                    "<c:bubble3D val=\"0\"/></c:dPt>");
    if (new_family != Family::kFill) {
      if (const xml_node marker = child_local(to, "marker")) {
        copy_at(p, marker, kPointOrder);
      }
      xml_node sp = p.append_child("c:spPr");
      if (const xml_node ln = child_local(series_sp, "ln")) {
        sp.append_copy(ln);
      }
      sp.append_child("a:effectLst");
    } else if (old_family == Family::kFill) {
      if (const xml_node sp = child_local(d, "spPr")) {
        copy_at(p, sp, kPointOrder);
      }
    } else if (old_family != Family::kLine || child_local(d, "spPr")) {
      // A line point without shape properties of its own stays bare.
      xml_node sp = p.append_child("c:spPr");
      for (xml_node e = series_sp.first_child(); e; e = e.next_sibling()) {
        if (rank(kFills, local_name(e)) >= 0) {
          sp.append_copy(e);
        }
      }
      const xml_node own = child_local(child_local(d, "spPr"), "ln");
      if (const xml_node ln = own ? own : child_local(series_sp, "ln")) {
        sp.append_copy(ln);
      }
      sp.append_child("a:effectLst");
    }
    if (const xml_node ext = child_local(d, "extLst")) {
      copy_at(p, ext, kPointOrder);
    }
    insert_point(to, p, idx);
  }
}

/// Hides the line of scatter shape properties `sp` as Excel does in a chart without style parts: 2 pt wide when the
/// line was visible with a width (for a point, its own and the series line both), 3 pt otherwise.
void hide_scatter_line(xml_node sp, bool sized) {
  drop(sp, "ln");
  xml_node ln = ensure(sp, "a:ln", kShapeOrder);
  ln.append_attribute("w").set_value(sized ? "25400" : "38100");
  ln.append_child("a:noFill");
}

/// Carries `from`'s formatting to `to` as Excel changes type in a chart without style parts: the shape properties
/// of the series and its points stay as they are, except that a line series without a visible series line loses
/// every line (and series shape properties left with neither a line nor effects) and takes no markers, and a
/// scatter series hides its lines and keeps only a shown marker.
void carry_unstyled(const xml_node& from, xml_node to, Family after, bool bar) {
  drop(to, "spPr");
  drop(to, "dPt");
  if (after != Family::kLine) {
    drop(to, "marker");
  }
  const xml_node series_ln = child_local(child_local(from, "spPr"), "ln");
  const bool shown = series_ln && !child_local(series_ln, "noFill");
  const bool sized = shown && series_ln.attribute("w");
  if (const xml_node sp = child_local(from, "spPr")) {
    copy_at(to, sp, kSerOrder);
  }
  if (after == Family::kLine) {
    if (xml_node sp = child_local(to, "spPr")) {
      if (!shown) {
        drop(sp, "ln");
      }
      if (!child_local(sp, "ln") && !child_local(sp, "effectLst")) {
        to.remove_child(sp);
      }
    }
  } else if (after == Family::kMarker) {
    hide_scatter_line(ensure(to, "c:spPr", kSerOrder), sized);
    const xml_node m = child_local(from, "marker");
    if (m && std::string_view(child_local(m, "symbol").attribute("val").value()) != "none") {
      xml_node out = copy_at(to, m, kSerOrder);
      if (!child_local(out, "size")) {
        put_val(out, "c:size", kMarkerOrder, "5");
      }
    }
  }
  for (const xml_node& d : children_local(from, "dPt")) {
    const std::uint32_t idx = val_u(d, "idx");
    const xml_node invert = child_local(d, "invertIfNegative");
    const std::string invert_val = invert ? invert.attribute("val").value() : "0";
    pugi::xml_document frag;
    xml_node p = fragment(frag, data_point_open(idx) +
                                    (bar ? "<c:invertIfNegative val=\"" + invert_val + "\"/>" : std::string()) +
                                    "<c:bubble3D val=\"0\"/></c:dPt>");
    if (after == Family::kLine || after == Family::kMarker) {
      if (const xml_node marker = child_local(d, "marker")) {
        copy_at(p, marker, kPointOrder);
      }
    }
    if (const xml_node sp = child_local(d, "spPr")) {
      xml_node out = copy_at(p, sp, kPointOrder);
      if (after == Family::kMarker) {
        const xml_node own = child_local(sp, "ln");
        hide_scatter_line(out, sized && own && !child_local(own, "noFill") && own.attribute("w"));
      } else if (after == Family::kLine && !shown) {
        drop(out, "ln");
      }
    }
    if (const xml_node ext = child_local(d, "extLst")) {
      copy_at(p, ext, kPointOrder);
    }
    insert_point(to, p, idx);
  }
}

/// The number format code of `ser`'s category data, as an axis linked to it shows: the cache's, General for text.
std::string category_format(const xml_node& ser) {
  const xml_node code = child_local(child_local(child_local(category_input(ser), "numRef"), "numCache"), "formatCode");
  const std::string text = code.child_value();
  return text.empty() ? "General" : text;
}

/// The base unit Excel gives a date axis over `ser`'s categories.
const char* base_time_unit(const xml_node& ser) {
  std::vector<double> dates;
  const xml_node cache = child_local(child_local(category_input(ser), "numRef"), "numCache");
  for (const xml_node& pt : children_local(cache, "pt")) {
    dates.push_back(child_local(pt, "v").text().as_double());
  }
  return chart_base_time_unit(dates);
}

/// Turns category or date axis `ax` into a value axis or back, as Excel changes type, keeping its identity, scaling
/// direction, gridlines, title, ticks and formatting. Label and mark intervals become the major and minor units; a
/// date axis keeps its units and loses its time units and bounds; a value axis loses its units, display units and
/// bounds, and its log base too on a bar plot. A value axis becomes a date axis with base unit `time_unit` when that
/// is set, else a category axis that is not multi-level, taking `tmpl`'s category settings. The number format links
/// to the category data, of format `format`.
void convert_axis(xml_node ax, const xml_node& tmpl, bool bar, const std::string& format, const char* time_unit) {
  const std::string name = ax.name();
  const std::string prefix = name.substr(0, name.find(':') + 1);
  if (local_name(tmpl) == "valAx") {
    const std::string major = child_local(ax, "tickLblSkip").attribute("val").value();
    const std::string minor = child_local(ax, "tickMarkSkip").attribute("val").value();
    for (const char* e : {"auto", "lblAlgn", "lblOffset", "baseTimeUnit", "majorTimeUnit", "minorTimeUnit",
                          "tickLblSkip", "tickMarkSkip", "noMultiLvlLbl"}) {
      drop(ax, e);
    }
    xml_node scaling = child_local(ax, "scaling");
    drop(scaling, "max");
    drop(scaling, "min");
    ax.set_name((prefix + "valAx").c_str());
    if (!major.empty()) {
      put_val(ax, "c:majorUnit", kAxisOrder, major);
    }
    if (!minor.empty()) {
      put_val(ax, "c:minorUnit", kAxisOrder, minor);
    }
  } else {
    for (const char* e : {"crossBetween", "majorUnit", "minorUnit", "dispUnits"}) {
      drop(ax, e);
    }
    xml_node scaling = child_local(ax, "scaling");
    drop(scaling, "max");
    drop(scaling, "min");
    if (bar) {
      drop(scaling, "logBase");
    }
    ax.set_name((prefix + (time_unit != nullptr ? "dateAx" : "catAx")).c_str());
    for (const char* e : {"auto", "lblAlgn", "lblOffset"}) {
      const xml_node v = child_local(tmpl, e);
      if (v && (time_unit == nullptr || std::string_view(e) != "lblAlgn")) {
        put_val(ax, (std::string("c:") + e).c_str(), kAxisOrder, v.attribute("val").value());
      }
    }
    if (time_unit != nullptr) {
      put_val(ax, "c:baseTimeUnit", kAxisOrder, time_unit);
    } else {
      put_val(ax, "c:noMultiLvlLbl", kAxisOrder, "1");
    }
  }
  drop(ax, "numFmt");
  xml_node fmt = ensure(ax, "c:numFmt", kAxisOrder);
  fmt.append_attribute("formatCode").set_value(format.c_str());
  fmt.append_attribute("sourceLinked").set_value("1");
}

}  // namespace

std::vector<xml_node> axes_of(const Chart& c, const std::vector<std::uint32_t>& ids) {
  std::vector<xml_node> out;
  for (const std::uint32_t id : ids) {
    for (xml_node n = c.area().first_child(); n; n = n.next_sibling()) {
      const std::string_view name = local_name(n);
      if (name.size() > 2 && strings::ends_with(name, "Ax") && val_u(n, "axId") == id) {
        out.push_back(n);
        break;
      }
    }
  }
  return out;
}

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
  if (!valid_kind_grouping(kind, grouping)) {
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
  for (const xml_node& a : children_local(plot, "axId")) {
    ids.push_back(a.attribute("val").as_uint());
  }
  for (const xml_node& other : plots_of(c)) {
    for (const xml_node& a : children_local(other, "axId")) {
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
  const std::vector<xml_node> sers = children_local(plot, "ser");
  ChartSplit split;
  for (const xml_node& s : sers) {
    const xml_node cat = category_input(s);
    const xml_node val = value_input(s);
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
  RETURN_IF_ERROR(template_doc(wb, c, kind, grouping, std::nullopt, ChartLegendPosition::kRight, split, t));
  xml_node tplot = template_plot(t);
  // A chart without style parts keeps its own formatting through a type change, as Excel does.
  Palette palette;
  load_palette(wb, c.path, palette);
  const std::vector<xml_node> fresh = children_local(tplot, "ser");
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
    const xml_node cat = category_input(from);
    const xml_node val = value_input(from);
    if (cat) {
      copy_at(to, cat, kSerOrder, scatter ? "c:xVal" : "c:cat");
    }
    copy_at(to, val, kSerOrder, scatter ? "c:yVal" : "c:val");
    if (const xml_node ext = child_local(from, "extLst")) {
      copy_at(to, ext, kSerOrder);
    }
    if (k == static_cast<int>(ChartInsertKind::kLine) && old == static_cast<int>(ChartInsertKind::kArea)) {
      put_val(to, "c:smooth", kSerOrder, "1");
    }
    if (!palette.styled && !(old_ring && new_ring)) {
      carry_unstyled(from, to, family_of(k), is_bar(tplot));
      continue;
    }
    // As Excel changes type: the series fill carries only between bar and area types; every other change
    // re-derives the series colour.
    if (family_of(old) == Family::kFill && family_of(k) == Family::kFill) {
      carry_series_fill(from, to);
    }
    if (!old_ring) {
      carry_series_line(from, to, old, k);
    }
    // As Excel changes type: line formatting follows the colours between series and points; pie and doughnut swap
    // points as-is.
    if (old_ring && new_ring) {
      drop(to, "dPt");
      for (const xml_node& d : children_local(from, "dPt")) {
        copy_at(to, d, kSerOrder);
      }
    } else if (old_ring) {
      if (family_of(k) == Family::kFill) {
        replace_ln(child_local(to, "spPr"), child_local(child_local(child_local(from, "dPt"), "spPr"), "ln"));
      }
      convert_points(from, to, Family::kNone, family_of(k), is_bar(tplot));
    } else if (new_ring) {
      ring_outline(from, to, family_of(old), palette);
    } else {
      convert_points(from, to, family_of(old), family_of(k), is_bar(tplot));
    }
  }

  // Axes: kept, converted in place between category and value kinds, else replaced by the template's.
  const std::vector<xml_node> old_axes = axes_of(c, ids);
  std::vector<xml_node> new_axes;
  for (xml_node n = child_local(template_chart(t), "plotArea").first_child(); n; n = n.next_sibling()) {
    if (local_name(n) == "catAx" || local_name(n) == "valAx" || local_name(n) == "dateAx") {
      new_axes.push_back(n);
    }
  }
  bool same = old_axes.size() == new_axes.size() && !old_axes.empty();
  for (std::size_t i = 0; same && i < old_axes.size(); ++i) {
    const std::string_view name = local_name(old_axes[i]);
    same = name == "catAx" || name == "valAx" || name == "dateAx";
  }
  xml_node area = c.area();
  std::vector<std::uint32_t> new_ids;
  if (same) {
    const std::string format = sers.empty() ? std::string("General") : category_format(sers[0]);
    // A value axis over dates becomes a date axis, as Excel judges the category data.
    const char* time_unit = chart_is_date_format(format) ? base_time_unit(sers[0]) : nullptr;
    for (std::size_t i = 0; i < old_axes.size(); ++i) {
      if ((local_name(old_axes[i]) == "valAx") != (local_name(new_axes[i]) == "valAx")) {
        convert_axis(old_axes[i], new_axes[i], is_bar(tplot), format, time_unit);
      } else {
        // A linked value-axis format follows the grouping, as Excel re-derives it: 0% for percent-stacked.
        xml_node fmt = child_local(old_axes[i], "numFmt");
        if (local_name(old_axes[i]) == "valAx" && fmt && flag_attr(fmt, "sourceLinked") &&
            (grouping_of(plot) == ChartGrouping::kPercentStacked) != (grouping == ChartGrouping::kPercentStacked)) {
          io::set_attr(fmt, "formatCode", child_local(new_axes[i], "numFmt").attribute("formatCode").value());
        }
      }
      // As Excel changes type: a category axis loses its label and mark intervals.
      drop(old_axes[i], "tickLblSkip");
      drop(old_axes[i], "tickMarkSkip");
      put_val(old_axes[i], "c:axPos", kAxisOrder, child_local(new_axes[i], "axPos").attribute("val").value());
      if (const xml_node between = child_local(new_axes[i], "crossBetween")) {
        put_val(old_axes[i], "c:crossBetween", kAxisOrder, between.attribute("val").value());
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
    // As Excel changes type from pie or doughnut: new axes take outside ticks and the default crossing, a bar
    // plot lists its value axis first, a scatter plot's axes have no outline and its horizontal axis no gridlines
    // or number format, and a chart without style parts leaves them unformatted.
    const bool horizontal = std::string_view(child_local(tplot, "barDir").attribute("val").value()) == "bar";
    std::vector<xml_node> made(new_axes.size());
    for (std::size_t n = 0; n < new_axes.size(); ++n) {
      const std::size_t i = old_ring && horizontal ? new_axes.size() - 1 - n : n;
      xml_node a = before ? area.insert_copy_before(new_axes[i], before) : area.append_copy(new_axes[i]);
      put_val(a, "c:axId", kAxisOrder, std::to_string(new_ids[i]));
      put_val(a, "c:crossAx", kAxisOrder, std::to_string(new_ids[1 - i]));
      made[i] = a;
    }
    for (std::size_t i = 0; old_ring && i < made.size(); ++i) {
      xml_node a = made[i];
      put_val(a, "c:majorTickMark", kAxisOrder, "out");
      drop(a, "crosses");
      if (scatter) {
        if (i == 0) {
          drop(a, "majorGridlines");
          drop(a, "numFmt");
        }
        if (xml_node sp = child_local(a, "spPr")) {
          pugi::xml_document frag;
          replace_ln(sp, fragment(frag, "<a:ln><a:noFill/></a:ln>"));
        }
      }
      if (!palette.styled) {
        drop(a, "spPr");
        drop(a, "txPr");
        drop(child_local(a, "majorGridlines"), "spPr");
      }
    }
  }
  const std::vector<xml_node> axis_refs = children_local(tplot, "axId");
  for (std::size_t i = 0; i < axis_refs.size() && i < new_ids.size(); ++i) {
    io::set_attr(axis_refs[i], "val", std::to_string(new_ids[i]));
  }
  // As Excel changes type: a new bar plot takes the schema default gap, 100 from pie or doughnut; a bar plot keeps
  // its own.
  if (is_bar(tplot)) {
    const xml_node gap = child_local(plot, "gapWidth");
    const char* fresh_gap = old_ring ? "100" : "150";
    put_val(tplot, "c:gapWidth", kPlotOrder, is_bar(plot) && gap ? gap.attribute("val").value() : fresh_gap);
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

}  // namespace chart_edit
}  // namespace formulon
