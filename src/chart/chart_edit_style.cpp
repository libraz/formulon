#include <algorithm>
#include <iterator>
#include <optional>
#include <string>

#include "chart/chart_edit_internal.h"
#include "chart/chart_ns.h"
#include "chart/chart_template.h"
#include "drawing/drawing_color.h"
#include "drawing/drawing_xml.h"
#include "eval/function_registry.h"
#include "io/xml_utils.h"
#include "sheet.h"
#include "utils/status_macros.h"
#include "workbook.h"

namespace formulon {
namespace chart_edit {

// ---------------------------------------------------------------------------
// Shape properties, markers, labels
// ---------------------------------------------------------------------------

namespace {

/// Removes every fill child of `node`.
void drop_fills(xml_node node) {
  for (Order f = kFills; *f != nullptr; ++f) {
    drop(node, *f);
  }
}

/// The `c:spPr` of `owner`, created with the `a:effectLst` Excel writes in every shape-properties element.
xml_node ensure_shape_properties(xml_node owner, Order order) {
  const bool fresh = !child_local(owner, "spPr");
  xml_node sp = ensure(owner, "c:spPr", order);
  if (fresh) {
    sp.append_child("a:effectLst");
  }
  return sp;
}

/// Applies a hidden / shown flag to `holder`: hidden swaps the fills for `a:noFill`, shown removes it.
Expected<void, Error> apply_hidden(xml_node holder, Order order, const std::optional<bool>& none,
                                   const ChartPatchField<ColorSpec>& color, const char* colour_message) {
  if (!none.has_value()) {
    return ok();
  }
  if (*none && color.is_set()) {
    return bad_arg(colour_message);
  }
  if (*none) {
    drop_fills(holder);
    ensure(holder, "a:noFill", order);
  } else {
    drop(holder, "noFill");
  }
  return ok();
}

/// Removes every child of `node` not named `keep`.
void remove_children_except(xml_node node, std::string_view keep) {
  for (xml_node n = node.first_child(); n;) {
    const xml_node next = n.next_sibling();
    if (local_name(n) != keep) {
      node.remove_child(n);
    }
    n = next;
  }
}

/// Colour and transparency of `holder`'s solid fill (`a:spPr`, `a:ln` or `a:defRPr`).
Expected<void, Error> apply_paint(xml_node holder, Order order, const ChartPatchField<ColorSpec>& color,
                                  const ChartPatchField<double>& transparency) {
  if (color.is_clear()) {
    drop(holder, "solidFill");
  } else if (color.is_set()) {
    const xml_node current = io::first_element_child(child_local(holder, "solidFill"));
    // An unchanged colour keeps its stored element (`lt1` stays `lt1` rather than becoming `bg1`).
    const auto stored = drawing_color::read_exact(current, false);
    if (!stored || !drawing_color::same_color(*stored, color.value)) {
      const std::string alpha = drawing_color::alpha_of(current);
      drop_fills(holder);
      drawing_color::write_color(ensure(holder, "a:solidFill", order), "a:", color.value, alpha);
    }
  }
  if (transparency.is_keep()) {
    return ok();
  }
  xml_node clr = io::first_element_child(child_local(holder, "solidFill"));
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
  io::set_attr(alpha ? alpha : clr.append_child("a:alpha"), "val", rounded((100.0 - transparency.value) * 1000.0));
  return ok();
}

}  // namespace

Expected<void, Error> apply_paint_properties(xml_node sp, const ChartFillPatch& fill, const ChartLinePatch& line) {
  RETURN_IF_ERROR(apply_hidden(sp, kShapeOrder, fill.none, fill.color, "chart: a hidden fill takes no colour"));
  RETURN_IF_ERROR(apply_paint(sp, kShapeOrder, fill.color, fill.transparency_pct));
  if (line_empty(line)) {
    return ok();
  }
  xml_node ln = ensure(sp, "a:ln", kShapeOrder);
  RETURN_IF_ERROR(apply_hidden(ln, kLineOrder, line.none, line.color, "chart: a hidden line takes no colour"));
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
    const int d = static_cast<int>(line.dash.value);
    if (!chart_token_valid(kChartDashTokens, d)) {
      return bad_arg("chart: line dash must be solid, dash, dot or dashDot");
    }
    drop(ln, "custDash");
    put_val(ln, "a:prstDash", kLineOrder, kChartDashTokens[d]);
  }
  return ok();
}

Expected<void, Error> apply_shape(xml_node owner, Order owner_order, const ChartFillPatch& fill,
                                  const ChartLinePatch& line) {
  if (fill_empty(fill) && line_empty(line)) {
    return ok();
  }
  return apply_paint_properties(ensure_shape_properties(owner, owner_order), fill, line);
}

void link_number_format(xml_node owner, Order order, const std::optional<bool>& linked) {
  if (!linked.has_value()) {
    return;
  }
  xml_node fmt = child_local(owner, "numFmt");
  if (!fmt) {
    if (!*linked) {
      return;
    }
    fmt = ensure(owner, "c:numFmt", order);
    io::set_attr(fmt, "formatCode", "General");
  }
  io::set_attr(fmt, "sourceLinked", *linked ? "1" : "0");
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
    const int s = static_cast<int>(*m.symbol);
    if (s < 0 || s > static_cast<int>(ChartMarkerSymbol::kDash)) {
      return bad_arg("chart: marker symbol must be a named shape");
    }
    put_val(node, "c:symbol", kMarkerOrder, kChartSymbolTokens[s]);
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

namespace {

// Symbols Excel gives an automatic marker, by series idx.
constexpr ChartMarkerSymbol kAutoSymbols[] = {
    ChartMarkerSymbol::kDiamond, ChartMarkerSymbol::kSquare, ChartMarkerSymbol::kTriangle,
    ChartMarkerSymbol::kX,       ChartMarkerSymbol::kStar,   ChartMarkerSymbol::kCircle,
    ChartMarkerSymbol::kPlus,    ChartMarkerSymbol::kDot,    ChartMarkerSymbol::kDash};

/// Appends the automatic colour of series `idx` to `fill`.
void append_series_color(xml_node fill, const Palette& p, std::uint32_t idx) {
  pugi::xml_document frag;
  if (const xml_node clr = fragment(frag, chart_cycle_color(p.colors.document_element(), idx))) {
    fill.append_copy(clr);
  }
}

}  // namespace

void settle_marker(xml_node ser, xml_node owner, const Palette& p) {
  xml_node m = child_local(owner, "marker");
  if (!m) {
    return;
  }
  const std::uint32_t idx = val_u(ser, "idx");
  std::string symbol = child_local(m, "symbol").attribute("val").value();
  if (symbol.empty() || symbol == "auto") {
    symbol = owner == ser ? "" : child_local(child_local(ser, "marker"), "symbol").attribute("val").value();
    if (symbol.empty() || symbol == "auto") {
      symbol = kChartSymbolTokens[static_cast<int>(kAutoSymbols[idx % std::size(kAutoSymbols)])];
    }
    put_val(m, "c:symbol", kMarkerOrder, symbol);
  }
  if (symbol == "none") {
    remove_children_except(m, "symbol");
    return;
  }
  if (!child_local(m, "size")) {
    put_val(m, "c:size", kMarkerOrder, "5");
  }
  if (!p.styled) {
    return;
  }
  // A styled chart keeps its style part only while every shown marker has both a fill and an outline.
  xml_node sp = ensure_shape_properties(m, kMarkerOrder);
  const auto filled = [](const xml_node& n) {
    return std::any_of(std::begin(kFills), std::end(kFills) - 1,
                       [&](const char* f) { return static_cast<bool>(child_local(n, f)); });
  };
  if (!filled(sp)) {
    append_series_color(ensure(sp, "a:solidFill", kShapeOrder), p, idx);
  }
  xml_node ln = ensure(sp, "a:ln", kShapeOrder);
  if (!ln.attribute("w")) {
    ln.prepend_attribute("w").set_value("9525");
  }
  if (!filled(ln)) {
    append_series_color(ensure(ln, "a:solidFill", kLineOrder), p, idx);
  }
}

namespace {

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
      lang + "\"/></a:p></c:txPr>" + kChartShowFlags;
  const std::string ext = std::string("<c:extLst><c:ext uri=\"") + kUriLeader + "\" xmlns:c15=\"" + kNsC15 + "\"";
  if (pie) {
    x += std::string("<c:showLeaderLines val=\"1\"/><c:leaderLines>") + leader_sp_pr() + "</c:leaderLines>" + ext +
         "/></c:extLst>";
  } else {
    x += "<c:showLeaderLines val=\"0\"/>" + ext + "><c15:showLeaderLines val=\"1\"/><c15:leaderLines>" +
         leader_sp_pr() + "</c15:leaderLines></c:ext></c:extLst>";
  }
  return x + "</c:dLbls>";
}

}  // namespace

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
  if (d && flag(child_local(d, "delete")) && !p.deleted.has_value()) {
    // Deleted labels store no other key, so only an explicit `deleted` changes them.
    const bool touches = p.show_val || p.show_cat || p.show_ser || p.show_pct || p.show_legend_key ||
                         !p.separator.is_keep() || !p.position.is_keep() || !p.number_format.is_keep() ||
                         !p.font_color.is_keep() || !p.font_size_pt.is_keep() || p.source_linked;
    return touches ? bad_arg("chart: the labels are deleted; set deleted to false to change them") : ok();
  }
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
    } else if (!child_local(d, local_name(kShow[i]))) {
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
    const int i = static_cast<int>(p.position.value);
    if (!chart_token_valid(kChartLabelPositionTokens, i)) {
      return bad_arg("chart: unknown label position");
    }
    put_val(d, "c:dLblPos", kLabelsOrder, kChartLabelPositionTokens[i]);
  }
  if (p.number_format.is_clear()) {
    drop(d, "numFmt");
  } else if (p.number_format.is_set()) {
    xml_node fmt = ensure(d, "c:numFmt", kLabelsOrder);
    io::set_attr(fmt, "formatCode", p.number_format.value);
    io::set_attr(fmt, "sourceLinked", "0");
  }
  link_number_format(d, kLabelsOrder, p.source_linked);
  if (!p.font_color.is_keep() || !p.font_size_pt.is_keep()) {
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
      io::set_attr(rpr, "sz", rounded(p.font_size_pt.value * 100.0));
    }
    RETURN_IF_ERROR(apply_paint(rpr, kRunOrder, p.font_color, ChartPatchField<double>{}));
  }
  if (p.deleted.value_or(false)) {
    // CT_DLbls holds either `c:delete` or the label group; only per-label `c:dLbl` entries stay beside it.
    remove_children_except(d, "dLbl");
    put_val(d, "c:delete", kLabelsOrder, "1");
  }
  return ok();
}

}  // namespace chart_edit

Expected<void, Error> write_paint_properties(pugi::xml_node sp_pr, const ChartFillPatch& fill,
                                             const ChartLinePatch& line) {
  return chart_edit::apply_paint_properties(sp_pr, fill, line);
}

}  // namespace formulon
