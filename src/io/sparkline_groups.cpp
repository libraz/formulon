#include "io/sparkline_groups.h"

#include <algorithm>
#include <cmath>
#include <cstring>

#include "io/color_spec_xml.h"
#include "io/sqref_text.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "utils/number_text.h"
#include "utils/status_macros.h"

namespace formulon::io {
namespace {

constexpr const char* kSparklineUri = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";
constexpr const char* kNsX14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
constexpr const char* kNsXm = "http://schemas.microsoft.com/office/excel/2006/main";
constexpr const char* kNsXr2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2";
constexpr std::size_t kMaxItems = 1U << 16;

constexpr const char* kTypeNames[] = {"line", "column", "stacked"};
constexpr const char* kAxisNames[] = {"individual", "group", "custom"};
constexpr const char* kEmptyNames[] = {"gap", "zero", "span"};

// Element-name helpers that ignore the prefix, since Excel and other writers bind the x14 and xm namespaces to
// different prefixes.
std::string_view local_name(const char* qname) {
  const char* colon = std::strchr(qname, ':');
  return colon != nullptr ? colon + 1 : qname;
}

std::string_view local_name(const pugi::xml_node& node) {
  return local_name(node.name());
}

std::vector<pugi::xml_node> children_local(const pugi::xml_node& parent, std::string_view name) {
  std::vector<pugi::xml_node> out;
  for (pugi::xml_node child = parent.first_child(); child; child = child.next_sibling()) {
    if (child.type() == pugi::node_element && local_name(child) == name) {
      out.push_back(child);
    }
  }
  return out;
}

pugi::xml_node child_local(const pugi::xml_node& parent, std::string_view name) {
  const std::vector<pugi::xml_node> found = children_local(parent, name);
  return found.empty() ? pugi::xml_node() : found.front();
}

// Child order of `CT_SparklineGroup`; anything else sorts last.
constexpr std::string_view kChildOrder[] = {"colorSeries", "colorNegative", "colorAxis", "colorMarkers",
                                            "colorFirst",  "colorLast",     "colorHigh", "colorLow",
                                            "f",           "sparklines"};

std::size_t child_rank(pugi::xml_node node) {
  const std::string_view name = local_name(node);
  return static_cast<std::size_t>(std::find(std::begin(kChildOrder), std::end(kChildOrder), name) -
                                  std::begin(kChildOrder));
}

Error invalid(const char* message) {
  return make_error(FormulonErrorCode::kInvalidArgument, message);
}

// `<x14:sparklineGroup>` nodes of `doc` in document order.
std::vector<pugi::xml_node> group_nodes(const pugi::xml_document& doc) {
  std::vector<pugi::xml_node> groups;
  for (pugi::xml_node ext : children_local(doc.document_element(), "ext")) {
    if (std::string_view(ext.attribute("uri").value()) != kSparklineUri) {
      continue;
    }
    for (pugi::xml_node list : children_local(ext, "sparklineGroups")) {
      for (pugi::xml_node group : children_local(list, "sparklineGroup")) {
        groups.push_back(group);
      }
    }
  }
  return groups;
}

template <std::size_t N>
std::int32_t enum_attr(pugi::xml_node node, const char* name, const char* const (&names)[N]) {
  const std::string_view value = node.attribute(name).value();
  for (std::size_t i = 0; i < N; ++i) {
    if (value == names[i]) {
      return static_cast<std::int32_t>(i);
    }
  }
  return 0;
}

ColorSpec read_color(pugi::xml_node group, std::string_view name) {
  return read_color_spec(child_local(group, name));
}

SparklineGroup read_group(pugi::xml_node node) {
  SparklineGroup g;
  g.type = static_cast<SparklineType>(enum_attr(node, "type", kTypeNames));
  g.color_series = read_color(node, "colorSeries");
  g.color_negative = read_color(node, "colorNegative");
  g.color_axis = read_color(node, "colorAxis");
  g.color_markers = read_color(node, "colorMarkers");
  g.color_first = read_color(node, "colorFirst");
  g.color_last = read_color(node, "colorLast");
  g.color_high = read_color(node, "colorHigh");
  g.color_low = read_color(node, "colorLow");
  g.markers = attr_bool(node, "markers");
  g.high = attr_bool(node, "high");
  g.low = attr_bool(node, "low");
  g.first = attr_bool(node, "first");
  g.last = attr_bool(node, "last");
  g.negative = attr_bool(node, "negative");
  g.display_x_axis = attr_bool(node, "displayXAxis");
  g.display_hidden = attr_bool(node, "displayHidden");
  g.min_axis_type = static_cast<SparklineAxisType>(enum_attr(node, "minAxisType", kAxisNames));
  g.max_axis_type = static_cast<SparklineAxisType>(enum_attr(node, "maxAxisType", kAxisNames));
  g.has_manual_min = node.attribute("manualMin") && std::isfinite(attr_f64(node, "manualMin", NAN));
  g.has_manual_max = node.attribute("manualMax") && std::isfinite(attr_f64(node, "manualMax", NAN));
  g.manual_min = g.has_manual_min ? attr_f64(node, "manualMin") : 0.0;
  g.manual_max = g.has_manual_max ? attr_f64(node, "manualMax") : 0.0;
  g.date_axis = attr_bool(node, "dateAxis");
  g.date_axis_ref = child_local(node, "f").text().get();
  g.line_weight = attr_f64(node, "lineWeight", 0.75);
  g.empty_cells_as = static_cast<SparklineEmptyCells>(enum_attr(node, "displayEmptyCellsAs", kEmptyNames));
  for (pugi::xml_node item : children_local(child_local(node, "sparklines"), "sparkline")) {
    g.items.push_back({child_local(item, "f").text().get(), child_local(item, "sqref").text().get()});
  }
  g.right_to_left = attr_bool(node, "rightToLeft");
  for (pugi::xml_attribute a = node.first_attribute(); a; a = a.next_attribute()) {
    if (local_name(a.name()) == "uid") {
      g.uid = a.value();
    }
  }
  return g;
}

bool load(pugi::xml_document& doc, const std::string& xml) {
  return !xml.empty() && doc.load_string(xml.c_str()) && local_name(doc.document_element()) == "extLst";
}

void store(std::string& xml, const pugi::xml_document& doc) {
  pugi::xml_node root = doc.document_element();
  xml = has_element_child(root) ? raw_xml(root) : std::string();
}

// Sets (or removes, for the default) a boolean attribute when its value changes.
void set_flag(pugi::xml_node node, const char* name, bool value) {
  if (attr_bool(node, name) == value) {
    return;
  }
  if (value) {
    (node.attribute(name) ? node.attribute(name) : node.append_attribute(name)).set_value("1");
  } else {
    node.remove_attribute(name);
  }
}

void set_text_attr(pugi::xml_node node, const char* name, const std::string& value, bool is_default) {
  if (is_default && !node.attribute(name)) {
    return;
  }
  (node.attribute(name) ? node.attribute(name) : node.append_attribute(name)).set_value(value.c_str());
}

std::string number_text(double value) {
  std::string out;
  append_xml_number(out, value);
  return out;
}

// Inserts a new child named `qname` before the first sibling of a later rank.
pugi::xml_node insert_ranked(pugi::xml_node group, const std::string& qname) {
  const std::size_t rank = static_cast<std::size_t>(
      std::find(std::begin(kChildOrder), std::end(kChildOrder), local_name(qname.c_str())) - std::begin(kChildOrder));
  for (pugi::xml_node child = group.first_child(); child; child = child.next_sibling()) {
    if (child.type() == pugi::node_element && child_rank(child) > rank) {
      return group.insert_child_before(qname.c_str(), child);
    }
  }
  return group.append_child(qname.c_str());
}

void set_color(pugi::xml_node group, const std::string& prefix, const char* name, const ColorSpec& spec) {
  pugi::xml_node node = child_local(group, name);
  if (spec.kind == ColorSpec::Kind::kNone) {
    group.remove_child(node);
    return;
  }
  if (!node) {
    node = insert_ranked(group, prefix + name);
  }
  while (node.first_attribute()) {
    node.remove_attribute(node.first_attribute());
  }
  switch (spec.kind) {
    case ColorSpec::Kind::kTheme:
      node.append_attribute("theme").set_value(spec.theme);
      if (spec.tint != 0.0) {
        node.append_attribute("tint").set_value(number_text(spec.tint).c_str());
      }
      break;
    case ColorSpec::Kind::kIndexed:
      node.append_attribute("indexed").set_value(spec.indexed);
      break;
    case ColorSpec::Kind::kAuto:
      node.append_attribute("auto").set_value("1");
      break;
    default: {
      char hex[16];
      format_hex(hex, sizeof hex, spec.rgb, 8, true);
      node.append_attribute("rgb").set_value(hex);
    }
  }
}

Expected<void, Error> validate_items(const std::vector<SparklineItem>& items) {
  if (items.empty() || items.size() > kMaxItems) {
    return invalid("sparkline group: needs between 1 and 65536 sparklines");
  }
  for (const SparklineItem& item : items) {
    std::vector<MergeRange> cells;
    if (item.formula.empty() || !parse_sqref_text(item.location, cells) || cells.size() != 1U ||
        cells.front().first_row != cells.front().last_row || cells.front().first_col != cells.front().last_col) {
      return invalid("sparkline group: each sparkline needs a formula and one location cell");
    }
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> validate_group(const SparklineGroup& g) {
  const auto bad_enum = [](std::int32_t v) { return v < 0 || v > 2; };
  if (bad_enum(static_cast<std::int32_t>(g.type)) || bad_enum(static_cast<std::int32_t>(g.min_axis_type)) ||
      bad_enum(static_cast<std::int32_t>(g.max_axis_type)) || bad_enum(static_cast<std::int32_t>(g.empty_cells_as))) {
    return invalid("sparkline group: enumeration out of range");
  }
  if (!std::isfinite(g.manual_min) || !std::isfinite(g.manual_max) || !std::isfinite(g.line_weight) ||
      g.line_weight < 0.0) {
    return invalid("sparkline group: number is not finite or the line weight is negative");
  }
  return Expected<void, Error>::Ok();
}

void apply_group(pugi::xml_node node, const SparklineGroup& g) {
  const std::string px = std::string(node.name()).substr(0, std::string(node.name()).find(':') + 1U);
  const std::string xm = [&node] {
    pugi::xml_node f = child_local(child_local(child_local(node, "sparklines"), "sparkline"), "f");
    const std::string_view name = f.name();
    return f ? std::string(name.substr(0, name.find(':') + 1U)) : std::string("xm:");
  }();
  set_text_attr(node, "type", kTypeNames[static_cast<int>(g.type)], g.type == SparklineType::kLine);
  set_text_attr(node, "displayEmptyCellsAs", kEmptyNames[static_cast<int>(g.empty_cells_as)],
                g.empty_cells_as == SparklineEmptyCells::kGap);
  set_text_attr(node, "minAxisType", kAxisNames[static_cast<int>(g.min_axis_type)],
                g.min_axis_type == SparklineAxisType::kIndividual);
  set_text_attr(node, "maxAxisType", kAxisNames[static_cast<int>(g.max_axis_type)],
                g.max_axis_type == SparklineAxisType::kIndividual);
  if (std::abs(attr_f64(node, "lineWeight", 0.75) - g.line_weight) > 0.0) {
    node.remove_attribute("lineWeight");
    node.append_attribute("lineWeight").set_value(number_text(g.line_weight).c_str());
  }
  const struct {
    const char* name;
    bool has;
    double value;
  } manual[] = {{"manualMin", g.has_manual_min, g.manual_min}, {"manualMax", g.has_manual_max, g.manual_max}};
  for (const auto& m : manual) {
    node.remove_attribute(m.name);
    if (m.has) {
      node.append_attribute(m.name).set_value(number_text(m.value).c_str());
    }
  }
  const struct {
    const char* name;
    bool value;
  } flags[] = {{"markers", g.markers},
               {"high", g.high},
               {"low", g.low},
               {"first", g.first},
               {"last", g.last},
               {"negative", g.negative},
               {"displayXAxis", g.display_x_axis},
               {"displayHidden", g.display_hidden},
               {"dateAxis", g.date_axis}};
  for (const auto& flag : flags) {
    set_flag(node, flag.name, flag.value);
  }
  const struct {
    const char* name;
    const ColorSpec* spec;
  } colors[] = {{"colorSeries", &g.color_series}, {"colorNegative", &g.color_negative},
                {"colorAxis", &g.color_axis},     {"colorMarkers", &g.color_markers},
                {"colorFirst", &g.color_first},   {"colorLast", &g.color_last},
                {"colorHigh", &g.color_high},     {"colorLow", &g.color_low}};
  for (const auto& color : colors) {
    set_color(node, px, color.name, *color.spec);
  }
  pugi::xml_node ref = child_local(node, "f");
  if (g.date_axis_ref.empty()) {
    node.remove_child(ref);
  } else {
    (ref ? ref : insert_ranked(node, xm + "f")).text().set(g.date_axis_ref.c_str());
  }
}

void apply_items(pugi::xml_node group, const std::vector<SparklineItem>& items) {
  const std::string px = std::string(group.name()).substr(0, std::string(group.name()).find(':') + 1U);
  pugi::xml_node list = child_local(group, "sparklines");
  std::string xm = "xm:";
  if (list) {
    const std::string_view name = child_local(child_local(list, "sparkline"), "f").name();
    xm = name.empty() ? xm : std::string(name.substr(0, name.find(':') + 1U));
    while (list.first_child()) {
      list.remove_child(list.first_child());
    }
  } else {
    list = insert_ranked(group, px + "sparklines");
  }
  for (const SparklineItem& item : items) {
    std::vector<MergeRange> cells;
    parse_sqref_text(item.location, cells);
    pugi::xml_node node = list.append_child((px + "sparkline").c_str());
    node.append_child((xm + "f").c_str()).text().set(item.formula.c_str());
    node.append_child((xm + "sqref").c_str()).text().set(format_sqref_text(cells).c_str());
  }
}

pugi::xml_node ensure_groups_list(pugi::xml_document& doc) {
  pugi::xml_node root = doc.document_element();
  if (!root) {
    root = doc.append_child("extLst");
  }
  for (pugi::xml_node ext : children_local(root, "ext")) {
    if (std::string_view(ext.attribute("uri").value()) == kSparklineUri) {
      if (pugi::xml_node list = child_local(ext, "sparklineGroups")) {
        return list;
      }
    }
  }
  pugi::xml_node ext = root.append_child("ext");
  ext.append_attribute("uri").set_value(kSparklineUri);
  ext.append_attribute("xmlns:x14").set_value(kNsX14);
  pugi::xml_node list = ext.append_child("x14:sparklineGroups");
  list.append_attribute("xmlns:xm").set_value(kNsXm);
  return list;
}

Expected<pugi::xml_node, Error> group_at(const pugi::xml_document& doc, std::size_t index) {
  const std::vector<pugi::xml_node> groups = group_nodes(doc);
  if (index >= groups.size()) {
    return invalid("sparkline group: index out of range");
  }
  return groups[index];
}

}  // namespace

std::vector<SparklineGroup> read_sparkline_groups(std::string_view ext_lst_xml) {
  std::vector<SparklineGroup> out;
  pugi::xml_document doc;
  if (ext_lst_xml.find("sparklineGroup") == std::string_view::npos || !load(doc, std::string(ext_lst_xml))) {
    return out;
  }
  for (pugi::xml_node node : group_nodes(doc)) {
    out.push_back(read_group(node));
  }
  return out;
}

Expected<void, Error> add_sparkline_group(std::string& ext_lst_xml, const SparklineGroup& group) {
  RETURN_IF_ERROR(validate_group(group));
  RETURN_IF_ERROR(validate_items(group.items));
  pugi::xml_document doc;
  if (!ext_lst_xml.empty() && !load(doc, ext_lst_xml)) {
    return invalid("sparkline group: the worksheet extension list does not parse");
  }
  pugi::xml_node list = ensure_groups_list(doc);
  const std::string px = std::string(list.name()).substr(0, std::string(list.name()).find(':') + 1U);
  pugi::xml_node node = list.append_child((px + "sparklineGroup").c_str());
  apply_group(node, group);
  set_flag(node, "rightToLeft", group.right_to_left);
  if (!group.uid.empty()) {
    node.append_attribute("xmlns:xr2").set_value(kNsXr2);
    node.append_attribute("xr2:uid").set_value(group.uid.c_str());
  }
  apply_items(node, group.items);
  store(ext_lst_xml, doc);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> update_sparkline_group(std::string& ext_lst_xml, std::size_t index, const SparklineGroup& group) {
  RETURN_IF_ERROR(validate_group(group));
  pugi::xml_document doc;
  if (!load(doc, ext_lst_xml)) {
    return invalid("sparkline group: index out of range");
  }
  ASSIGN_OR_RETURN(pugi::xml_node node, group_at(doc, index));
  apply_group(node, group);
  store(ext_lst_xml, doc);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> set_group_sparklines(std::string& ext_lst_xml, std::size_t index,
                                           const std::vector<SparklineItem>& items) {
  RETURN_IF_ERROR(validate_items(items));
  pugi::xml_document doc;
  if (!load(doc, ext_lst_xml)) {
    return invalid("sparkline group: index out of range");
  }
  ASSIGN_OR_RETURN(pugi::xml_node node, group_at(doc, index));
  apply_items(node, items);
  store(ext_lst_xml, doc);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> remove_sparkline_group(std::string& ext_lst_xml, std::size_t index) {
  pugi::xml_document doc;
  if (!load(doc, ext_lst_xml)) {
    return invalid("sparkline group: index out of range");
  }
  ASSIGN_OR_RETURN(pugi::xml_node node, group_at(doc, index));
  // Climb while the parent would be left without element children.
  pugi::xml_node doomed = node;
  while (doomed.parent() != doc.document_element()) {
    pugi::xml_node parent = doomed.parent();
    parent.remove_child(doomed);
    if (has_element_child(parent)) {
      store(ext_lst_xml, doc);
      return Expected<void, Error>::Ok();
    }
    doomed = parent;
  }
  doc.document_element().remove_child(doomed);
  store(ext_lst_xml, doc);
  return Expected<void, Error>::Ok();
}

}  // namespace formulon::io
