#include "auto_filter.h"

#include <array>
#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/xml_utils.h"
#include "io/xsd_double.h"
#include "pugixml.hpp"
#include "utils/a1_column.h"
#include "utils/a1_ref.h"

namespace formulon {

namespace {

constexpr std::string_view kRichFilterExtUri = "{1AD28BCE-077C-4C59-8B6E-1921CE8616D4}";
constexpr std::string_view kRichDataNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2";

constexpr std::array<std::string_view, 6> kGroupingNames = {"year", "month", "day", "hour", "minute", "second"};
constexpr std::array<std::string_view, 6> kOperatorNames = {"equal",    "lessThan",           "lessThanOrEqual",
                                                            "notEqual", "greaterThanOrEqual", "greaterThan"};
constexpr std::array<std::string_view, 35> kDynamicNames = {
    "null",        "aboveAverage", "belowAverage", "tomorrow",  "today",      "yesterday",   "nextWeek",
    "thisWeek",    "lastWeek",     "nextMonth",    "thisMonth", "lastMonth",  "nextQuarter", "thisQuarter",
    "lastQuarter", "nextYear",     "thisYear",     "lastYear",  "yearToDate", "Q1",          "Q2",
    "Q3",          "Q4",           "M1",           "M2",        "M3",         "M4",          "M5",
    "M6",          "M7",           "M8",           "M9",        "M10",        "M11",         "M12"};
constexpr std::array<std::string_view, 4> kSortByNames = {"value", "cellColor", "fontColor", "icon"};
constexpr std::array<std::string_view, 3> kSortMethodNames = {"none", "pinYin", "stroke"};

template <std::size_t N>
std::optional<std::size_t> lookup(const std::array<std::string_view, N>& names, std::string_view text) {
  for (std::size_t i = 0; i < N; ++i) {
    if (names[i] == text) {
      return i;
    }
  }
  return std::nullopt;
}

/// Parse state: `ok` turns false on the first construct the model cannot carry.
struct Reader {
  bool ok = true;

  void fail() { ok = false; }

  /// Rejects any attribute not in `known`.
  void only(const pugi::xml_node& node, std::initializer_list<std::string_view> known) {
    for (const pugi::xml_attribute& attr : node.attributes()) {
      bool listed = false;
      for (const std::string_view name : known) {
        listed = listed || name == attr.name();
      }
      if (!listed) {
        fail();
      }
    }
  }

  /// Retains every attribute not in `known`, verbatim and in order.
  static RawAttributes extras(const pugi::xml_node& node, std::initializer_list<std::string_view> known) {
    RawAttributes out;
    for (const pugi::xml_attribute& attr : node.attributes()) {
      bool listed = false;
      for (const std::string_view name : known) {
        listed = listed || name == attr.name();
      }
      if (!listed) {
        out.emplace_back(attr.name(), attr.value());
      }
    }
    return out;
  }

  bool boolean(const pugi::xml_node& node, const char* name, bool def) {
    const pugi::xml_attribute attr = node.attribute(name);
    if (!attr) {
      return def;
    }
    const std::string_view v = attr.value();
    if (v == "1" || v == "true") {
      return true;
    }
    if (v == "0" || v == "false") {
      return false;
    }
    fail();
    return def;
  }

  std::optional<std::uint32_t> u32(const pugi::xml_node& node, const char* name) {
    const pugi::xml_attribute attr = node.attribute(name);
    if (!attr) {
      return std::nullopt;
    }
    const std::optional<std::uint32_t> v = io::parse_xml_u32_attr_strict(attr);
    if (!v) {
      fail();
    }
    return v;
  }

  std::optional<double> number(const pugi::xml_node& node, const char* name) {
    const pugi::xml_attribute attr = node.attribute(name);
    if (!attr) {
      return std::nullopt;
    }
    double v = 0.0;
    if (!io::parse_xsd_double(attr.value(), &v)) {
      fail();
      return std::nullopt;
    }
    return v;
  }

  template <std::size_t N>
  std::size_t choice(const pugi::xml_node& node, const char* name, const std::array<std::string_view, N>& names,
                     std::size_t def) {
    const pugi::xml_attribute attr = node.attribute(name);
    if (!attr) {
      return def;
    }
    const std::optional<std::size_t> v = lookup(names, attr.value());
    if (!v) {
      fail();
      return def;
    }
    return *v;
  }

  MergeRange rect(const pugi::xml_node& node) {
    const std::optional<MergeRange> r = parse_a1_rectangle(node.attribute("ref").value());
    if (!r) {
      fail();
      return {};
    }
    return *r;
  }
};

bool is_element(const pugi::xml_node& node) {
  return node.type() == pugi::node_element;
}

/// Parses one `<dateGroupItem>` (any prefix); fields present must be exactly
/// those its grouping uses.
std::optional<DateGroupItem> read_date_group(Reader& r, const pugi::xml_node& node) {
  r.only(node, {"year", "month", "day", "hour", "minute", "second", "dateTimeGrouping"});
  const std::optional<std::size_t> grouping = lookup(kGroupingNames, node.attribute("dateTimeGrouping").value());
  if (!grouping) {
    r.fail();
    return std::nullopt;
  }
  constexpr std::array<const char*, 6> kFields = {"year", "month", "day", "hour", "minute", "second"};
  constexpr std::array<std::uint32_t, 6> kMax = {9999U, 12U, 31U, 23U, 59U, 59U};
  std::array<std::uint32_t, 6> fields{};
  for (std::size_t i = 0; i < kFields.size(); ++i) {
    const std::optional<std::uint32_t> v = r.u32(node, kFields[i]);
    if (v.has_value() != (i <= *grouping) || (v && *v > kMax[i])) {
      r.fail();
      return std::nullopt;
    }
    fields[i] = v.value_or(0U);
  }
  DateGroupItem item;
  item.year = static_cast<std::uint16_t>(fields[0]);
  item.month = static_cast<std::uint8_t>(fields[1]);
  item.day = static_cast<std::uint8_t>(fields[2]);
  item.hour = static_cast<std::uint8_t>(fields[3]);
  item.minute = static_cast<std::uint8_t>(fields[4]);
  item.second = static_cast<std::uint8_t>(fields[5]);
  item.grouping = static_cast<DateTimeGrouping>(*grouping);
  return item;
}

/// Parses `<filters>` children named `<prefix>filter` / `<prefix>dateGroupItem`.
void read_value_filters(Reader& r, const pugi::xml_node& node, std::string_view prefix, ValueFilters& out) {
  r.only(node, {"blank", "calendarType"});
  out.blank = r.boolean(node, "blank", false);
  out.calendar_type = node.attribute("calendarType").value();
  if (out.calendar_type == "none") {
    out.calendar_type.clear();
  }
  for (const pugi::xml_node& child : node.children()) {
    if (!is_element(child)) {
      continue;
    }
    const std::string_view name = child.name();
    if (name.substr(0, prefix.size()) != prefix) {
      r.fail();
      return;
    }
    const std::string_view local = name.substr(prefix.size());
    if (local == "filter") {
      r.only(child, {"val"});
      out.values.emplace_back(child.attribute("val").value());
    } else if (local == "dateGroupItem") {
      if (std::optional<DateGroupItem> item = read_date_group(r, child)) {
        out.date_groups.push_back(*item);
      }
    } else {
      r.fail();
      return;
    }
  }
}

/// Lifts the date-group extension Excel writes into `out`. Returns false
/// (leaving `out` untouched) when `ext` is any other extension or carries
/// content the model does not name.
bool lift_rich_filter_ext(const pugi::xml_node& ext, ValueFilters& out) {
  if (std::string_view(ext.attribute("uri").value()) != kRichFilterExtUri) {
    return false;
  }
  std::string prefix;
  for (const pugi::xml_attribute& attr : ext.attributes()) {
    const std::string_view name = attr.name();
    if (name == "uri") {
      continue;
    }
    if (name.substr(0, 6) != "xmlns:" || std::string_view(attr.value()) != kRichDataNamespace || !prefix.empty()) {
      return false;
    }
    prefix = std::string(name.substr(6)) + ":";
  }
  if (prefix.empty()) {
    return false;
  }
  const pugi::xml_node column = ext.first_child();
  if (!is_element(column) || column.next_sibling() || column.first_attribute() ||
      std::string_view(column.name()) != prefix + "filterColumn") {
    return false;
  }
  const pugi::xml_node filters = column.first_child();
  if (!is_element(filters) || filters.next_sibling() || std::string_view(filters.name()) != prefix + "filters") {
    return false;
  }
  Reader r;
  ValueFilters lifted;
  read_value_filters(r, filters, prefix, lifted);
  if (!r.ok || lifted.date_groups.empty()) {
    return false;
  }
  out = std::move(lifted);
  return true;
}

void read_filter_column(Reader& r, const pugi::xml_node& node, FilterColumn& col) {
  const std::optional<std::uint32_t> col_id = r.u32(node, "colId");
  if (!col_id) {
    r.fail();
    return;
  }
  col.col_id = *col_id;
  col.hidden_button = r.boolean(node, "hiddenButton", false);
  col.show_button = r.boolean(node, "showButton", true);
  col.extra_attrs = Reader::extras(node, {"colId", "hiddenButton", "showButton"});
  for (const pugi::xml_node& child : node.children()) {
    if (!is_element(child)) {
      continue;
    }
    const std::string_view name = child.name();
    const bool criterion = name == "filters" || name == "customFilters" || name == "top10" || name == "dynamicFilter" ||
                           name == "colorFilter" || name == "iconFilter";
    if (criterion && col.kind != FilterKind::kNone) {
      r.fail();
      return;
    }
    if (name == "filters") {
      col.kind = FilterKind::kValues;
      read_value_filters(r, child, "", col.values);
    } else if (name == "customFilters") {
      col.kind = FilterKind::kCustom;
      r.only(child, {"and"});
      col.custom.and_join = r.boolean(child, "and", false);
      for (const pugi::xml_node& item : child.children()) {
        if (!is_element(item)) {
          continue;
        }
        if (std::string_view(item.name()) != "customFilter") {
          r.fail();
          return;
        }
        r.only(item, {"operator", "val"});
        CustomFilter filter;
        filter.op = static_cast<FilterOperator>(r.choice(item, "operator", kOperatorNames, 0U));
        filter.val = item.attribute("val").value();
        col.custom.filters.push_back(std::move(filter));
      }
      if (col.custom.filters.empty() || col.custom.filters.size() > 2U) {
        r.fail();
      }
    } else if (name == "top10") {
      col.kind = FilterKind::kTop10;
      r.only(child, {"top", "percent", "val", "filterVal"});
      col.top10.top = r.boolean(child, "top", true);
      col.top10.percent = r.boolean(child, "percent", false);
      const std::optional<double> val = r.number(child, "val");
      if (!val) {
        r.fail();
      }
      col.top10.val = val.value_or(0.0);
      col.top10.filter_val = r.number(child, "filterVal");
    } else if (name == "dynamicFilter") {
      col.kind = FilterKind::kDynamic;
      r.only(child, {"type", "val", "valIso", "maxVal", "maxValIso"});
      if (!child.attribute("type")) {
        r.fail();
      }
      col.dynamic.type = static_cast<DynamicFilterType>(r.choice(child, "type", kDynamicNames, 0U));
      col.dynamic.val = r.number(child, "val");
      col.dynamic.val_iso = child.attribute("valIso").value();
      col.dynamic.max_val = r.number(child, "maxVal");
      col.dynamic.max_val_iso = child.attribute("maxValIso").value();
    } else if (name == "colorFilter") {
      col.kind = FilterKind::kColor;
      r.only(child, {"dxfId", "cellColor"});
      col.color.dxf_id = r.u32(child, "dxfId");
      col.color.cell_color = r.boolean(child, "cellColor", true);
    } else if (name == "iconFilter") {
      col.kind = FilterKind::kIcon;
      r.only(child, {"iconSet", "iconId"});
      if (!child.attribute("iconSet")) {
        r.fail();
      }
      col.icon.icon_set = static_cast<cf::IconSetName>(r.choice(child, "iconSet", cf::kIconSetNames, 0U));
      col.icon.icon_id = r.u32(child, "iconId");
    } else if (name == "extLst") {
      for (const pugi::xml_node& ext : child.children()) {
        if (!is_element(ext)) {
          continue;
        }
        ValueFilters lifted;
        if (col.kind == FilterKind::kNone && lift_rich_filter_ext(ext, lifted)) {
          col.kind = FilterKind::kValues;
          col.values = std::move(lifted);
        } else {
          io::append_raw_xml(col.ext_xml, ext);
        }
      }
    } else {
      io::append_raw_xml(col.extra_xml, child);
    }
    if (!r.ok) {
      return;
    }
  }
}

void read_sort_state(Reader& r, const pugi::xml_node& node, SortState& sort) {
  sort.ref = r.rect(node);
  sort.column_sort = r.boolean(node, "columnSort", false);
  sort.case_sensitive = r.boolean(node, "caseSensitive", false);
  sort.sort_method = static_cast<SortMethod>(r.choice(node, "sortMethod", kSortMethodNames, 0U));
  sort.extra_attrs = Reader::extras(node, {"ref", "columnSort", "caseSensitive", "sortMethod"});
  for (const pugi::xml_node& child : node.children()) {
    if (!is_element(child)) {
      continue;
    }
    const std::string_view name = child.name();
    if (name == "sortCondition") {
      r.only(child, {"descending", "sortBy", "ref", "customList", "dxfId", "iconSet", "iconId"});
      SortCondition cond;
      cond.ref = r.rect(child);
      cond.descending = r.boolean(child, "descending", false);
      cond.sort_by = static_cast<SortBy>(r.choice(child, "sortBy", kSortByNames, 0U));
      cond.custom_list = child.attribute("customList").value();
      cond.dxf_id = r.u32(child, "dxfId");
      cond.icon_set = static_cast<cf::IconSetName>(r.choice(child, "iconSet", cf::kIconSetNames, 0U));
      cond.icon_id = r.u32(child, "iconId");
      sort.conditions.push_back(std::move(cond));
    } else if (name == "extLst" && sort.ext_lst_xml.empty()) {
      sort.ext_lst_xml = io::raw_xml(child);
    } else {
      r.fail();
    }
  }
}

void append_bool(std::string& out, const char* name, bool value, bool def) {
  if (value != def) {
    io::append_xml_attr(out, name, value ? "1" : "0");
  }
}

void append_number(std::string& out, const char* name, double value) {
  out.push_back(' ');
  out.append(name);
  out.append("=\"");
  io::append_xml_number(out, value);
  out.push_back('"');
}

void append_ref(std::string& out, const MergeRange& rect) {
  std::string ref;
  append_a1_rectangle(ref, rect);
  io::append_xml_attr(out, "ref", ref);
}

void append_value_filters(std::string& out, const ValueFilters& values, std::string_view prefix) {
  out.push_back('<');
  out.append(prefix);
  out.append("filters");
  append_bool(out, "blank", values.blank, false);
  if (!values.calendar_type.empty()) {
    io::append_xml_attr(out, "calendarType", values.calendar_type);
  }
  if (values.values.empty() && values.date_groups.empty()) {
    out.append("/>");
    return;
  }
  out.push_back('>');
  for (const std::string& value : values.values) {
    out.push_back('<');
    out.append(prefix);
    out.append("filter");
    io::append_xml_attr(out, "val", value);
    out.append("/>");
  }
  for (const DateGroupItem& item : values.date_groups) {
    out.push_back('<');
    out.append(prefix);
    out.append("dateGroupItem");
    const std::array<std::uint32_t, 6> fields = {item.year, item.month, item.day, item.hour, item.minute, item.second};
    constexpr std::array<const char*, 6> kFields = {"year", "month", "day", "hour", "minute", "second"};
    for (std::size_t i = 0; i <= static_cast<std::size_t>(item.grouping); ++i) {
      io::append_xml_attr_uint(out, kFields[i], fields[i]);
    }
    io::append_xml_attr(out, "dateTimeGrouping", kGroupingNames[static_cast<std::size_t>(item.grouping)]);
    out.append("/>");
  }
  out.append("</");
  out.append(prefix);
  out.append("filters>");
}

void append_filter_column(std::string& out, const FilterColumn& col) {
  out.append("<filterColumn");
  io::append_xml_attr_uint(out, "colId", col.col_id);
  append_bool(out, "hiddenButton", col.hidden_button, false);
  append_bool(out, "showButton", col.show_button, true);
  io::append_raw_attrs(out, col.extra_attrs);
  const bool rich_dates = col.kind == FilterKind::kValues && !col.values.date_groups.empty();
  std::string body;
  switch (col.kind) {
    case FilterKind::kNone:
      break;
    case FilterKind::kValues:
      if (!rich_dates) {
        append_value_filters(body, col.values, "");
      }
      break;
    case FilterKind::kCustom:
      body.append("<customFilters");
      append_bool(body, "and", col.custom.and_join, false);
      body.push_back('>');
      for (const CustomFilter& filter : col.custom.filters) {
        body.append("<customFilter");
        if (filter.op != FilterOperator::kEqual) {
          io::append_xml_attr(body, "operator", kOperatorNames[static_cast<std::size_t>(filter.op)]);
        }
        io::append_xml_attr(body, "val", filter.val);
        body.append("/>");
      }
      body.append("</customFilters>");
      break;
    case FilterKind::kTop10:
      body.append("<top10");
      append_bool(body, "top", col.top10.top, true);
      append_bool(body, "percent", col.top10.percent, false);
      append_number(body, "val", col.top10.val);
      if (col.top10.filter_val) {
        append_number(body, "filterVal", *col.top10.filter_val);
      }
      body.append("/>");
      break;
    case FilterKind::kDynamic:
      body.append("<dynamicFilter");
      io::append_xml_attr(body, "type", kDynamicNames[static_cast<std::size_t>(col.dynamic.type)]);
      if (col.dynamic.val) {
        append_number(body, "val", *col.dynamic.val);
      }
      if (!col.dynamic.val_iso.empty()) {
        io::append_xml_attr(body, "valIso", col.dynamic.val_iso);
      }
      if (col.dynamic.max_val) {
        append_number(body, "maxVal", *col.dynamic.max_val);
      }
      if (!col.dynamic.max_val_iso.empty()) {
        io::append_xml_attr(body, "maxValIso", col.dynamic.max_val_iso);
      }
      body.append("/>");
      break;
    case FilterKind::kColor:
      body.append("<colorFilter");
      if (col.color.dxf_id) {
        io::append_xml_attr_uint(body, "dxfId", *col.color.dxf_id);
      }
      append_bool(body, "cellColor", col.color.cell_color, true);
      body.append("/>");
      break;
    case FilterKind::kIcon:
      body.append("<iconFilter");
      io::append_xml_attr(body, "iconSet", cf::kIconSetNames[static_cast<std::size_t>(col.icon.icon_set)]);
      if (col.icon.icon_id) {
        io::append_xml_attr_uint(body, "iconId", *col.icon.icon_id);
      }
      body.append("/>");
      break;
  }
  body.append(col.extra_xml);
  if (rich_dates || !col.ext_xml.empty()) {
    body.append("<extLst>");
    if (rich_dates) {
      body.append("<ext");
      io::append_xml_attr(body, "uri", kRichFilterExtUri);
      io::append_xml_attr(body, "xmlns:xlrd2", kRichDataNamespace);
      body.append("><xlrd2:filterColumn>");
      append_value_filters(body, col.values, "xlrd2:");
      body.append("</xlrd2:filterColumn></ext>");
    }
    body.append(col.ext_xml);
    body.append("</extLst>");
  }
  if (body.empty()) {
    out.append("/>");
    return;
  }
  out.push_back('>');
  out.append(body);
  out.append("</filterColumn>");
}

void append_sort_state(std::string& out, const SortState& sort) {
  out.append("<sortState");
  append_bool(out, "columnSort", sort.column_sort, false);
  append_bool(out, "caseSensitive", sort.case_sensitive, false);
  if (sort.sort_method != SortMethod::kNone) {
    io::append_xml_attr(out, "sortMethod", kSortMethodNames[static_cast<std::size_t>(sort.sort_method)]);
  }
  io::append_raw_attrs(out, sort.extra_attrs);
  append_ref(out, sort.ref);
  if (sort.conditions.empty() && sort.ext_lst_xml.empty()) {
    out.append("/>");
    return;
  }
  out.push_back('>');
  for (const SortCondition& cond : sort.conditions) {
    out.append("<sortCondition");
    append_bool(out, "descending", cond.descending, false);
    if (cond.sort_by != SortBy::kValue) {
      io::append_xml_attr(out, "sortBy", kSortByNames[static_cast<std::size_t>(cond.sort_by)]);
    }
    append_ref(out, cond.ref);
    if (!cond.custom_list.empty()) {
      io::append_xml_attr(out, "customList", cond.custom_list);
    }
    if (cond.dxf_id) {
      io::append_xml_attr_uint(out, "dxfId", *cond.dxf_id);
    }
    if (cond.sort_by == SortBy::kIcon || cond.icon_set != cf::IconSetName::Three_Arrows) {
      io::append_xml_attr(out, "iconSet", cf::kIconSetNames[static_cast<std::size_t>(cond.icon_set)]);
    }
    if (cond.icon_id) {
      io::append_xml_attr_uint(out, "iconId", *cond.icon_id);
    }
    out.append("/>");
  }
  out.append(sort.ext_lst_xml);
  out.append("</sortState>");
}

Error invalid(std::string message) {
  return make_error(FormulonErrorCode::kAutoFilterInvalid, std::move(message));
}

}  // namespace

bool AutoFilter::has_criteria() const noexcept {
  for (const FilterColumn& col : columns) {
    if (col.kind != FilterKind::kNone) {
      return true;
    }
  }
  return false;
}

std::optional<MergeRange> parse_a1_rectangle(std::string_view text) noexcept {
  const std::size_t colon = text.find(':');
  MergeRange rect;
  if (!a1::parse_a1_ref(text.substr(0, colon), &rect.first_row, &rect.first_col)) {
    return std::nullopt;
  }
  if (colon == std::string_view::npos) {
    rect.last_row = rect.first_row;
    rect.last_col = rect.first_col;
  } else if (!a1::parse_a1_ref(text.substr(colon + 1U), &rect.last_row, &rect.last_col)) {
    return std::nullopt;
  }
  if (rect.first_row > rect.last_row || rect.first_col > rect.last_col) {
    return std::nullopt;
  }
  return rect;
}

bool append_a1_rectangle(std::string& out, const MergeRange& rect) {
  const auto append_cell = [&out](std::uint32_t row, std::uint32_t col) {
    if (!a1::append_column_letters(out, col)) {
      return false;
    }
    out += std::to_string(static_cast<std::uint64_t>(row) + 1U);
    return true;
  };
  if (!append_cell(rect.first_row, rect.first_col)) {
    return false;
  }
  if (rect.first_row == rect.last_row && rect.first_col == rect.last_col) {
    return true;
  }
  out += ':';
  return append_cell(rect.last_row, rect.last_col);
}

Expected<AutoFilter, Error> parse_auto_filter_xml(std::string_view xml) {
  pugi::xml_document doc;
  const pugi::xml_parse_result parsed =
      doc.load_buffer(xml.data(), xml.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parsed) {
    return invalid(std::string("autoFilter: XML parse failed: ") + parsed.description());
  }
  const pugi::xml_node root = doc.first_child();
  if (!is_element(root) || root.next_sibling() || std::string_view(root.name()) != "autoFilter") {
    return invalid("autoFilter: fragment is not a single autoFilter element");
  }
  Reader r;
  AutoFilter filter;
  filter.range = r.rect(root);
  filter.extra_attrs = Reader::extras(root, {"ref"});
  for (const pugi::xml_node& child : root.children()) {
    if (!is_element(child) || !r.ok) {
      continue;
    }
    const std::string_view name = child.name();
    if (name == "filterColumn" && !filter.sort && filter.ext_lst_xml.empty()) {
      filter.columns.emplace_back();
      read_filter_column(r, child, filter.columns.back());
    } else if (name == "sortState" && !filter.sort && filter.ext_lst_xml.empty()) {
      filter.sort.emplace();
      read_sort_state(r, child, *filter.sort);
    } else if (name == "extLst" && filter.ext_lst_xml.empty()) {
      filter.ext_lst_xml = io::raw_xml(child);
    } else {
      r.fail();
    }
  }
  if (!r.ok) {
    return invalid("autoFilter: element carries content the model cannot represent");
  }
  return filter;
}

std::string serialize_auto_filter(const AutoFilter& filter) {
  if (filter.is_opaque()) {
    return filter.opaque_xml;
  }
  std::string out = "<autoFilter";
  append_ref(out, filter.range);
  io::append_raw_attrs(out, filter.extra_attrs);
  if (filter.columns.empty() && !filter.sort && filter.ext_lst_xml.empty()) {
    out.append("/>");
    return out;
  }
  out.push_back('>');
  for (const FilterColumn& col : filter.columns) {
    append_filter_column(out, col);
  }
  if (filter.sort) {
    append_sort_state(out, *filter.sort);
  }
  out.append(filter.ext_lst_xml);
  out.append("</autoFilter>");
  return out;
}

Expected<void, Error> validate_auto_filter(const AutoFilter& filter) {
  if (filter.is_opaque()) {
    return invalid("autoFilter: model holds an unparsed fragment");
  }
  std::string probe;
  if (filter.range.first_row > filter.range.last_row || filter.range.first_col > filter.range.last_col ||
      !append_a1_rectangle(probe, filter.range) || !parse_a1_rectangle(probe)) {
    return invalid("autoFilter: range is not an in-grid rectangle");
  }
  const std::uint64_t width = static_cast<std::uint64_t>(filter.range.last_col) - filter.range.first_col + 1U;
  for (std::size_t i = 0; i < filter.columns.size(); ++i) {
    const FilterColumn& col = filter.columns[i];
    if (col.col_id >= width || (i > 0U && col.col_id <= filter.columns[i - 1U].col_id)) {
      return invalid("autoFilter: column ids must be ascending and inside the range");
    }
    if (col.kind == FilterKind::kCustom && (col.custom.filters.empty() || col.custom.filters.size() > 2U)) {
      return invalid("autoFilter: a custom filter holds one or two conditions");
    }
    for (const DateGroupItem& item : col.values.date_groups) {
      const auto level = static_cast<std::size_t>(item.grouping);
      if (level >= kGroupingNames.size() || item.year > 9999U ||
          (level >= 1U && (item.month < 1U || item.month > 12U)) ||
          (level >= 2U && (item.day < 1U || item.day > 31U)) || item.hour > 23U || item.minute > 59U ||
          item.second > 59U) {
        return invalid("autoFilter: date group item fields do not fit its grouping");
      }
    }
  }
  return Expected<void, Error>::Ok();
}

void AutoFilterSlot::set_xml(std::string_view xml) {
  if (xml.empty()) {
    value_.reset();
    return;
  }
  Expected<AutoFilter, Error> parsed = parse_auto_filter_xml(xml);
  if (parsed) {
    value_ = std::move(parsed.value());
    return;
  }
  AutoFilter opaque;
  opaque.opaque_xml.assign(xml);
  value_ = std::move(opaque);
}

}  // namespace formulon
