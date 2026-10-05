#include "auto_filter.h"

#include <array>
#include <cstddef>
#include <cstdint>
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

std::optional<std::size_t> lookup(const std::string_view* names, std::size_t count, std::string_view text) {
  for (std::size_t i = 0; i < count; ++i) {
    if (names[i] == text) {
      return i;
    }
  }
  return std::nullopt;
}

/// How one attribute is parsed and serialized. The target a table entry binds
/// to has the type named in each comment.
enum class AttrKind : std::uint8_t {
  kBool,            ///< `bool`; written when it differs from `bool_default`.
  kU32,             ///< `std::optional<std::uint32_t>`; written when engaged.
  kU32Required,     ///< `std::uint32_t`; absence fails, always written.
  kNumber,          ///< `std::optional<double>`; written when engaged.
  kNumberRequired,  ///< `double`; absence fails, always written.
  kText,            ///< `std::string`; written when non-empty.
  kTextAlways,      ///< `std::string`; always written.
  kChoice,          ///< One-byte enum indexing `choices`; written unless 0.
  kChoiceRequired,  ///< One-byte enum indexing `choices`; absence fails, always written.
  kRect,            ///< `MergeRange` as an A1 `ref`; always written.
  kExtras,          ///< `RawAttributes`: every attribute the table does not name.
};

/// One attribute of an element. A table lists them in serialization order; an
/// element whose table has no `kExtras` entry rejects unlisted attributes.
struct AttrSpec {
  const char* name;
  AttrKind kind;
  bool bool_default = false;
  const std::string_view* choices = nullptr;
  std::size_t choice_count = 0;
};

constexpr AttrSpec choice_attr(const char* name, AttrKind kind, const std::string_view* names, std::size_t count) {
  return AttrSpec{name, kind, false, names, count};
}

constexpr AttrSpec kRootAttrs[] = {{"ref", AttrKind::kRect}, {nullptr, AttrKind::kExtras}};
constexpr AttrSpec kColumnAttrs[] = {{"colId", AttrKind::kU32Required},
                                     {"hiddenButton", AttrKind::kBool, false},
                                     {"showButton", AttrKind::kBool, true},
                                     {nullptr, AttrKind::kExtras}};
constexpr AttrSpec kFiltersAttrs[] = {{"blank", AttrKind::kBool, false}, {"calendarType", AttrKind::kText}};
constexpr AttrSpec kFilterAttrs[] = {{"val", AttrKind::kTextAlways}};
constexpr AttrSpec kDateGroupAttrs[] = {
    {"year", AttrKind::kU32},
    {"month", AttrKind::kU32},
    {"day", AttrKind::kU32},
    {"hour", AttrKind::kU32},
    {"minute", AttrKind::kU32},
    {"second", AttrKind::kU32},
    choice_attr("dateTimeGrouping", AttrKind::kChoiceRequired, kGroupingNames.data(), kGroupingNames.size())};
constexpr AttrSpec kCustomFiltersAttrs[] = {{"and", AttrKind::kBool, false}};
constexpr AttrSpec kCustomFilterAttrs[] = {
    choice_attr("operator", AttrKind::kChoice, kOperatorNames.data(), kOperatorNames.size()),
    {"val", AttrKind::kTextAlways}};
constexpr AttrSpec kTop10Attrs[] = {{"top", AttrKind::kBool, true},
                                    {"percent", AttrKind::kBool, false},
                                    {"val", AttrKind::kNumberRequired},
                                    {"filterVal", AttrKind::kNumber}};
constexpr AttrSpec kDynamicAttrs[] = {
    choice_attr("type", AttrKind::kChoiceRequired, kDynamicNames.data(), kDynamicNames.size()),
    {"val", AttrKind::kNumber},
    {"valIso", AttrKind::kText},
    {"maxVal", AttrKind::kNumber},
    {"maxValIso", AttrKind::kText}};
constexpr AttrSpec kColorAttrs[] = {{"dxfId", AttrKind::kU32}, {"cellColor", AttrKind::kBool, true}};
constexpr AttrSpec kIconAttrs[] = {
    choice_attr("iconSet", AttrKind::kChoiceRequired, cf::kIconSetNames.data(), cf::kIconSetNames.size()),
    {"iconId", AttrKind::kU32}};
constexpr AttrSpec kSortStateAttrs[] = {
    {"columnSort", AttrKind::kBool, false},
    {"caseSensitive", AttrKind::kBool, false},
    choice_attr("sortMethod", AttrKind::kChoice, kSortMethodNames.data(), kSortMethodNames.size()),
    {nullptr, AttrKind::kExtras},
    {"ref", AttrKind::kRect}};
/// `iconSet` (index `kSortConditionIconSet`) is written by hand: an icon sort
/// names its set even when it is the default one.
constexpr AttrSpec kSortConditionAttrs[] = {
    {"descending", AttrKind::kBool, false},
    choice_attr("sortBy", AttrKind::kChoice, kSortByNames.data(), kSortByNames.size()),
    {"ref", AttrKind::kRect},
    {"customList", AttrKind::kText},
    {"dxfId", AttrKind::kU32},
    choice_attr("iconSet", AttrKind::kChoice, cf::kIconSetNames.data(), cf::kIconSetNames.size()),
    {"iconId", AttrKind::kU32}};
constexpr std::size_t kSortConditionIconSet = 5U;

static_assert(sizeof(FilterOperator) == 1U && sizeof(DynamicFilterType) == 1U && sizeof(cf::IconSetName) == 1U &&
                  sizeof(SortBy) == 1U && sizeof(SortMethod) == 1U && sizeof(DateTimeGrouping) == 1U,
              "kChoice targets are stored through one byte");

/// Parse state: `ok` turns false on the first construct the model cannot carry.
struct Reader {
  bool ok = true;

  void fail() { ok = false; }

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

  MergeRange rect(const pugi::xml_node& node) {
    const std::optional<MergeRange> r = parse_a1_rectangle(node.attribute("ref").value());
    if (!r) {
      fail();
      return {};
    }
    return *r;
  }

  /// Reads every attribute `specs` names into the matching `targets` entry.
  void attrs(const pugi::xml_node& node, const AttrSpec* specs, std::size_t n, void* const* targets) {
    RawAttributes* extras = nullptr;
    for (std::size_t i = 0; i < n; ++i) {
      if (specs[i].kind == AttrKind::kExtras) {
        extras = static_cast<RawAttributes*>(targets[i]);
        extras->clear();
      }
    }
    for (const pugi::xml_attribute& attr : node.attributes()) {
      bool listed = false;
      for (std::size_t i = 0; i < n; ++i) {
        listed = listed || (specs[i].name != nullptr && std::string_view(specs[i].name) == attr.name());
      }
      if (listed) {
        continue;
      }
      if (extras != nullptr) {
        extras->emplace_back(attr.name(), attr.value());
      } else {
        fail();
      }
    }
    for (std::size_t i = 0; i < n; ++i) {
      const AttrSpec& spec = specs[i];
      void* target = targets[i];
      switch (spec.kind) {
        case AttrKind::kBool:
          *static_cast<bool*>(target) = boolean(node, spec.name, spec.bool_default);
          break;
        case AttrKind::kU32:
          *static_cast<std::optional<std::uint32_t>*>(target) = u32(node, spec.name);
          break;
        case AttrKind::kU32Required: {
          const std::optional<std::uint32_t> v = u32(node, spec.name);
          if (!v) {
            fail();
          }
          *static_cast<std::uint32_t*>(target) = v.value_or(0U);
          break;
        }
        case AttrKind::kNumber:
          *static_cast<std::optional<double>*>(target) = number(node, spec.name);
          break;
        case AttrKind::kNumberRequired: {
          const std::optional<double> v = number(node, spec.name);
          if (!v) {
            fail();
          }
          *static_cast<double*>(target) = v.value_or(0.0);
          break;
        }
        case AttrKind::kText:
        case AttrKind::kTextAlways:
          *static_cast<std::string*>(target) = node.attribute(spec.name).value();
          break;
        case AttrKind::kChoice:
        case AttrKind::kChoiceRequired: {
          const pugi::xml_attribute attr = node.attribute(spec.name);
          std::size_t index = 0;
          if (attr) {
            const std::optional<std::size_t> v = lookup(spec.choices, spec.choice_count, attr.value());
            if (v) {
              index = *v;
            } else {
              fail();
            }
          } else if (spec.kind == AttrKind::kChoiceRequired) {
            fail();
          }
          *static_cast<std::uint8_t*>(target) = static_cast<std::uint8_t>(index);
          break;
        }
        case AttrKind::kRect:
          *static_cast<MergeRange*>(target) = rect(node);
          break;
        case AttrKind::kExtras:
          break;
      }
    }
  }

  template <std::size_t N>
  void attrs(const pugi::xml_node& node, const AttrSpec (&specs)[N], void* const (&targets)[N]) {
    attrs(node, specs, N, targets);
  }
};

bool is_element(const pugi::xml_node& node) {
  return node.type() == pugi::node_element;
}

/// Parses one `<dateGroupItem>` (any prefix); fields present must be exactly
/// those its grouping uses.
std::optional<DateGroupItem> read_date_group(Reader& r, const pugi::xml_node& node) {
  std::array<std::optional<std::uint32_t>, 6> fields;
  DateGroupItem item;
  r.attrs(node, kDateGroupAttrs,
          {&fields[0], &fields[1], &fields[2], &fields[3], &fields[4], &fields[5], &item.grouping});
  if (!r.ok) {
    return std::nullopt;
  }
  constexpr std::array<std::uint32_t, 6> kMax = {9999U, 12U, 31U, 23U, 59U, 59U};
  for (std::size_t i = 0; i < fields.size(); ++i) {
    if (fields[i].has_value() != (i <= static_cast<std::size_t>(item.grouping)) ||
        (fields[i] && *fields[i] > kMax[i])) {
      r.fail();
      return std::nullopt;
    }
  }
  item.year = static_cast<std::uint16_t>(fields[0].value_or(0U));
  item.month = static_cast<std::uint8_t>(fields[1].value_or(0U));
  item.day = static_cast<std::uint8_t>(fields[2].value_or(0U));
  item.hour = static_cast<std::uint8_t>(fields[3].value_or(0U));
  item.minute = static_cast<std::uint8_t>(fields[4].value_or(0U));
  item.second = static_cast<std::uint8_t>(fields[5].value_or(0U));
  return item;
}

/// Parses `<filters>` children named `<prefix>filter` / `<prefix>dateGroupItem`.
void read_value_filters(Reader& r, const pugi::xml_node& node, std::string_view prefix, ValueFilters& out) {
  r.attrs(node, kFiltersAttrs, {&out.blank, &out.calendar_type});
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
      std::string& value = out.values.emplace_back();
      r.attrs(child, kFilterAttrs, {&value});
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
  r.attrs(node, kColumnAttrs, {&col.col_id, &col.hidden_button, &col.show_button, &col.extra_attrs});
  if (!r.ok) {
    return;
  }
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
      r.attrs(child, kCustomFiltersAttrs, {&col.custom.and_join});
      for (const pugi::xml_node& item : child.children()) {
        if (!is_element(item)) {
          continue;
        }
        if (std::string_view(item.name()) != "customFilter") {
          r.fail();
          return;
        }
        CustomFilter& filter = col.custom.filters.emplace_back();
        r.attrs(item, kCustomFilterAttrs, {&filter.op, &filter.val});
      }
      if (col.custom.filters.empty() || col.custom.filters.size() > 2U) {
        r.fail();
      }
    } else if (name == "top10") {
      col.kind = FilterKind::kTop10;
      r.attrs(child, kTop10Attrs, {&col.top10.top, &col.top10.percent, &col.top10.val, &col.top10.filter_val});
    } else if (name == "dynamicFilter") {
      col.kind = FilterKind::kDynamic;
      r.attrs(
          child, kDynamicAttrs,
          {&col.dynamic.type, &col.dynamic.val, &col.dynamic.val_iso, &col.dynamic.max_val, &col.dynamic.max_val_iso});
    } else if (name == "colorFilter") {
      col.kind = FilterKind::kColor;
      r.attrs(child, kColorAttrs, {&col.color.dxf_id, &col.color.cell_color});
    } else if (name == "iconFilter") {
      col.kind = FilterKind::kIcon;
      r.attrs(child, kIconAttrs, {&col.icon.icon_set, &col.icon.icon_id});
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
  r.attrs(node, kSortStateAttrs,
          {&sort.column_sort, &sort.case_sensitive, &sort.sort_method, &sort.extra_attrs, &sort.ref});
  for (const pugi::xml_node& child : node.children()) {
    if (!is_element(child)) {
      continue;
    }
    const std::string_view name = child.name();
    if (name == "sortCondition") {
      SortCondition& cond = sort.conditions.emplace_back();
      r.attrs(
          child, kSortConditionAttrs,
          {&cond.descending, &cond.sort_by, &cond.ref, &cond.custom_list, &cond.dxf_id, &cond.icon_set, &cond.icon_id});
    } else if (name == "extLst" && sort.ext_lst_xml.empty()) {
      sort.ext_lst_xml = io::raw_xml(child);
    } else {
      r.fail();
    }
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

/// Writes the attributes `specs` describes, in table order, from `targets`.
void write_attrs(std::string& out, const AttrSpec* specs, std::size_t n, const void* const* targets) {
  for (std::size_t i = 0; i < n; ++i) {
    const AttrSpec& spec = specs[i];
    const void* target = targets[i];
    switch (spec.kind) {
      case AttrKind::kBool: {
        const bool value = *static_cast<const bool*>(target);
        if (value != spec.bool_default) {
          io::append_xml_attr(out, spec.name, value ? "1" : "0");
        }
        break;
      }
      case AttrKind::kU32:
        if (const auto& value = *static_cast<const std::optional<std::uint32_t>*>(target)) {
          io::append_xml_attr_uint(out, spec.name, *value);
        }
        break;
      case AttrKind::kU32Required:
        io::append_xml_attr_uint(out, spec.name, *static_cast<const std::uint32_t*>(target));
        break;
      case AttrKind::kNumber:
        if (const auto& value = *static_cast<const std::optional<double>*>(target)) {
          append_number(out, spec.name, *value);
        }
        break;
      case AttrKind::kNumberRequired:
        append_number(out, spec.name, *static_cast<const double*>(target));
        break;
      case AttrKind::kText:
      case AttrKind::kTextAlways: {
        const std::string& value = *static_cast<const std::string*>(target);
        if (spec.kind == AttrKind::kTextAlways || !value.empty()) {
          io::append_xml_attr(out, spec.name, value);
        }
        break;
      }
      case AttrKind::kChoice:
      case AttrKind::kChoiceRequired: {
        const std::uint8_t index = *static_cast<const std::uint8_t*>(target);
        if (spec.kind == AttrKind::kChoiceRequired || index != 0U) {
          io::append_xml_attr(out, spec.name, spec.choices[index]);
        }
        break;
      }
      case AttrKind::kRect:
        append_ref(out, *static_cast<const MergeRange*>(target));
        break;
      case AttrKind::kExtras:
        io::append_raw_attrs(out, *static_cast<const RawAttributes*>(target));
        break;
    }
  }
}

template <std::size_t N>
void write_attrs(std::string& out, const AttrSpec (&specs)[N], const void* const (&targets)[N]) {
  write_attrs(out, specs, N, targets);
}

void append_value_filters(std::string& out, const ValueFilters& values, std::string_view prefix) {
  out.push_back('<');
  out.append(prefix);
  out.append("filters");
  write_attrs(out, kFiltersAttrs, {&values.blank, &values.calendar_type});
  if (values.values.empty() && values.date_groups.empty()) {
    out.append("/>");
    return;
  }
  out.push_back('>');
  for (const std::string& value : values.values) {
    out.push_back('<');
    out.append(prefix);
    out.append("filter");
    write_attrs(out, kFilterAttrs, {&value});
    out.append("/>");
  }
  for (const DateGroupItem& item : values.date_groups) {
    out.push_back('<');
    out.append(prefix);
    out.append("dateGroupItem");
    const std::array<std::uint32_t, 6> fields = {item.year, item.month, item.day, item.hour, item.minute, item.second};
    for (std::size_t i = 0; i <= static_cast<std::size_t>(item.grouping); ++i) {
      io::append_xml_attr_uint(out, kDateGroupAttrs[i].name, fields[i]);
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
  write_attrs(out, kColumnAttrs, {&col.col_id, &col.hidden_button, &col.show_button, &col.extra_attrs});
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
      write_attrs(body, kCustomFiltersAttrs, {&col.custom.and_join});
      body.push_back('>');
      for (const CustomFilter& filter : col.custom.filters) {
        body.append("<customFilter");
        write_attrs(body, kCustomFilterAttrs, {&filter.op, &filter.val});
        body.append("/>");
      }
      body.append("</customFilters>");
      break;
    case FilterKind::kTop10:
      body.append("<top10");
      write_attrs(body, kTop10Attrs, {&col.top10.top, &col.top10.percent, &col.top10.val, &col.top10.filter_val});
      body.append("/>");
      break;
    case FilterKind::kDynamic:
      body.append("<dynamicFilter");
      write_attrs(
          body, kDynamicAttrs,
          {&col.dynamic.type, &col.dynamic.val, &col.dynamic.val_iso, &col.dynamic.max_val, &col.dynamic.max_val_iso});
      body.append("/>");
      break;
    case FilterKind::kColor:
      body.append("<colorFilter");
      write_attrs(body, kColorAttrs, {&col.color.dxf_id, &col.color.cell_color});
      body.append("/>");
      break;
    case FilterKind::kIcon:
      body.append("<iconFilter");
      write_attrs(body, kIconAttrs, {&col.icon.icon_set, &col.icon.icon_id});
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
  write_attrs(out, kSortStateAttrs,
              {&sort.column_sort, &sort.case_sensitive, &sort.sort_method, &sort.extra_attrs, &sort.ref});
  if (sort.conditions.empty() && sort.ext_lst_xml.empty()) {
    out.append("/>");
    return;
  }
  out.push_back('>');
  for (const SortCondition& cond : sort.conditions) {
    out.append("<sortCondition");
    const void* const targets[] = {&cond.descending, &cond.sort_by,  &cond.ref,    &cond.custom_list,
                                   &cond.dxf_id,     &cond.icon_set, &cond.icon_id};
    write_attrs(out, kSortConditionAttrs, kSortConditionIconSet, targets);
    if (cond.sort_by == SortBy::kIcon || cond.icon_set != cf::IconSetName::Three_Arrows) {
      io::append_xml_attr(out, "iconSet", cf::kIconSetNames[static_cast<std::size_t>(cond.icon_set)]);
    }
    write_attrs(out, kSortConditionAttrs + kSortConditionIconSet + 1U, 1U, targets + kSortConditionIconSet + 1U);
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
  r.attrs(root, kRootAttrs, {&filter.range, &filter.extra_attrs});
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
  write_attrs(out, kRootAttrs, {&filter.range, &filter.extra_attrs});
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
