//
// C ABI - custom table / PivotTable styles (`<tableStyles>` in styles.xml).

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "styles.h"
#include "utils/error.h"
#include "utils/strings.h"
#include "workbook.h"

using formulon::c_api::parts::check_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;

namespace {

/// `ST_TableStyleType`, in schema order.
constexpr std::string_view kTableStyleTypes[] = {
    "wholeTable",
    "headerRow",
    "totalRow",
    "firstColumn",
    "lastColumn",
    "firstRowStripe",
    "secondRowStripe",
    "firstColumnStripe",
    "secondColumnStripe",
    "firstHeaderCell",
    "lastHeaderCell",
    "firstTotalCell",
    "lastTotalCell",
    "firstSubtotalColumn",
    "secondSubtotalColumn",
    "thirdSubtotalColumn",
    "firstSubtotalRow",
    "secondSubtotalRow",
    "thirdSubtotalRow",
    "blankRow",
    "firstColumnSubheading",
    "secondColumnSubheading",
    "thirdColumnSubheading",
    "firstRowSubheading",
    "secondRowSubheading",
    "thirdRowSubheading",
    "pageFieldLabels",
    "pageFieldValues",
};

bool is_table_style_type(std::string_view type) {
  for (const std::string_view known : kTableStyleTypes) {
    if (known == type) {
      return true;
    }
  }
  return false;
}

/// True for a built-in style name such as `TableStyleMedium2` or
/// `PivotStyleDark28`, which a custom style may not take.
bool is_builtin_style_name(std::string_view name) {
  struct Family {
    std::string_view prefix;
    unsigned count;
  };
  constexpr Family kFamilies[] = {
      {"TableStyleLight", 21U}, {"TableStyleMedium", 28U}, {"TableStyleDark", 11U},
      {"PivotStyleLight", 28U}, {"PivotStyleMedium", 28U}, {"PivotStyleDark", 28U},
  };
  for (const Family& family : kFamilies) {
    if (name.size() <= family.prefix.size() ||
        !formulon::strings::case_insensitive_eq(name.substr(0, family.prefix.size()), family.prefix)) {
      continue;
    }
    const std::string_view digits = name.substr(family.prefix.size());
    if (digits.size() > 2U || digits.front() == '0') {
      continue;
    }
    unsigned value = 0;
    bool numeric = true;
    for (const char ch : digits) {
      if (ch < '0' || ch > '9') {
        numeric = false;
        break;
      }
      value = value * 10U + static_cast<unsigned>(ch - '0');
    }
    if (numeric && value >= 1U && value <= family.count) {
      return true;
    }
  }
  return false;
}

std::size_t find_style(const formulon::TableStylesSection& section, std::string_view name) {
  for (std::size_t i = 0; i < section.styles.size(); ++i) {
    if (formulon::strings::case_insensitive_eq(section.styles[i].name, name)) {
      return i;
    }
  }
  return section.styles.size();
}

/// Built-in defaults the desktop application writes into a new workbook.
constexpr std::string_view kDefaultTableStyle = "TableStyleMedium2";
constexpr std::string_view kDefaultPivotStyle = "PivotStyleLight16";

}  // namespace

extern "C" fm_status_t fm_styles_get_table_style_count(fm_workbook_t* wb, uint32_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_table_style_count: NULL argument");
  }
  *out_count = static_cast<uint32_t>(wb->workbook().styles().table_styles.styles.size());
  return 0;
}

extern "C" fm_status_t fm_styles_get_table_style(fm_workbook_t* wb, uint32_t index, fm_table_style_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_table_style: NULL argument");
  }
  const auto& styles = wb->workbook().styles().table_styles.styles;
  if (auto rc = check_index(index, styles.size(), "fm_styles_get_table_style", "index"); rc != 0) {
    return rc;
  }
  const formulon::CustomTableStyle& style = styles[index];
  std::vector<fm_table_style_element_t> elements;
  elements.reserve(style.elements.size());
  for (const formulon::TableStyleElement& element : style.elements) {
    fm_table_style_element_t c{};
    c.type = element.type.c_str();
    c.size = element.size;
    c.has_dxf_id = element.dxf_id.has_value() ? 1 : 0;
    c.dxf_id = element.dxf_id.value_or(0U);
    elements.push_back(c);
  }
  wb->read_scratch.clear();
  wb->table_style_element_scratch.clear();
  *out = fm_table_style_t{};
  out->name = style.name.c_str();
  out->pivot = style.pivot ? 1 : 0;
  out->table = style.table ? 1 : 0;
  out->element_count = static_cast<uint32_t>(elements.size());
  out->elements = wb->table_style_element_scratch.adopt(std::move(elements));
  return 0;
}

extern "C" fm_status_t fm_styles_set_table_style(fm_workbook_t* wb, const fm_table_style_t* style) {
  clear_last_error();
  if (wb == nullptr || style == nullptr || style->name == nullptr ||
      (style->element_count != 0U && style->elements == nullptr)) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_set_table_style: NULL argument");
  }
  const std::string_view name = style->name;
  if (name.empty() || is_builtin_style_name(name)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_styles_set_table_style: name is empty or a built-in style name",
                             "name=" + std::string(name));
  }
  formulon::StylesTable& styles = wb->workbook().mutable_styles();
  formulon::CustomTableStyle next;
  next.name = std::string(name);
  next.pivot = style->pivot != 0;
  next.table = style->table != 0;
  next.elements.reserve(style->element_count);
  for (uint32_t i = 0; i < style->element_count; ++i) {
    const fm_table_style_element_t& in = style->elements[i];
    const std::string_view type = in.type != nullptr ? in.type : "";
    const std::string context = "element=" + std::to_string(i) + " type=" + std::string(type);
    if (!is_table_style_type(type)) {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               "fm_styles_set_table_style: unknown element type", context);
    }
    for (const formulon::TableStyleElement& earlier : next.elements) {
      if (earlier.type == type) {
        return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                                 "fm_styles_set_table_style: element type repeated", context);
      }
    }
    if (in.size == 0U) {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               "fm_styles_set_table_style: element size must be at least 1", context);
    }
    if (in.has_dxf_id != 0 && in.dxf_id >= styles.dxfs.size()) {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               "fm_styles_set_table_style: dxf_id out of range",
                               context + " dxf_id=" + std::to_string(in.dxf_id));
    }
    formulon::TableStyleElement element;
    element.type = std::string(type);
    element.size = in.size;
    if (in.has_dxf_id != 0) {
      element.dxf_id = in.dxf_id;
    }
    next.elements.push_back(std::move(element));
  }

  formulon::TableStylesSection& section = styles.table_styles;
  if (!section.present) {
    section.present = true;
    section.default_table_style = std::string(kDefaultTableStyle);
    section.default_pivot_style = std::string(kDefaultPivotStyle);
  }
  const std::size_t existing = find_style(section, name);
  if (existing < section.styles.size()) {
    // Attributes the record does not carry (`xr9:uid`) stay with the style.
    next.extra_attrs = std::move(section.styles[existing].extra_attrs);
    section.styles[existing] = std::move(next);
  } else {
    section.styles.push_back(std::move(next));
  }
  return 0;
}

extern "C" fm_status_t fm_styles_remove_table_style(fm_workbook_t* wb, const char* name) {
  clear_last_error();
  if (wb == nullptr || name == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_remove_table_style: NULL argument");
  }
  formulon::TableStylesSection& section = wb->workbook().mutable_styles().table_styles;
  const std::size_t index = find_style(section, name);
  if (index >= section.styles.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_styles_remove_table_style: no custom table style has that name",
                             "name=" + std::string(name));
  }
  section.styles.erase(section.styles.begin() + static_cast<std::ptrdiff_t>(index));
  if (formulon::strings::case_insensitive_eq(section.default_table_style, name)) {
    section.default_table_style = std::string(kDefaultTableStyle);
  }
  if (formulon::strings::case_insensitive_eq(section.default_pivot_style, name)) {
    section.default_pivot_style = std::string(kDefaultPivotStyle);
  }
  return 0;
}

extern "C" fm_status_t fm_styles_get_table_style_defaults(fm_workbook_t* wb, const char** out_table_style,
                                                          const char** out_pivot_style) {
  clear_last_error();
  if (wb == nullptr || out_table_style == nullptr || out_pivot_style == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_table_style_defaults: NULL argument");
  }
  const formulon::TableStylesSection& section = wb->workbook().styles().table_styles;
  *out_table_style = section.default_table_style.c_str();
  *out_pivot_style = section.default_pivot_style.c_str();
  return 0;
}

extern "C" fm_status_t fm_styles_set_table_style_defaults(fm_workbook_t* wb, const char* table_style,
                                                          const char* pivot_style) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_set_table_style_defaults: wb is NULL");
  }
  formulon::TableStylesSection& section = wb->workbook().mutable_styles().table_styles;
  if (table_style != nullptr) {
    section.default_table_style = table_style;
  }
  if (pivot_style != nullptr) {
    section.default_pivot_style = pivot_style;
  }
  if (table_style != nullptr || pivot_style != nullptr) {
    section.present = true;
  }
  return 0;
}
