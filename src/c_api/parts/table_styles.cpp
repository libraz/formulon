//
// C ABI - custom table / PivotTable styles (`<tableStyles>` in styles.xml).

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <tuple>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/enum_map.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "c_api/parts/table_style_defaults.h"
#include "sheet_name.h"
#include "styles.h"
#include "utils/error.h"
#include "utils/strings.h"
#include "workbook.h"

using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::fresh_arena;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::kTableStyleElementTypeTokens;
using formulon::c_api::parts::not_found;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::ordinal_of;
using formulon::c_api::parts::patch_op_ok;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::token_of;

namespace {

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

// The custom style `name` resolves to (`sheet_names::resolve_named`), or `section.styles.size()` when none.
std::size_t find_style(const formulon::TableStylesSection& section, std::string_view name) {
  const std::size_t found = formulon::sheet_names::resolve_named(section.styles, name);
  return found == formulon::sheet_names::kNotFound ? section.styles.size() : found;
}

/// Built-in defaults the desktop application writes into a new workbook.
constexpr std::string_view kDefaultTableStyle = "TableStyleMedium2";
constexpr std::string_view kDefaultPivotStyle = "PivotStyleLight16";

}  // namespace

namespace formulon::c_api::parts {

std::string_view default_table_style_name(const formulon::Workbook& book) {
  const std::string& name = book.styles().table_styles.default_table_style;
  return name.empty() ? kDefaultTableStyle : std::string_view(name);
}

std::string_view default_pivot_style_name(const formulon::Workbook& book) {
  const std::string& name = book.styles().table_styles.default_pivot_style;
  return name.empty() ? kDefaultPivotStyle : std::string_view(name);
}

}  // namespace formulon::c_api::parts

extern "C" fm_status_t fm_workbook_get_table_style_count(const fm_workbook_t* wb, size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return null_argument();
  }
  *out_count = wb->workbook().styles().table_styles.styles.size();
  return 0;
}

extern "C" fm_status_t fm_workbook_get_table_style_at(const fm_workbook_t* wb, size_t index, fm_table_style_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  const auto& styles = wb->workbook().styles().table_styles.styles;
  if (auto rc = check_index(index, styles.size(), "index"); rc != 0) {
    return rc;
  }
  const formulon::CustomTableStyle& style = styles[index];
  std::vector<fm_table_style_element_t> elements;
  elements.reserve(style.elements.size());
  for (const formulon::TableStyleElement& element : style.elements) {
    fm_table_style_element_t c{};
    c.type = ordinal_of(kTableStyleElementTypeTokens, element.type, -1);
    c.size = element.size;
    c.has_dxf_index = element.dxf_id.has_value() ? 1 : 0;
    c.dxf_index = element.dxf_id.value_or(0U);
    elements.push_back(c);
  }
  fresh_read_scratch(wb);
  auto& element_arena = fresh_arena(wb, &fm_workbook_t::table_style_element_scratch);
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_table_style_t{};
  out->struct_size = struct_size;
  out->name = style.name.c_str();
  out->pivot = style.pivot ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
  out->table = style.table ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
  out->element_count = static_cast<uint32_t>(elements.size());
  out->elements = element_arena.adopt(std::move(elements));
  return 0;
}

extern "C" fm_status_t fm_workbook_get_table_style(const fm_workbook_t* wb, const char* name, int32_t* out_present,
                                                   fm_table_style_t* out) {
  clear_last_error();
  if (wb == nullptr || name == nullptr || out_present == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({name}); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  *out_present = 0;
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_table_style_t{};
  out->struct_size = struct_size;
  const formulon::TableStylesSection& section = wb->workbook().styles().table_styles;
  const std::size_t index = find_style(section, name);
  if (index >= section.styles.size()) {
    return 0;
  }
  *out_present = 1;
  return fm_workbook_get_table_style_at(wb, index, out);
}

extern "C" fm_status_t fm_workbook_set_table_style(fm_workbook_t* wb, const fm_table_style_t* style) {
  clear_last_error();
  if (wb == nullptr || style == nullptr || style->name == nullptr ||
      (style->element_count != 0U && style->elements == nullptr)) {
    return null_argument();
  }
  if (auto rc = check_record_size(*style); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({style->name}); rc != 0) {
    return rc;
  }
  const std::string_view name = style->name;
  if (name.empty() || is_builtin_style_name(name)) {
    return invalid_argument("name is empty or a built-in style name", "name=" + std::string(name));
  }
  if (auto rc = check_enum_domain(style->pivot, FM_TRISTATE_DEFAULT, FM_TRISTATE_TRUE, "pivot"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(style->table, FM_TRISTATE_DEFAULT, FM_TRISTATE_TRUE, "table"); rc != 0) {
    return rc;
  }
  formulon::StylesTable& styles = wb->workbook().mutable_styles();
  formulon::CustomTableStyle next;
  next.name = std::string(name);
  next.pivot = style->pivot != FM_TRISTATE_FALSE;
  next.table = style->table != FM_TRISTATE_FALSE;
  next.elements.reserve(style->element_count);
  for (uint32_t i = 0; i < style->element_count; ++i) {
    const fm_table_style_element_t& in = style->elements[i];
    if (auto rc = check_enum_domain(in.type, FM_TABLE_STYLE_ELEMENT_TYPE_WHOLE_TABLE,
                                    FM_TABLE_STYLE_ELEMENT_TYPE_PAGE_FIELD_VALUES, "type");
        rc != 0) {
      return rc;
    }
    const std::string_view type = token_of(kTableStyleElementTypeTokens, in.type);
    const std::string context = "element=" + std::to_string(i) + " type=" + std::string(type);
    for (const formulon::TableStyleElement& earlier : next.elements) {
      if (earlier.type == type) {
        return invalid_argument("element type repeated", context);
      }
    }
    if (in.size == 0U) {
      return invalid_argument("element size must be at least 1", context);
    }
    if (in.has_dxf_index != 0 && in.dxf_index >= styles.dxfs.size()) {
      return invalid_argument("dxf_index out of range", context + " dxf_index=" + std::to_string(in.dxf_index));
    }
    formulon::TableStyleElement element;
    element.type = std::string(type);
    element.size = in.size;
    if (in.has_dxf_index != 0) {
      element.dxf_id = in.dxf_index;
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
  } else if (std::any_of(section.styles.begin(), section.styles.end(), [&](const auto& style) {
               return formulon::sheet_names::write_collides(style.name, next.name);
             })) {
    return set_binding_error(formulon::FormulonErrorCode::kAlreadyExists, "name already in use", "name=" + next.name);
  } else {
    section.styles.push_back(std::move(next));
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_remove_table_style(fm_workbook_t* wb, const char* name) {
  clear_last_error();
  if (wb == nullptr || name == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({name}); rc != 0) {
    return rc;
  }
  formulon::TableStylesSection& section = wb->workbook().mutable_styles().table_styles;
  const std::size_t index = find_style(section, name);
  if (index >= section.styles.size()) {
    return not_found("no custom table style has that name", "name=" + std::string(name));
  }
  // A default naming the removed style falls back to the built-in one.
  const bool table_default = find_style(section, section.default_table_style) == index;
  const bool pivot_default = find_style(section, section.default_pivot_style) == index;
  section.styles.erase(section.styles.begin() + static_cast<std::ptrdiff_t>(index));
  if (table_default) {
    section.default_table_style = std::string(kDefaultTableStyle);
  }
  if (pivot_default) {
    section.default_pivot_style = std::string(kDefaultPivotStyle);
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_get_table_style_defaults(const fm_workbook_t* wb, fm_table_style_defaults_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  const formulon::TableStylesSection& section = wb->workbook().styles().table_styles;
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_table_style_defaults_t{};
  out->struct_size = struct_size;
  out->table_style_name = section.default_table_style.empty() ? nullptr : section.default_table_style.c_str();
  out->pivot_style_name = section.default_pivot_style.empty() ? nullptr : section.default_pivot_style.c_str();
  return 0;
}

extern "C" fm_status_t fm_workbook_update_table_style_defaults(fm_workbook_t* wb,
                                                               const fm_table_style_defaults_patch_t* patch) {
  clear_last_error();
  if (wb == nullptr || patch == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  // The whole patch is checked before anything changes.
  for (const auto& [op, name, field] : {std::tuple<fm_patch_op_t, const char*, const char*>{
                                            patch->table_style_name_op, patch->table_style_name, "table_style_name"},
                                        {patch->pivot_style_name_op, patch->pivot_style_name, "pivot_style_name"}}) {
    if (!patch_op_ok(op)) {
      return invalid_argument("patch op out of range", std::string("field=") + field);
    }
    if (op == FM_PATCH_OP_SET) {
      if (name == nullptr || name[0] == '\0') {
        return invalid_argument("SET needs a non-empty style name", std::string("field=") + field);
      }
      if (auto rc = check_utf8(name, "field", field); rc != 0) {
        return rc;
      }
    }
  }
  if (patch->table_style_name_op == FM_PATCH_OP_KEEP && patch->pivot_style_name_op == FM_PATCH_OP_KEEP) {
    return 0;
  }
  formulon::TableStylesSection& section = wb->workbook().mutable_styles().table_styles;
  if (patch->table_style_name_op != FM_PATCH_OP_KEEP) {
    section.default_table_style = patch->table_style_name_op == FM_PATCH_OP_SET ? patch->table_style_name : "";
  }
  if (patch->pivot_style_name_op != FM_PATCH_OP_KEEP) {
    section.default_pivot_style = patch->pivot_style_name_op == FM_PATCH_OP_SET ? patch->pivot_style_name : "";
  }
  section.present = true;
  return 0;
}
