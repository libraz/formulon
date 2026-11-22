//
// C ABI - field validation for the records callers pass in: conditional
// formatting rules, data validations, fonts, fills, borders and pivot specs.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/enum_map.h"
#include "cf/cf_types.h"
#include "styles.h"
#include "utils/error.h"
#include "utils/resource_budget.h"
#include "utils/strings.h"

namespace formulon {
namespace c_api {
namespace parts {

namespace {

// Rejects a string that is not GUID-shaped (`{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`,
// hexadecimal digits). NULL passes: the entry point synthesizes an id. `""` is refused like any other non-GUID,
// since `ST_Guid` attributes holding anything else make Excel drop the whole block.
fm_status_t check_guid(const char* text, const char* field) {
  if (text == nullptr) {
    return 0;
  }
  // `{8-4-4-4-12}`: 32 hexadecimal digits, four hyphens, two braces.
  static constexpr std::string_view kShape = "{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}";
  const std::string_view value(text);
  bool shaped = value.size() == kShape.size();
  for (std::size_t i = 0; shaped && i < kShape.size(); ++i) {
    switch (kShape[i]) {
      case 'X':
        shaped = formulon::strings::hex_digit_value(value[i]) >= 0;
        break;
      default:
        shaped = value[i] == kShape[i];
        break;
    }
  }
  if (shaped) {
    return 0;
  }
  return invalid_argument((std::string(field) + " is not a GUID").c_str(),
                          std::string(field) + "=" + std::string(value));
}

}  // namespace

namespace {

// Checks the `type` and `gte` (`fm_tristate_t`) domains of every `fm_cfvo_t`
// in a threshold array and refuses an empty `value`; `field` names the array in the diagnostic.
// `max_type` is the highest `type` the array accepts.
// The array bound is the caller's, so an entry point that has not yet
// established `count` against its pointer must not reach here.
fm_status_t check_cfvos(const fm_cfvo_t* thresholds, std::uint32_t count, const char* field,
                        cf::CfvoType max_type = cf::CfvoType::AutoMax) {
  const std::string type_field = std::string(field) + ".type";
  const std::string gte_field = std::string(field) + ".gte";
  for (std::uint32_t i = 0; i < count; ++i) {
    if (thresholds[i].value != nullptr && thresholds[i].value[0] == '\0') {
      return invalid_argument((std::string(field) + ".value must not be empty").c_str(),
                              std::string("field=") + field + ".value");
    }
    if (auto rc = check_enum_domain(thresholds[i].type, 0, static_cast<std::int64_t>(max_type), type_field.c_str());
        rc != 0) {
      return rc;
    }
    if (auto rc = check_enum_domain(thresholds[i].gte, FM_TRISTATE_DEFAULT, FM_TRISTATE_TRUE, gte_field.c_str());
        rc != 0) {
      return rc;
    }
  }
  return 0;
}

}  // namespace

fm_status_t validate(const fm_cf_rule_t& rule) {
  // A range the A1 encoder cannot spell reaches the writer as an empty
  // reference, which Excel reads as a broken conditional-formatting
  // block. Rejecting it here names the offending rectangle at set time
  // instead of losing the format at save time.
  if (auto rc = check_rect_list(rule.ranges, rule.range_count); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(rule.type, 0, static_cast<std::int64_t>(cf::RuleType::UniqueValues), "type");
      rc != 0) {
    return rc;
  }
  if (rule.has_op != 0) {
    if (auto rc = check_enum_domain(rule.op, 0, FM_COMPARISON_OPERATOR_GREATER_THAN_OR_EQUAL, "op"); rc != 0) {
      return rc;
    }
  }
  if (rule.has_time_period != 0) {
    if (auto rc = check_enum_domain(rule.time_period, 0, FM_DATE_PERIOD_NEXT_WEEK, "time_period"); rc != 0) {
      return rc;
    }
  }
  if (rule.has_data_bar != 0) {
    if (auto rc = check_cfvos(&rule.data_bar_min, 1U, "data_bar_min"); rc != 0) {
      return rc;
    }
    if (auto rc = check_cfvos(&rule.data_bar_max, 1U, "data_bar_max"); rc != 0) {
      return rc;
    }
    if (rule.has_data_bar_axis_position != 0) {
      if (auto rc = check_enum_domain(rule.data_bar_axis_position, 0, FM_CF_DATA_BAR_AXIS_POSITION_NONE,
                                      "data_bar_axis_position");
          rc != 0) {
        return rc;
      }
    }
    if (auto rc =
            check_enum_domain(rule.data_bar_direction, 0, FM_CF_DATA_BAR_DIRECTION_RIGHT_TO_LEFT, "data_bar_direction");
        rc != 0) {
      return rc;
    }
  }
  // The colour-scale scan is bounded by the array shape the entry point
  // accepts, so a record that declares more thresholds than the schema
  // allows is rejected there rather than scanned here. AutoMin and AutoMax are data-bar types.
  if (rule.color_scale_thresholds != nullptr && rule.color_scale_count >= 2U && rule.color_scale_count <= 3U) {
    if (auto rc = check_cfvos(rule.color_scale_thresholds, rule.color_scale_count, "color_scale_thresholds[]",
                              cf::CfvoType::Formula);
        rc != 0) {
      return rc;
    }
  }
  if (rule.has_icon_set != 0) {
    if (auto rc =
            check_enum_domain(rule.icon_set, 0, static_cast<std::int64_t>(cf::IconSetName::Five_Quarters), "icon_set");
        rc != 0) {
      return rc;
    }
    if (rule.icon_set_thresholds != nullptr) {
      if (auto rc = check_cfvos(rule.icon_set_thresholds, rule.icon_set_threshold_count, "icon_set_thresholds[]");
          rc != 0) {
        return rc;
      }
    }
    if (rule.has_icon_set_floor != 0) {
      if (auto rc = check_cfvos(&rule.icon_set_floor, 1U, "icon_set_floor"); rc != 0) {
        return rc;
      }
    }
    if (rule.icon_set_icon_count != 0U) {
      // One icon per level, each naming an icon its source set has.
      const std::size_t levels = cf::icon_count(static_cast<cf::IconSetName>(rule.icon_set));
      if (rule.icon_set_icons == nullptr || rule.icon_set_icon_count != levels) {
        return invalid_argument(
            "icon_set_icons must hold one icon per level",
            "icon_set_icons count=" + std::to_string(rule.icon_set_icon_count) + " levels=" + std::to_string(levels));
      }
      for (std::uint32_t i = 0; i < rule.icon_set_icon_count; ++i) {
        const fm_cf_icon_t& icon = rule.icon_set_icons[i];
        const std::size_t available = icon.icon_set < 0 || icon.icon_set > FM_ICON_SET_NAME_NO_ICONS
                                          ? 0U
                                          : cf::cf_icon_set_icon_count(token_of(kIconSetNameTokens, icon.icon_set));
        if (available == 0U || icon.icon_index < 0 || static_cast<std::size_t>(icon.icon_index) >= available) {
          return invalid_argument(("icon_set_icons[" + std::to_string(i) + "] names no icon").c_str(),
                                  "icon_set_icons[" + std::to_string(i) + "]");
        }
      }
    }
  }
  if (auto rc = check_utf8({rule.text}); rc != 0) {
    return rc;
  }
  for (const auto& [value, field] : {std::pair<const char*, const char*>{rule.formula1, "formula1"},
                                     std::pair<const char*, const char*>{rule.formula2, "formula2"}}) {
    if (value != nullptr && value[0] == '\0') {
      return invalid_argument((std::string(field) + " must not be empty").c_str(), std::string("field=") + field);
    }
  }
  // The id reaches the file as `<x14:cfRule id="...">` and `<x14:id>`,
  // both of schema type `ST_Guid`. Excel repairs a workbook whose
  // extension block holds anything else, and repairing means dropping
  // every conditional format in it.
  return check_guid(rule.id, "id");
}

fm_status_t validate(const fm_data_validation_t& v) {
  if (auto rc = check_rect_list(v.ranges, v.range_count); rc != 0) {
    return rc;
  }
  // The three domains the writer's attribute tables spell out
  // (`sheet_xml_builder.cpp`): a value past them is emitted as an absent
  // attribute, so a mistyped `type` saves as "no validation at all".
  if (auto rc = check_enum_domain(v.type, 0, 7, "type"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(v.op, 0, FM_COMPARISON_OPERATOR_GREATER_THAN_OR_EQUAL, "op"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(v.error_style, 0, 2, "error_style"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(v.allow_blank, FM_TRISTATE_DEFAULT, FM_TRISTATE_TRUE, "allow_blank"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(v.show_input_message, FM_TRISTATE_DEFAULT, FM_TRISTATE_TRUE, "show_input_message");
      rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(v.show_error_message, FM_TRISTATE_DEFAULT, FM_TRISTATE_TRUE, "show_error_message");
      rc != 0) {
    return rc;
  }
  if (auto rc =
          check_enum_domain(v.show_in_cell_dropdown, FM_TRISTATE_DEFAULT, FM_TRISTATE_TRUE, "show_in_cell_dropdown");
      rc != 0) {
    return rc;
  }
  return check_utf8({v.error_title, v.error_message, v.prompt_title, v.prompt_message});
}

fm_status_t validate(const fm_font_t& font) {
  if (fm_status_t rc = check_finite_non_negative(font.size_pt, "font.size_pt"); rc != 0) {
    return rc;
  }
  if (fm_status_t rc = check_font_enums(font); rc != 0) {
    return rc;
  }
  return check_utf8({font.name});
}

fm_status_t check_font_enums(const fm_font_t& font) {
  if (fm_status_t rc = check_enum_domain(font.underline, 0, FM_UNDERLINE_STYLE_DOUBLE_ACCOUNTING, "font.underline");
      rc != 0) {
    return rc;
  }
  if (fm_status_t rc = check_enum_domain(font.vert_align, 0, FM_FONT_VERTICAL_ALIGN_SUBSCRIPT, "font.vert_align");
      rc != 0) {
    return rc;
  }
  if (fm_status_t rc = check_enum_domain(font.family, 0, FM_FONT_FAMILY_DECORATIVE, "font.family"); rc != 0) {
    return rc;
  }
  return check_enum_domain(font.scheme, 0, FM_FONT_SCHEME_MINOR, "font.scheme");
}

fm_status_t validate(const fm_fill_t& fill) {
  if (fm_status_t rc = check_enum_domain(fill.pattern, 0, FM_FILL_PATTERN_GRAY0625, "fill.pattern"); rc != 0) {
    return rc;
  }
  if (fm_status_t rc = check_enum_domain(fill.gradient_type, 0, FM_GRADIENT_TYPE_PATH, "fill.gradient_type"); rc != 0) {
    return rc;
  }
  if (fill.has_gradient != 0 && fill.gradient_stops != nullptr &&
      fill.gradient_stop_count > formulon::kMaxGradientStops) {
    return set_binding_error(
        formulon::FormulonErrorCode::kSecurityResourceLimit, "gradient_stop_count is above the limit",
        formulon::limit_context("gradient_stops", fill.gradient_stop_count, formulon::kMaxGradientStops));
  }
  return 0;
}

fm_status_t validate(const fm_border_t& border) {
  const std::pair<const fm_border_side_t*, const char*> sides[] = {{&border.left, "border.left.style"},
                                                                   {&border.right, "border.right.style"},
                                                                   {&border.top, "border.top.style"},
                                                                   {&border.bottom, "border.bottom.style"},
                                                                   {&border.diagonal, "border.diagonal.style"}};
  for (const auto& [side, field] : sides) {
    if (fm_status_t rc = check_enum_domain(side->style, 0, FM_BORDER_STYLE_SLANT_DASH_DOT, field); rc != 0) {
      return rc;
    }
  }
  return 0;
}

fm_status_t validate(const fm_pivot_data_field_t& spec) {
  if (auto rc = check_utf8({spec.name}); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(spec.aggregation, 0, FM_PIVOT_AGGREGATION_VARP, "aggregation"); rc != 0) {
    return rc;
  }
  return check_enum_domain(spec.show_as, 0, FM_PIVOT_SHOW_VALUES_AS_RANK_DESCENDING, "show_as");
}

fm_status_t validate(const fm_pivot_filter_t& spec) {
  if (auto rc = check_utf8({spec.value_text, spec.value_high_text}); rc != 0) {
    return rc;
  }
  return check_enum_domain(spec.kind, 0, FM_PIVOT_FILTER_KIND_M12, "kind");
}

}  // namespace parts
}  // namespace c_api
}  // namespace formulon
