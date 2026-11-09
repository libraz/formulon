//
// JsWorkbook conditional-formatting surface: read / add / remove / clear
// rules plus `evaluateCfRange` which projects matched rules onto a
// requested cell window. Visual rule kinds (`ColorScale`, `DataBar`,
// `IconSet`) are exposed as structured payload objects on both read and
// add paths.

#include <emscripten/val.h>

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "c_api/borrowed_arena.h"
#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {
namespace {

fm_cfvo_t js_pull_cfvo(emscripten::val v, formulon::c_api::BorrowedStringArena* strings,
                       JsNarrowNumericReader& reader) {
  fm_cfvo_t out{};
  out.type = reader.u8(v, "type", 0U, "cfvo.type");
  out.gte = reader.boolean(v, "gte", true, "cfvo.gte") ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
  const emscripten::val value = reader.value(v, "value", "cfvo.value");
  if (js_value_present(value)) {
    out.value = strings->emplace(reader.string_value(value, "cfvo.value"));
  }
  return out;
}

emscripten::val cfvo_to_js(const fm_cfvo_t& cfvo) {
  emscripten::val out = emscripten::val::object();
  out.set("type", static_cast<uint32_t>(cfvo.type));
  out.set("gte", cfvo.gte != FM_TRISTATE_FALSE);
  out.set("value", cfvo.value != nullptr ? emscripten::val(std::string(cfvo.value)) : emscripten::val::null());
  return out;
}

}  // namespace

emscripten::val JsWorkbook::evaluateCfRange(uint32_t sheet, emscripten::val range, emscripten::val options) const {
  emscripten::val cells = emscripten::val::array();
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  JsNarrowNumericReader reader("Sheet.evaluateCfRange");
  const fm_cell_rect_t rect = js_pull_range(range, &reader);
  // `options` is `{ timePeriodRules?, todaySerial? }`; an omitted key is the C default: `timePeriod` rules on,
  // measured from the pinned clock, else today.
  fm_cf_range_evaluation_spec_t spec = FM_STRUCT_INIT(fm_cf_range_evaluation_spec_t);
  if (js_value_present(options) && reader.require_object(options, "options")) {
    const emscripten::val rules = reader.value(options, "timePeriodRules", "options.timePeriodRules");
    if (js_value_present(rules)) {
      spec.time_period_rules = reader.boolean_value(rules, true) ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
    }
    const emscripten::val today = reader.value(options, "todaySerial", "options.todaySerial");
    if (js_value_present(today)) {
      spec.has_today_serial = 1;
      spec.today_serial = reader.number_value(today, 0.0, "options.todaySerial");
    }
  }
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  fm_cf_results_t* results = nullptr;
  fm_status_t rc = fm_sheet_evaluate_cf_range(handle_, sheet, &rect, &spec, &results);
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  std::size_t cell_count = 0;
  rc = fm_cf_results_get_cell_count(results, &cell_count);
  if (!succeeded(rc)) {
    fm_cf_results_destroy(results);
    return emscripten::val::undefined();
  }
  std::size_t emitted_cells = 0;
  for (std::size_t i = 0; i < cell_count; ++i) {
    uint32_t row = 0;
    uint32_t col = 0;
    std::size_t match_count = 0;
    if (fm_cf_results_get_cell_at(results, i, &row, &col, &match_count) != 0) {
      // Skip entries the C ABI declines to materialise; this is purely
      // defensive -- the index is always valid by construction.
      continue;
    }
    emscripten::val cell = emscripten::val::object();
    emscripten::val matches = emscripten::val::array();
    std::size_t emitted_matches = 0;
    for (std::size_t j = 0; j < match_count; ++j) {
      fm_cf_match_info_t m = FM_STRUCT_INIT(fm_cf_match_info_t);
      if (fm_cf_results_get_match_at(results, i, j, &m) != 0) {
        continue;
      }
      matches.set(emitted_matches, translate_cf_match(m));
      ++emitted_matches;
    }
    cell.set("row", row);
    cell.set("col", col);
    cell.set("matches", matches);
    cells.set(emitted_cells, cell);
    ++emitted_cells;
  }
  fm_cf_results_destroy(results);
  return cells;
}

namespace {

// Reads CF rule `index` of `sheet` into its `ConditionalFormatEntry` shape.
emscripten::val cf_rule_to_val(const fm_cf_rule_t& rule) {
  const emscripten::val null = emscripten::val::null();
  emscripten::val item = emscripten::val::object();
  js_set_cstr(item, "id", rule.id);
  item.set("type", static_cast<uint32_t>(rule.type));
  item.set("priority", rule.priority);
  item.set("stopIfTrue", rule.stop_if_true != 0);
  item.set("ranges", cell_rects_to_val(rule.ranges, rule.range_count));
  js_set_present_or_null(item, "dxfIndex", rule.has_dxf_index, rule.dxf_index);
  js_set_cstr_or_null(item, "formula1", rule.formula1);
  js_set_cstr_or_null(item, "formula2", rule.formula2);
  js_set_present_or_null(item, "operator", rule.has_op, static_cast<uint32_t>(rule.op));
  js_set_present_or_null(item, "rank", rule.has_rank, rule.rank);
  item.set("percent", rule.has_rank != 0 ? emscripten::val(rule.percent != 0) : null);
  item.set("bottom", rule.has_rank != 0 ? emscripten::val(rule.bottom != 0) : null);
  // The average flags belong to the AboveAverage rule type only.
  const bool above_average_rule = rule.type == 6 /* AboveAverage */;
  item.set("aboveAverage", above_average_rule ? emscripten::val(rule.above_average != 0) : null);
  item.set("equalAverage", above_average_rule ? emscripten::val(rule.equal_average != 0) : null);
  js_set_present_or_null(item, "stdDev", above_average_rule ? rule.has_std_dev : 0, rule.std_dev);
  js_set_cstr_or_null(item, "text", rule.text);
  js_set_present_or_null(item, "timePeriod", rule.has_time_period, static_cast<uint32_t>(rule.time_period));
  if (rule.color_scale_count > 0 && rule.color_scale_thresholds != nullptr && rule.color_scale_colors != nullptr) {
    emscripten::val color_scale = emscripten::val::object();
    emscripten::val thresholds = emscripten::val::array();
    emscripten::val colors = emscripten::val::array();
    for (uint32_t j = 0; j < rule.color_scale_count; ++j) {
      thresholds.set(j, cfvo_to_js(rule.color_scale_thresholds[j]));
      colors.set(j, js_color_spec(rule.color_scale_colors[j]));
    }
    color_scale.set("thresholds", thresholds);
    color_scale.set("colors", colors);
    item.set("colorScale", color_scale);
  } else {
    item.set("colorScale", null);
  }
  if (rule.has_data_bar != 0) {
    emscripten::val data_bar = emscripten::val::object();
    data_bar.set("min", cfvo_to_js(rule.data_bar_min));
    data_bar.set("max", cfvo_to_js(rule.data_bar_max));
    data_bar.set("fill", js_color_spec(rule.data_bar_fill));
    data_bar.set("showValue", rule.data_bar_show_value != 0);
    data_bar.set("minLengthPct", static_cast<uint32_t>(rule.data_bar_min_length_pct));
    data_bar.set("maxLengthPct", static_cast<uint32_t>(rule.data_bar_max_length_pct));
    // The six `x14` keys read `null` when the rule does not override the model default.
    data_bar.set("gradient", rule.has_data_bar_gradient != 0 ? emscripten::val(rule.data_bar_gradient != 0) : null);
    js_set_present_or_null(data_bar, "axisPosition", rule.has_data_bar_axis_position,
                           static_cast<uint32_t>(rule.data_bar_axis_position));
    const auto color_or_null = [&null](const fm_color_t& color) {
      return color.kind != FM_COLOR_KIND_NONE ? js_color_spec(color) : null;
    };
    data_bar.set("negativeFill", color_or_null(rule.data_bar_negative_fill));
    data_bar.set("border", color_or_null(rule.data_bar_border));
    data_bar.set("negativeBorder", color_or_null(rule.data_bar_negative_border));
    data_bar.set("axisColor", color_or_null(rule.data_bar_axis_color));
    data_bar.set("direction", static_cast<uint32_t>(rule.data_bar_direction));
    item.set("dataBar", data_bar);
  } else {
    item.set("dataBar", null);
  }
  if (rule.has_icon_set != 0) {
    emscripten::val icon_set = emscripten::val::object();
    icon_set.set("name", static_cast<uint32_t>(rule.icon_set));
    emscripten::val thresholds = emscripten::val::array();
    for (uint32_t j = 0; j < rule.icon_set_threshold_count; ++j) {
      thresholds.set(j, cfvo_to_js(rule.icon_set_thresholds[j]));
    }
    icon_set.set("thresholds", thresholds);
    icon_set.set("reverse", rule.icon_set_reverse != 0);
    icon_set.set("showValue", rule.icon_set_show_value != 0);
    icon_set.set("percent", rule.icon_set_percent != 0);
    icon_set.set("floor", rule.has_icon_set_floor != 0 ? cfvo_to_js(rule.icon_set_floor) : null);
    if (rule.icon_set_icon_count != 0U) {
      emscripten::val icons = emscripten::val::array();
      for (uint32_t j = 0; j < rule.icon_set_icon_count; ++j) {
        emscripten::val icon = emscripten::val::object();
        icon.set("iconSet", rule.icon_set_icons[j].icon_set);
        icon.set("iconIndex", rule.icon_set_icons[j].icon_index);
        icons.set(j, icon);
      }
      icon_set.set("icons", icons);
    } else {
      icon_set.set("icons", null);
    }
    item.set("iconSet", icon_set);
  } else {
    item.set("iconSet", null);
  }
  return item;
}

}  // namespace

double JsWorkbook::getCfRuleCount(uint32_t sheet) const {
  return js_count(handle_, sheet, &fm_sheet_get_cf_rule_count);
}

emscripten::val JsWorkbook::getCfRule(uint32_t sheet, emscripten::val cfRule) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  if (cfRule.isString()) {
    const std::string id = cfRule.as<std::string>();
    int32_t present = 0;
    const fm_status_t rc = fm_sheet_get_cf_rule(handle_, sheet, id.c_str(), &present, &rule);
    return optional_or_null(rc, present, rc == 0 && present != 0 ? cf_rule_to_val(rule) : emscripten::val::null());
  }
  JsNarrowNumericReader reader("Sheet.getCfRule");
  if (!cfRule.isNumber()) {
    reader.fail_type("cfRule", "a number or string");
  }
  const uint32_t index = reader.u32_value(cfRule, 0U, "cfRule");
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  if (!succeeded(fm_sheet_get_cf_rule_at(handle_, sheet, index, &rule))) {
    return emscripten::val::undefined();
  }
  return cf_rule_to_val(rule);
}

emscripten::val JsWorkbook::addConditionalFormat(uint32_t sheet, emscripten::val v) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  // Pull every JS field into local storage; the C ABI receives
  // borrowed `const char*` views that must stay valid for the
  // duration of the call. Threshold value strings go into an arena
  // rather than a vector: the colorScale, dataBar and iconSet branches
  // below all feed the same store, and a vector would relocate the
  // earlier strings as the later branches add theirs.
  std::vector<fm_cfvo_t> color_scale_thresholds;
  std::vector<fm_color_t> color_scale_colors;
  std::vector<fm_cfvo_t> icon_set_thresholds;
  std::vector<fm_cf_icon_t> icon_set_icons;
  formulon::c_api::BorrowedStringArena cfvo_strings;
  JsNarrowNumericReader reader("Sheet.addCfRule");
  const emscripten::val ranges_value = reader.value(v, "ranges", "rule.ranges");
  std::vector<fm_cell_rect_t> merge_ranges;
  if (reader.require_array(ranges_value, "rule.ranges")) {
    merge_ranges = js_pull_range_array(ranges_value, "rule.ranges", &reader);
  }
  std::vector<fm_cell_rect_t> ranges_buf;
  ranges_buf.reserve(merge_ranges.size());
  for (const fm_cell_rect_t& m : merge_ranges) {
    ranges_buf.push_back(fm_cell_rect_t{m.first_row, m.first_col, m.last_row, m.last_col});
  }
  const std::string id = reader.string(v, "id", "rule.id");
  const std::string formula1 = reader.string(v, "formula1", "rule.formula1");
  const std::string formula2 = reader.string(v, "formula2", "rule.formula2");
  const std::string text = reader.string(v, "text", "rule.text");

  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.id = id.empty() ? nullptr : id.c_str();
  rule.type = reader.u8(v, "type", 0U, "rule.type");
  rule.priority = reader.i32(v, "priority", 0, "rule.priority");
  rule.stop_if_true = reader.boolean(v, "stopIfTrue", false, "rule.stopIfTrue") ? 1 : 0;
  const emscripten::val dxf_index = reader.value(v, "dxfIndex", "rule.dxfIndex");
  if (js_value_present(dxf_index)) {
    rule.has_dxf_index = 1;
    rule.dxf_index = reader.u32_value(dxf_index, 0U, "rule.dxfIndex");
  }
  rule.ranges = ranges_buf.empty() ? nullptr : ranges_buf.data();
  rule.range_count = ranges_buf.size();
  rule.formula1 = formula1.empty() ? nullptr : formula1.c_str();
  rule.formula2 = formula2.empty() ? nullptr : formula2.c_str();
  const emscripten::val op = reader.value(v, "operator", "rule.operator");
  if (js_value_present(op)) {
    rule.has_op = 1;
    rule.op = reader.u8_value(op, 0U, "rule.operator");
  }
  const emscripten::val rank = reader.value(v, "rank", "rule.rank");
  if (js_value_present(rank)) {
    rule.has_rank = 1;
    rule.rank = reader.i32_value(rank, 0, "rule.rank");
  }
  rule.percent = reader.boolean(v, "percent", false, "rule.percent") ? 1 : 0;
  rule.bottom = reader.boolean(v, "bottom", false, "rule.bottom") ? 1 : 0;
  rule.above_average = reader.boolean(v, "aboveAverage", true, "rule.aboveAverage") ? 1 : 0;
  rule.equal_average = reader.boolean(v, "equalAverage", false, "rule.equalAverage") ? 1 : 0;
  const emscripten::val std_dev = reader.value(v, "stdDev", "rule.stdDev");
  if (js_value_present(std_dev)) {
    rule.has_std_dev = 1;
    rule.std_dev = reader.number_value(std_dev, 0.0, "rule.stdDev");
  }
  rule.text = text.empty() ? nullptr : text.c_str();
  const emscripten::val time_period = reader.value(v, "timePeriod", "rule.timePeriod");
  if (js_value_present(time_period)) {
    rule.has_time_period = 1;
    rule.time_period = reader.u8_value(time_period, 0U, "rule.timePeriod");
  }
  const emscripten::val color_scale = reader.value(v, "colorScale", "colorScale");
  if (js_value_present(color_scale)) {
    const emscripten::val thresholds = reader.value(color_scale, "thresholds", "colorScale.thresholds");
    if (reader.is_array(thresholds, "colorScale.thresholds")) {
      const uint32_t n = reader.length(thresholds, "colorScale.thresholds");
      color_scale_thresholds.reserve(n);
      for (uint32_t i = 0; i < n; ++i) {
        color_scale_thresholds.push_back(
            js_pull_cfvo(reader.array_element(thresholds, i, "colorScale.thresholds[]"), &cfvo_strings, reader));
      }
    }
    const emscripten::val colors = reader.value(color_scale, "colors", "colorScale.colors");
    if (reader.is_array(colors, "colorScale.colors")) {
      const uint32_t n = reader.length(colors, "colorScale.colors");
      color_scale_colors.reserve(n);
      for (uint32_t i = 0; i < n; ++i) {
        color_scale_colors.push_back(js_pull_color_spec_value(reader.array_element(colors, i, "colorScale.colors[]"),
                                                              "colorScale.colors[]", &reader));
      }
    }
    rule.color_scale_thresholds = color_scale_thresholds.empty() ? nullptr : color_scale_thresholds.data();
    rule.color_scale_colors = color_scale_colors.empty() ? nullptr : color_scale_colors.data();
    rule.color_scale_count = static_cast<uint32_t>(color_scale_thresholds.size());
  }
  const emscripten::val data_bar = reader.value(v, "dataBar", "dataBar");
  if (js_value_present(data_bar)) {
    rule.has_data_bar = 1;
    rule.data_bar_min = js_pull_cfvo(reader.value(data_bar, "min", "dataBar.min"), &cfvo_strings, reader);
    rule.data_bar_max = js_pull_cfvo(reader.value(data_bar, "max", "dataBar.max"), &cfvo_strings, reader);
    rule.data_bar_fill = js_pull_color_spec(data_bar, "fill", &reader);
    rule.data_bar_show_value = reader.boolean(data_bar, "showValue", true, "dataBar.showValue") ? 1 : 0;
    rule.data_bar_min_length_pct = reader.u8(data_bar, "minLengthPct", 10U, "dataBar.minLengthPct");
    rule.data_bar_max_length_pct = reader.u8(data_bar, "maxLengthPct", 90U, "dataBar.maxLengthPct");
    // `x14` extension payload. An omitted key leaves the corresponding
    // `has_*` flag clear, which the C ABI reads as "keep the model
    // default" (gradient on, automatic axis, negative fill equal to the
    // positive fill, no border, black axis).
    const emscripten::val gradient = reader.value(data_bar, "gradient", "dataBar.gradient");
    if (js_value_present(gradient)) {
      rule.has_data_bar_gradient = 1;
      rule.data_bar_gradient = reader.boolean_value(gradient, true) ? 1 : 0;
    }
    const emscripten::val axis_position = reader.value(data_bar, "axisPosition", "dataBar.axisPosition");
    if (js_value_present(axis_position)) {
      rule.has_data_bar_axis_position = 1;
      rule.data_bar_axis_position = reader.u8_value(axis_position, 0U, "dataBar.axisPosition");
    }
    const emscripten::val negative_fill = reader.value(data_bar, "negativeFill", "dataBar.negativeFill");
    if (js_value_present(negative_fill)) {
      rule.data_bar_negative_fill = js_pull_color_spec_value(negative_fill, "dataBar.negativeFill", &reader);
    }
    const emscripten::val border = reader.value(data_bar, "border", "dataBar.border");
    if (js_value_present(border)) {
      rule.data_bar_border = js_pull_color_spec_value(border, "dataBar.border", &reader);
    }
    const emscripten::val negative_border = reader.value(data_bar, "negativeBorder", "dataBar.negativeBorder");
    if (js_value_present(negative_border)) {
      rule.data_bar_negative_border = js_pull_color_spec_value(negative_border, "dataBar.negativeBorder", &reader);
    }
    const emscripten::val axis_color = reader.value(data_bar, "axisColor", "dataBar.axisColor");
    if (js_value_present(axis_color)) {
      rule.data_bar_axis_color = js_pull_color_spec_value(axis_color, "dataBar.axisColor", &reader);
    }
    rule.data_bar_direction = reader.u8(data_bar, "direction", 0U, "dataBar.direction");
  }
  const emscripten::val icon_set = reader.value(v, "iconSet", "iconSet");
  if (js_value_present(icon_set)) {
    rule.has_icon_set = 1;
    rule.icon_set = reader.u8(icon_set, "name", 0U, "iconSet.name");
    const emscripten::val thresholds = reader.value(icon_set, "thresholds", "iconSet.thresholds");
    if (reader.is_array(thresholds, "iconSet.thresholds")) {
      const uint32_t n = reader.length(thresholds, "iconSet.thresholds");
      icon_set_thresholds.reserve(n);
      for (uint32_t i = 0; i < n; ++i) {
        icon_set_thresholds.push_back(
            js_pull_cfvo(reader.array_element(thresholds, i, "iconSet.thresholds[]"), &cfvo_strings, reader));
      }
    }
    rule.icon_set_thresholds = icon_set_thresholds.empty() ? nullptr : icon_set_thresholds.data();
    rule.icon_set_threshold_count = static_cast<uint32_t>(icon_set_thresholds.size());
    rule.icon_set_reverse = reader.boolean(icon_set, "reverse", false, "iconSet.reverse") ? 1 : 0;
    rule.icon_set_show_value = reader.boolean(icon_set, "showValue", true, "iconSet.showValue") ? 1 : 0;
    rule.icon_set_percent = reader.boolean(icon_set, "percent", true, "iconSet.percent") ? 1 : 0;
    // An omitted floor keeps Excel's default, `percent 0`.
    const emscripten::val floor = reader.value(icon_set, "floor", "iconSet.floor");
    if (js_value_present(floor)) {
      rule.has_icon_set_floor = 1;
      rule.icon_set_floor = js_pull_cfvo(floor, &cfvo_strings, reader);
    }
    const emscripten::val icons = reader.value(icon_set, "icons", "iconSet.icons");
    if (js_value_present(icons) && reader.is_array(icons, "iconSet.icons")) {
      const uint32_t n = reader.length(icons, "iconSet.icons");
      icon_set_icons.reserve(n);
      for (uint32_t i = 0; i < n; ++i) {
        const emscripten::val icon = reader.array_element(icons, i, "iconSet.icons[]");
        icon_set_icons.push_back(fm_cf_icon_t{reader.i32(icon, "iconSet", 0, "iconSet.icons[].iconSet"),
                                              reader.i32(icon, "iconIndex", 0, "iconSet.icons[].iconIndex")});
      }
      rule.icon_set_icons = icon_set_icons.empty() ? nullptr : icon_set_icons.data();
      rule.icon_set_icon_count = static_cast<uint32_t>(icon_set_icons.size());
    }
  }
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }

  const char* out_id = nullptr;
  if (!succeeded(fm_sheet_add_cf_rule(handle_, sheet, &rule, &out_id))) {
    return emscripten::val::undefined();
  }
  return emscripten::val(string_from_cstr(out_id));
}

void JsWorkbook::updateCfRule(uint32_t sheet, const std::string& cfRuleId, emscripten::val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Sheet.updateCfRule");
  formulon::c_api::BorrowedStringArena cfvo_strings;
  std::vector<fm_cfvo_t> color_scale_thresholds;
  std::vector<fm_color_t> color_scale_colors;
  std::vector<fm_cfvo_t> icon_set_thresholds;
  std::vector<fm_cf_icon_t> icon_set_icons;
  std::vector<fm_cell_rect_t> ranges;
  std::string text[3];
  fm_cf_rule_patch_t p = FM_STRUCT_INIT(fm_cf_rule_patch_t);

  const auto u32 = [&](const emscripten::val& owner, const char* key, const char* field, fm_patch_op_t* op,
                       uint32_t* value, bool nullable) {
    const emscripten::val v = reader.value(owner, key, field);
    switch (js_patch_key(reader, v, field, nullable)) {
      case JsPatchKey::kClear:
        *op = FM_PATCH_OP_CLEAR;
        break;
      case JsPatchKey::kSet:
        *op = FM_PATCH_OP_SET;
        *value = reader.u32_value(v, 0U, field);
        break;
      case JsPatchKey::kKeep:
        break;
    }
  };
  const auto i32 = [&](const emscripten::val& owner, const char* key, const char* field, fm_patch_op_t* op,
                       int32_t* value) { js_patch_int(reader, owner, key, field, false, op, value); };
  const auto u8 = [&](const emscripten::val& owner, const char* key, const char* field, fm_patch_op_t* op,
                      int32_t* value) {
    const emscripten::val v = reader.value(owner, key, field);
    if (js_patch_key(reader, v, field, false) == JsPatchKey::kSet) {
      *op = FM_PATCH_OP_SET;
      *value = static_cast<int32_t>(reader.u8_value(v, 0U, field));
    }
  };
  const auto u8_pct = [&](const emscripten::val& owner, const char* key, const char* field, fm_patch_op_t* op,
                          uint32_t* value) {
    const emscripten::val v = reader.value(owner, key, field);
    if (js_patch_key(reader, v, field, false) == JsPatchKey::kSet) {
      *op = FM_PATCH_OP_SET;
      *value = reader.u8_value(v, 0U, field);
    }
  };
  const auto cfvo = [&](const emscripten::val& owner, const char* key, const char* field, fm_patch_op_t* op,
                        fm_cfvo_t* value) {
    const emscripten::val v = reader.value(owner, key, field);
    if (js_patch_key(reader, v, field, false) == JsPatchKey::kSet) {
      *op = FM_PATCH_OP_SET;
      *value = js_pull_cfvo(v, &cfvo_strings, reader);
    }
  };
  // `null` on an optional bar colour restores its default, which the C patch spells as a colour of kind none.
  const auto color = [&](const emscripten::val& owner, const char* key, const char* field, fm_patch_op_t* op,
                         fm_color_t* value, bool nullable) {
    const emscripten::val v = reader.value(owner, key, field);
    if (v.isUndefined()) {
      return;
    }
    if (v.isNull() && !nullable) {
      js_patch_key(reader, v, field, false);
      return;
    }
    *op = FM_PATCH_OP_SET;
    *value = js_pull_color_spec_value(v, field, &reader);
  };
  const auto cfvo_list = [&](const emscripten::val& owner, const char* key, const char* field, fm_patch_op_t* op,
                             std::vector<fm_cfvo_t>& store, const fm_cfvo_t** data, uint32_t* count) {
    const emscripten::val v = reader.value(owner, key, field);
    if (js_patch_key(reader, v, field, false) != JsPatchKey::kSet || !reader.is_array(v, field)) {
      return;
    }
    const uint32_t n = reader.length(v, field);
    store.reserve(n);
    for (uint32_t i = 0; i < n; ++i) {
      store.push_back(js_pull_cfvo(reader.array_element(v, i, field), &cfvo_strings, reader));
    }
    *op = FM_PATCH_OP_SET;
    *data = store.empty() ? nullptr : store.data();
    *count = static_cast<uint32_t>(store.size());
  };
  // A nested payload patch: an object whose own keys are the payload's fields.
  const auto nested = [&](const char* key, const char* field) {
    const emscripten::val v = reader.value(patch, key, field);
    if (js_patch_key(reader, v, field, false) != JsPatchKey::kSet || !reader.require_object(v, field)) {
      return emscripten::val::undefined();
    }
    return v;
  };

  i32(patch, "priority", "priority", &p.priority_op, &p.priority);
  js_patch_flag(reader, patch, "stopIfTrue", "stopIfTrue", &p.stop_if_true_op, &p.stop_if_true);
  u32(patch, "dxfIndex", "dxfIndex", &p.dxf_index_op, &p.dxf_index, true);
  const emscripten::val ranges_value = reader.value(patch, "ranges", "ranges");
  if (reader.ok() && js_patch_key(reader, ranges_value, "ranges", false) == JsPatchKey::kSet &&
      reader.require_array(ranges_value, "ranges")) {
    ranges = js_pull_range_array(ranges_value, "ranges", &reader);
    p.ranges_op = FM_PATCH_OP_SET;
    p.ranges = ranges.empty() ? nullptr : ranges.data();
    p.range_count = ranges.size();
  }
  js_patch_string(reader, patch, "formula1", "formula1", true, text[0], &p.formula1_op, &p.formula1);
  js_patch_string(reader, patch, "formula2", "formula2", true, text[1], &p.formula2_op, &p.formula2);
  u8(patch, "operator", "operator", &p.op_op, &p.op);
  i32(patch, "rank", "rank", &p.rank_op, &p.rank);
  js_patch_flag(reader, patch, "percent", "percent", &p.percent_op, &p.percent);
  js_patch_flag(reader, patch, "bottom", "bottom", &p.bottom_op, &p.bottom);
  js_patch_flag(reader, patch, "aboveAverage", "aboveAverage", &p.above_average_op, &p.above_average);
  js_patch_flag(reader, patch, "equalAverage", "equalAverage", &p.equal_average_op, &p.equal_average);
  js_patch_number(reader, patch, "stdDev", "stdDev", false, &p.std_dev_op, &p.std_dev);
  js_patch_string(reader, patch, "text", "text", false, text[2], &p.text_op, &p.text);
  u8(patch, "timePeriod", "timePeriod", &p.time_period_op, &p.time_period);

  const emscripten::val color_scale = nested("colorScale", "colorScale");
  if (!color_scale.isUndefined()) {
    cfvo_list(color_scale, "thresholds", "colorScale.thresholds", &p.color_scale_thresholds_op, color_scale_thresholds,
              &p.color_scale_thresholds, &p.color_scale_thresholds_count);
    const emscripten::val colors = reader.value(color_scale, "colors", "colorScale.colors");
    if (js_patch_key(reader, colors, "colorScale.colors", false) == JsPatchKey::kSet &&
        reader.is_array(colors, "colorScale.colors")) {
      const uint32_t n = reader.length(colors, "colorScale.colors");
      color_scale_colors.reserve(n);
      for (uint32_t i = 0; i < n; ++i) {
        color_scale_colors.push_back(js_pull_color_spec_value(reader.array_element(colors, i, "colorScale.colors[]"),
                                                              "colorScale.colors[]", &reader));
      }
      p.color_scale_colors_op = FM_PATCH_OP_SET;
      p.color_scale_colors = color_scale_colors.empty() ? nullptr : color_scale_colors.data();
      p.color_scale_colors_count = static_cast<uint32_t>(color_scale_colors.size());
    }
  }
  const emscripten::val data_bar = nested("dataBar", "dataBar");
  if (!data_bar.isUndefined()) {
    cfvo(data_bar, "min", "dataBar.min", &p.data_bar_min_op, &p.data_bar_min);
    cfvo(data_bar, "max", "dataBar.max", &p.data_bar_max_op, &p.data_bar_max);
    color(data_bar, "fill", "dataBar.fill", &p.data_bar_fill_op, &p.data_bar_fill, false);
    js_patch_flag(reader, data_bar, "showValue", "dataBar.showValue", &p.data_bar_show_value_op,
                  &p.data_bar_show_value);
    u8_pct(data_bar, "minLengthPct", "dataBar.minLengthPct", &p.data_bar_min_length_pct_op, &p.data_bar_min_length_pct);
    u8_pct(data_bar, "maxLengthPct", "dataBar.maxLengthPct", &p.data_bar_max_length_pct_op, &p.data_bar_max_length_pct);
    // `null` on these three is "not specified": the patch sets the default the add path applies.
    const emscripten::val gradient = reader.value(data_bar, "gradient", "dataBar.gradient");
    if (gradient.isNull()) {
      p.data_bar_gradient_op = FM_PATCH_OP_SET;
      p.data_bar_gradient = 1;
    } else {
      js_patch_flag(reader, data_bar, "gradient", "dataBar.gradient", &p.data_bar_gradient_op, &p.data_bar_gradient);
    }
    if (reader.value(data_bar, "axisPosition", "dataBar.axisPosition").isNull()) {
      p.data_bar_axis_position_op = FM_PATCH_OP_SET;
      p.data_bar_axis_position = FM_CF_DATA_BAR_AXIS_POSITION_AUTOMATIC;
    } else {
      u8(data_bar, "axisPosition", "dataBar.axisPosition", &p.data_bar_axis_position_op, &p.data_bar_axis_position);
    }
    color(data_bar, "negativeFill", "dataBar.negativeFill", &p.data_bar_negative_fill_op, &p.data_bar_negative_fill,
          true);
    color(data_bar, "border", "dataBar.border", &p.data_bar_border_op, &p.data_bar_border, true);
    color(data_bar, "negativeBorder", "dataBar.negativeBorder", &p.data_bar_negative_border_op,
          &p.data_bar_negative_border, true);
    color(data_bar, "axisColor", "dataBar.axisColor", &p.data_bar_axis_color_op, &p.data_bar_axis_color, true);
    u8(data_bar, "direction", "dataBar.direction", &p.data_bar_direction_op, &p.data_bar_direction);
  }
  const emscripten::val icon_set = nested("iconSet", "iconSet");
  if (!icon_set.isUndefined()) {
    u8(icon_set, "name", "iconSet.name", &p.icon_set_op, &p.icon_set);
    cfvo_list(icon_set, "thresholds", "iconSet.thresholds", &p.icon_set_thresholds_op, icon_set_thresholds,
              &p.icon_set_thresholds, &p.icon_set_thresholds_count);
    js_patch_flag(reader, icon_set, "reverse", "iconSet.reverse", &p.icon_set_reverse_op, &p.icon_set_reverse);
    js_patch_flag(reader, icon_set, "showValue", "iconSet.showValue", &p.icon_set_show_value_op,
                  &p.icon_set_show_value);
    js_patch_flag(reader, icon_set, "percent", "iconSet.percent", &p.icon_set_percent_op, &p.icon_set_percent);
    if (reader.value(icon_set, "floor", "iconSet.floor").isNull()) {
      p.icon_set_floor_op = FM_PATCH_OP_SET;
      p.icon_set_floor = fm_cfvo_t{FM_CF_VALUE_OBJECT_TYPE_PERCENT, FM_TRISTATE_TRUE, nullptr};
    } else {
      cfvo(icon_set, "floor", "iconSet.floor", &p.icon_set_floor_op, &p.icon_set_floor);
    }
    const emscripten::val icons = reader.value(icon_set, "icons", "iconSet.icons");
    if (!icons.isNull() && js_patch_key(reader, icons, "iconSet.icons", false) == JsPatchKey::kSet &&
        reader.is_array(icons, "iconSet.icons")) {
      const uint32_t n = reader.length(icons, "iconSet.icons");
      icon_set_icons.reserve(n);
      for (uint32_t i = 0; i < n; ++i) {
        const emscripten::val icon = reader.array_element(icons, i, "iconSet.icons[]");
        icon_set_icons.push_back(fm_cf_icon_t{reader.i32(icon, "iconSet", 0, "iconSet.icons[].iconSet"),
                                              reader.i32(icon, "iconIndex", 0, "iconSet.icons[].iconIndex")});
      }
      p.icon_set_icons_op = FM_PATCH_OP_SET;
      p.icon_set_icons = icon_set_icons.empty() ? nullptr : icon_set_icons.data();
      p.icon_set_icons_count = static_cast<uint32_t>(icon_set_icons.size());
    }
  }
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_cf_rule(handle_, sheet, cfRuleId.c_str(), &p));
}

void JsWorkbook::removeConditionalFormatAt(uint32_t sheet, const std::string& cfRuleId) {
  succeeded(handle_ != nullptr ? fm_sheet_remove_cf_rule(handle_, sheet, cfRuleId.c_str()) : kBindingInvalidHandle);
}

void JsWorkbook::clearConditionalFormats(uint32_t sheet) {
  succeeded(handle_ != nullptr ? fm_sheet_clear_cf_rules(handle_, sheet) : kBindingInvalidHandle);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
