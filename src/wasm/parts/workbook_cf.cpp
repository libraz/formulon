// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {
namespace {

fm_cf_color_t js_pull_cf_color(emscripten::val v) {
  fm_cf_color_t out{};
  out.r = js_pull_u8(v, "r", 0U);
  out.g = js_pull_u8(v, "g", 0U);
  out.b = js_pull_u8(v, "b", 0U);
  out.a = js_pull_u8(v, "a", 255U);
  return out;
}

fm_cfvo_t js_pull_cfvo(emscripten::val v, std::vector<std::string>* strings) {
  fm_cfvo_t out{};
  out.type = js_pull_u8(v, "type", 0U);
  out.gte = js_pull_bool(v, "gte", true) ? 1 : 0;
  if (!v["value"].isUndefined() && !v["value"].isNull()) {
    strings->push_back(v["value"].as<std::string>());
    out.value = strings->back().c_str();
  }
  return out;
}

emscripten::val cf_color_to_js(fm_cf_color_t color) {
  emscripten::val out = emscripten::val::object();
  out.set("r", color.r);
  out.set("g", color.g);
  out.set("b", color.b);
  out.set("a", color.a);
  return out;
}

emscripten::val cfvo_to_js(const fm_cfvo_t& cfvo) {
  emscripten::val out = emscripten::val::object();
  out.set("type", static_cast<uint32_t>(cfvo.type));
  out.set("gte", cfvo.gte != 0);
  if (cfvo.value != nullptr) {
    out.set("value", std::string(cfvo.value));
  }
  return out;
}

}  // namespace

JsCfRangeResult JsWorkbook::evaluateCfRange(uint32_t sheet, uint32_t firstRow, uint32_t firstCol, uint32_t lastRow,
                                            uint32_t lastCol, double todaySerial) const {
  JsCfRangeResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  fm_cf_results_t* results = nullptr;
  fm_status_t rc =
      fm_workbook_cf_evaluate_range(handle_, sheet, firstRow, firstCol, lastRow, lastCol, todaySerial, &results);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  const std::size_t cell_count = fm_cf_results_cell_count(results);
  r.cells.reserve(cell_count);
  for (std::size_t i = 0; i < cell_count; ++i) {
    JsCfCellResult cell;
    uint32_t row = 0;
    uint32_t col = 0;
    std::size_t match_count = 0;
    if (fm_cf_results_cell_at(results, i, &row, &col, &match_count) != 0) {
      // Skip entries the C ABI declines to materialise; this is purely
      // defensive -- the index is always valid by construction.
      continue;
    }
    cell.row = row;
    cell.col = col;
    cell.matches.reserve(match_count);
    for (std::size_t j = 0; j < match_count; ++j) {
      fm_cf_match_t m{};
      if (fm_cf_results_match_at(results, i, j, &m) != 0) {
        continue;
      }
      cell.matches.push_back(translate_cf_match(m));
    }
    r.cells.push_back(std::move(cell));
  }
  fm_cf_results_destroy(results);
  r.status = ok_status();
  return r;
}

emscripten::val JsWorkbook::getConditionalFormats(uint32_t sheet) const {
  emscripten::val arr = emscripten::val::array();
  if (handle_ == nullptr) {
    return arr;
  }
  std::size_t count = 0;
  if (fm_sheet_cf_count(handle_, sheet, &count) != 0) {
    return arr;
  }
  for (std::size_t i = 0; i < count; ++i) {
    fm_cf_rule_t rule{};
    if (fm_sheet_cf_get_at(handle_, sheet, i, &rule) != 0) {
      continue;
    }
    emscripten::val item = emscripten::val::object();
    item.set("id", rule.id != nullptr ? std::string(rule.id) : std::string());
    item.set("type", static_cast<uint32_t>(rule.type));
    item.set("priority", rule.priority);
    item.set("stopIfTrue", rule.stop_if_true != 0);
    emscripten::val sqref = emscripten::val::array();
    for (uint32_t r = 0; r < rule.sqref_count; ++r) {
      emscripten::val rng = emscripten::val::object();
      rng.set("firstRow", rule.sqref[r].first_row);
      rng.set("firstCol", rule.sqref[r].first_col);
      rng.set("lastRow", rule.sqref[r].last_row);
      rng.set("lastCol", rule.sqref[r].last_col);
      sqref.set(r, rng);
    }
    item.set("sqref", sqref);
    if (rule.dxf_id_engaged != 0) {
      item.set("dxfId", rule.dxf_id);
    }
    if (rule.formula1 != nullptr) {
      item.set("formula1", std::string(rule.formula1));
    }
    if (rule.formula2 != nullptr) {
      item.set("formula2", std::string(rule.formula2));
    }
    if (rule.op_engaged != 0) {
      item.set("op", static_cast<uint32_t>(rule.op));
    }
    if (rule.rank_engaged != 0) {
      item.set("rank", rule.rank);
      item.set("percent", rule.percent != 0);
      item.set("bottom", rule.bottom != 0);
    }
    // aboveAverage flags are always present (default-engineered), but
    // we only surface them for the AboveAverage rule type to avoid
    // confusing FE consumers.
    if (rule.type == 6 /* AboveAverage */) {
      item.set("aboveAverage", rule.above_average != 0);
      item.set("equalAverage", rule.equal_average != 0);
      if (rule.std_dev_engaged != 0) {
        item.set("stdDev", rule.std_dev);
      }
    }
    if (rule.text != nullptr) {
      item.set("text", std::string(rule.text));
    }
    if (rule.time_period_engaged != 0) {
      item.set("timePeriod", static_cast<uint32_t>(rule.time_period));
    }
    if (rule.color_scale_count > 0 && rule.color_scale_thresholds != nullptr && rule.color_scale_colors != nullptr) {
      emscripten::val color_scale = emscripten::val::object();
      emscripten::val thresholds = emscripten::val::array();
      emscripten::val colors = emscripten::val::array();
      for (uint32_t j = 0; j < rule.color_scale_count; ++j) {
        thresholds.set(j, cfvo_to_js(rule.color_scale_thresholds[j]));
        colors.set(j, cf_color_to_js(rule.color_scale_colors[j]));
      }
      color_scale.set("thresholds", thresholds);
      color_scale.set("colors", colors);
      item.set("colorScale", color_scale);
    }
    if (rule.data_bar_engaged != 0) {
      emscripten::val data_bar = emscripten::val::object();
      data_bar.set("min", cfvo_to_js(rule.data_bar_min));
      data_bar.set("max", cfvo_to_js(rule.data_bar_max));
      data_bar.set("fill", cf_color_to_js(rule.data_bar_fill));
      data_bar.set("showValue", rule.data_bar_show_value != 0);
      data_bar.set("minLengthPct", static_cast<uint32_t>(rule.data_bar_min_length_pct));
      data_bar.set("maxLengthPct", static_cast<uint32_t>(rule.data_bar_max_length_pct));
      item.set("dataBar", data_bar);
    }
    if (rule.icon_set_engaged != 0) {
      emscripten::val icon_set = emscripten::val::object();
      icon_set.set("name", static_cast<uint32_t>(rule.icon_set_name));
      emscripten::val thresholds = emscripten::val::array();
      for (uint32_t j = 0; j < rule.icon_set_threshold_count; ++j) {
        thresholds.set(j, cfvo_to_js(rule.icon_set_thresholds[j]));
      }
      icon_set.set("thresholds", thresholds);
      icon_set.set("reverse", rule.icon_set_reverse != 0);
      icon_set.set("showValue", rule.icon_set_show_value != 0);
      icon_set.set("percent", rule.icon_set_percent != 0);
      item.set("iconSet", icon_set);
    }
    arr.set(static_cast<uint32_t>(i), item);
  }
  return arr;
}

JsAddStyleResult JsWorkbook::addConditionalFormat(uint32_t sheet, emscripten::val v) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  // Pull every JS field into local storage; the C ABI receives
  // borrowed `const char*` views that must stay valid for the
  // duration of the call.
  std::vector<fm_cf_cell_range_t> ranges_buf;
  std::vector<fm_cfvo_t> color_scale_thresholds;
  std::vector<fm_cf_color_t> color_scale_colors;
  std::vector<fm_cfvo_t> icon_set_thresholds;
  std::vector<std::string> cfvo_strings;
  if (v.hasOwnProperty("sqref")) {
    emscripten::val sqref_js = v["sqref"];
    if (sqref_js.isArray()) {
      const uint32_t n = sqref_js["length"].as<uint32_t>();
      ranges_buf.reserve(n);
      for (uint32_t i = 0; i < n; ++i) {
        emscripten::val rng = sqref_js[i];
        fm_cf_cell_range_t r{};
        r.first_row = rng["firstRow"].as<uint32_t>();
        r.first_col = rng["firstCol"].as<uint32_t>();
        r.last_row = rng["lastRow"].as<uint32_t>();
        r.last_col = rng["lastCol"].as<uint32_t>();
        ranges_buf.push_back(r);
      }
    }
  }
  const std::string id = js_pull_string(v, "id");
  const std::string formula1 = js_pull_string(v, "formula1");
  const std::string formula2 = js_pull_string(v, "formula2");
  const std::string text = js_pull_string(v, "text");

  fm_cf_rule_t rule{};
  rule.id = id.empty() ? nullptr : id.c_str();
  rule.type = js_pull_u8(v, "type", 0U);
  rule.priority = js_pull_u32(v, "priority", 0U) != 0 ? static_cast<int32_t>(js_pull_u32(v, "priority", 0U)) : 0;
  rule.stop_if_true = js_pull_bool(v, "stopIfTrue", false) ? 1 : 0;
  if (!v["dxfId"].isUndefined() && !v["dxfId"].isNull()) {
    rule.dxf_id_engaged = 1;
    rule.dxf_id = v["dxfId"].as<uint32_t>();
  }
  rule.sqref = ranges_buf.empty() ? nullptr : ranges_buf.data();
  rule.sqref_count = static_cast<uint32_t>(ranges_buf.size());
  rule.formula1 = formula1.empty() ? nullptr : formula1.c_str();
  rule.formula2 = formula2.empty() ? nullptr : formula2.c_str();
  if (!v["op"].isUndefined() && !v["op"].isNull()) {
    rule.op_engaged = 1;
    rule.op = js_pull_u8(v, "op", 0U);
  }
  if (!v["rank"].isUndefined() && !v["rank"].isNull()) {
    rule.rank_engaged = 1;
    rule.rank = static_cast<int32_t>(v["rank"].as<int32_t>());
  }
  rule.percent = js_pull_bool(v, "percent", false) ? 1 : 0;
  rule.bottom = js_pull_bool(v, "bottom", false) ? 1 : 0;
  rule.above_average = js_pull_bool(v, "aboveAverage", true) ? 1 : 0;
  rule.equal_average = js_pull_bool(v, "equalAverage", false) ? 1 : 0;
  if (!v["stdDev"].isUndefined() && !v["stdDev"].isNull()) {
    rule.std_dev_engaged = 1;
    rule.std_dev = v["stdDev"].as<double>();
  }
  rule.text = text.empty() ? nullptr : text.c_str();
  if (!v["timePeriod"].isUndefined() && !v["timePeriod"].isNull()) {
    rule.time_period_engaged = 1;
    rule.time_period = js_pull_u8(v, "timePeriod", 0U);
  }
  if (!v["colorScale"].isUndefined() && !v["colorScale"].isNull()) {
    emscripten::val cs = v["colorScale"];
    if (cs.hasOwnProperty("thresholds") && cs["thresholds"].isArray()) {
      const uint32_t n = cs["thresholds"]["length"].as<uint32_t>();
      color_scale_thresholds.reserve(n);
      for (uint32_t i = 0; i < n; ++i) {
        color_scale_thresholds.push_back(js_pull_cfvo(cs["thresholds"][i], &cfvo_strings));
      }
    }
    if (cs.hasOwnProperty("colors") && cs["colors"].isArray()) {
      const uint32_t n = cs["colors"]["length"].as<uint32_t>();
      color_scale_colors.reserve(n);
      for (uint32_t i = 0; i < n; ++i) {
        color_scale_colors.push_back(js_pull_cf_color(cs["colors"][i]));
      }
    }
    rule.color_scale_thresholds = color_scale_thresholds.empty() ? nullptr : color_scale_thresholds.data();
    rule.color_scale_colors = color_scale_colors.empty() ? nullptr : color_scale_colors.data();
    rule.color_scale_count = static_cast<uint32_t>(color_scale_thresholds.size());
  }
  if (!v["dataBar"].isUndefined() && !v["dataBar"].isNull()) {
    emscripten::val db = v["dataBar"];
    rule.data_bar_engaged = 1;
    rule.data_bar_min = js_pull_cfvo(db["min"], &cfvo_strings);
    rule.data_bar_max = js_pull_cfvo(db["max"], &cfvo_strings);
    rule.data_bar_fill = js_pull_cf_color(db["fill"]);
    rule.data_bar_show_value = js_pull_bool(db, "showValue", true) ? 1 : 0;
    rule.data_bar_min_length_pct = js_pull_u8(db, "minLengthPct", 10U);
    rule.data_bar_max_length_pct = js_pull_u8(db, "maxLengthPct", 90U);
  }
  if (!v["iconSet"].isUndefined() && !v["iconSet"].isNull()) {
    emscripten::val is = v["iconSet"];
    rule.icon_set_engaged = 1;
    rule.icon_set_name = js_pull_u8(is, "name", 0U);
    if (is.hasOwnProperty("thresholds") && is["thresholds"].isArray()) {
      const uint32_t n = is["thresholds"]["length"].as<uint32_t>();
      icon_set_thresholds.reserve(n);
      for (uint32_t i = 0; i < n; ++i) {
        icon_set_thresholds.push_back(js_pull_cfvo(is["thresholds"][i], &cfvo_strings));
      }
    }
    rule.icon_set_thresholds = icon_set_thresholds.empty() ? nullptr : icon_set_thresholds.data();
    rule.icon_set_threshold_count = static_cast<uint32_t>(icon_set_thresholds.size());
    rule.icon_set_reverse = js_pull_bool(is, "reverse", false) ? 1 : 0;
    rule.icon_set_show_value = js_pull_bool(is, "showValue", true) ? 1 : 0;
    rule.icon_set_percent = js_pull_bool(is, "percent", true) ? 1 : 0;
  }
  std::size_t new_index = 0;
  fm_status_t rc = fm_sheet_cf_add_rule(handle_, sheet, rule, &new_index);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.status = ok_status();
  r.index = static_cast<uint32_t>(new_index);
  return r;
}

JsStatus JsWorkbook::removeConditionalFormatAt(uint32_t sheet, uint32_t index) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_cf_remove_at(handle_, sheet, index);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::clearConditionalFormats(uint32_t sheet) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_cf_clear(handle_, sheet);
  return status_from_rc(rc);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
