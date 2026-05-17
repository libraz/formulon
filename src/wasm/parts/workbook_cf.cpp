// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// JsWorkbook conditional-formatting surface: read / add / remove / clear
// rules plus `evaluateCfRange` which projects matched rules onto a
// requested cell window. Visual rule kinds (`ColorScale`, `DataBar`,
// `IconSet`) round-trip through OOXML reader / writer but cannot be
// constructed via the embind add path yet -- the read surface returns
// `type` populated but with the visual sub-spec fields omitted.

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
    arr.set(static_cast<uint32_t>(i), item);
  }
  return arr;
}

JsStatus JsWorkbook::addConditionalFormat(uint32_t sheet, emscripten::val v) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  // Pull every JS field into local storage; the C ABI receives
  // borrowed `const char*` views that must stay valid for the
  // duration of the call.
  std::vector<fm_cf_cell_range_t> ranges_buf;
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
  fm_status_t rc = fm_sheet_cf_add_rule(handle_, sheet, rule);
  return status_from_rc(rc);
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
