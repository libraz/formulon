// Conditional-formatting bindings: the `evaluateCfRange` evaluator plus the
// read / mutate surface for per-sheet conditional-format rules.

#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <vector>

#include "c_api/borrowed_arena.h"
#include "node_addon/parts/workbook_class.h"

namespace formulon_node {
namespace {

fm_cf_color_t PullCfColor(CheckedSpecReader& reader, const Napi::Object& spec) {
  fm_cf_color_t out{};
  out.r = reader.U8(spec, "r", 0U);
  out.g = reader.U8(spec, "g", 0U);
  out.b = reader.U8(spec, "b", 0U);
  out.a = reader.U8(spec, "a", 255U);
  return out;
}

fm_cfvo_t PullCfvo(CheckedSpecReader& reader, const Napi::Object& spec, formulon::c_api::BorrowedStringArena* strings) {
  fm_cfvo_t out{};
  out.type = reader.U8(spec, "type", 0U);
  out.gte = reader.Bool(spec, "gte", true) ? 1 : 0;
  std::string value;
  if (reader.String(spec, "value", &value)) {
    out.value = strings->emplace(value);
  }
  return out;
}

Napi::Object CfvoToJs(Napi::Env env, const fm_cfvo_t& cfvo) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("type", Napi::Number::New(env, static_cast<uint32_t>(cfvo.type)));
  out.Set("gte", Napi::Boolean::New(env, cfvo.gte != 0));
  if (cfvo.value != nullptr) {
    out.Set("value", Napi::String::New(env, cfvo.value));
  }
  return out;
}

}  // namespace

// ---- Conditional formatting -----------------------------------------

Napi::Value Workbook::EvaluateCfRange(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeFieldResult(env, NullHandleError(env), "cells", Napi::Array::New(env));
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t first_row = ArgU32(info, 1);
  const uint32_t first_col = ArgU32(info, 2);
  const uint32_t last_row = ArgU32(info, 3);
  const uint32_t last_col = ArgU32(info, 4);
  // A missing `todaySerial` must disable `timePeriod` rules, matching the
  // Python binding and the C ABI contract. `ArgDouble`'s 0.0 fallback is a
  // valid Excel serial (1899-12-30), so it would silently evaluate those
  // rules against that date instead of skipping them.
  const double today_serial =
      info.Length() > 5 ? info[5].ToNumber().DoubleValue() : std::numeric_limits<double>::quiet_NaN();
  fm_cf_results_t* results = nullptr;
  fm_status_t rc =
      fm_workbook_cf_evaluate_range(handle_, sheet, first_row, first_col, last_row, last_col, today_serial, &results);
  if (rc != 0) {
    return MakeFieldResult(env, MakeErrorStatus(env, rc), "cells", Napi::Array::New(env));
  }
  const std::size_t cell_count = fm_cf_results_cell_count(results);
  Napi::Array cells = Napi::Array::New(env, cell_count);
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < cell_count; ++i) {
    uint32_t row = 0;
    uint32_t col = 0;
    std::size_t match_count = 0;
    if (fm_cf_results_cell_at(results, i, &row, &col, &match_count) != 0) {
      // Defensive: skip entries the C ABI declines to materialise.
      continue;
    }
    Napi::Array matches = Napi::Array::New(env, match_count);
    std::size_t mj = 0;
    for (std::size_t j = 0; j < match_count; ++j) {
      fm_cf_match_t m{};
      if (fm_cf_results_match_at(results, i, j, &m) != 0) {
        continue;
      }
      matches.Set(static_cast<uint32_t>(mj), TranslateCfMatch(env, m));
      ++mj;
    }
    Napi::Object cell = Napi::Object::New(env);
    cell.Set("row", Napi::Number::New(env, row));
    cell.Set("col", Napi::Number::New(env, col));
    cell.Set("matches", matches);
    cells.Set(static_cast<uint32_t>(emitted), cell);
    ++emitted;
  }
  fm_cf_results_destroy(results);
  return MakeFieldResult(env, MakeOkStatus(env), "cells", cells);
}

// ---- Conditional formatting (read / mutate) -------------------------

Napi::Value Workbook::GetConditionalFormats(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return FinishListResult(env, arr, kBindingInvalidHandle);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  std::size_t count = 0;
  fm_status_t rc = fm_sheet_cf_count(handle_, sheet, &count);
  if (rc != 0) {
    return FinishListResult(env, arr, rc);
  }
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < count; ++i) {
    fm_cf_rule_t rule{};
    rc = fm_sheet_cf_get_at(handle_, sheet, i, &rule);
    if (rc != 0) {
      return FinishListResult(env, arr, rc);
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("id", Napi::String::New(env, rule.id != nullptr ? rule.id : ""));
    item.Set("type", Napi::Number::New(env, static_cast<uint32_t>(rule.type)));
    item.Set("priority", Napi::Number::New(env, rule.priority));
    item.Set("stopIfTrue", Napi::Boolean::New(env, rule.stop_if_true != 0));
    Napi::Array sqref = Napi::Array::New(env, rule.sqref_count);
    for (uint32_t r = 0; r < rule.sqref_count; ++r) {
      Napi::Object rng = Napi::Object::New(env);
      rng.Set("firstRow", Napi::Number::New(env, rule.sqref[r].first_row));
      rng.Set("firstCol", Napi::Number::New(env, rule.sqref[r].first_col));
      rng.Set("lastRow", Napi::Number::New(env, rule.sqref[r].last_row));
      rng.Set("lastCol", Napi::Number::New(env, rule.sqref[r].last_col));
      sqref.Set(r, rng);
    }
    item.Set("sqref", sqref);
    if (rule.dxf_id_engaged != 0) {
      item.Set("dxfId", Napi::Number::New(env, rule.dxf_id));
    }
    if (rule.formula1 != nullptr) {
      item.Set("formula1", Napi::String::New(env, rule.formula1));
    }
    if (rule.formula2 != nullptr) {
      item.Set("formula2", Napi::String::New(env, rule.formula2));
    }
    if (rule.op_engaged != 0) {
      item.Set("op", Napi::Number::New(env, static_cast<uint32_t>(rule.op)));
    }
    if (rule.rank_engaged != 0) {
      item.Set("rank", Napi::Number::New(env, rule.rank));
      item.Set("percent", Napi::Boolean::New(env, rule.percent != 0));
      item.Set("bottom", Napi::Boolean::New(env, rule.bottom != 0));
    }
    // aboveAverage flags are always engineered but only meaningful for
    // the AboveAverage rule type; surface them only there to mirror the
    // embind shape.
    if (rule.type == 6 /* AboveAverage */) {
      item.Set("aboveAverage", Napi::Boolean::New(env, rule.above_average != 0));
      item.Set("equalAverage", Napi::Boolean::New(env, rule.equal_average != 0));
      if (rule.std_dev_engaged != 0) {
        item.Set("stdDev", Napi::Number::New(env, rule.std_dev));
      }
    }
    if (rule.text != nullptr) {
      item.Set("text", Napi::String::New(env, rule.text));
    }
    if (rule.time_period_engaged != 0) {
      item.Set("timePeriod", Napi::Number::New(env, static_cast<uint32_t>(rule.time_period)));
    }
    if (rule.color_scale_count > 0 && rule.color_scale_thresholds != nullptr && rule.color_scale_colors != nullptr) {
      Napi::Object color_scale = Napi::Object::New(env);
      Napi::Array thresholds = Napi::Array::New(env, rule.color_scale_count);
      Napi::Array colors = Napi::Array::New(env, rule.color_scale_count);
      for (uint32_t j = 0; j < rule.color_scale_count; ++j) {
        thresholds.Set(j, CfvoToJs(env, rule.color_scale_thresholds[j]));
        colors.Set(j, TranslateCfColor(env, rule.color_scale_colors[j]));
      }
      color_scale.Set("thresholds", thresholds);
      color_scale.Set("colors", colors);
      item.Set("colorScale", color_scale);
    }
    if (rule.data_bar_engaged != 0) {
      Napi::Object data_bar = Napi::Object::New(env);
      data_bar.Set("min", CfvoToJs(env, rule.data_bar_min));
      data_bar.Set("max", CfvoToJs(env, rule.data_bar_max));
      data_bar.Set("fill", TranslateCfColor(env, rule.data_bar_fill));
      data_bar.Set("showValue", Napi::Boolean::New(env, rule.data_bar_show_value != 0));
      data_bar.Set("minLengthPct", Napi::Number::New(env, static_cast<uint32_t>(rule.data_bar_min_length_pct)));
      data_bar.Set("maxLengthPct", Napi::Number::New(env, static_cast<uint32_t>(rule.data_bar_max_length_pct)));
      // `x14` extension payload. Each key appears only when the C ABI
      // reports the matching `*_engaged` flag, so an absent key means the
      // rule keeps the model default and the object round-trips through
      // `addConditionalFormat` unchanged. The getter engages all six.
      if (rule.data_bar_gradient_engaged != 0) {
        data_bar.Set("gradient", Napi::Boolean::New(env, rule.data_bar_gradient != 0));
      }
      if (rule.data_bar_axis_position_engaged != 0) {
        data_bar.Set("axisPosition", Napi::Number::New(env, static_cast<uint32_t>(rule.data_bar_axis_position)));
      }
      if (rule.data_bar_negative_fill_engaged != 0) {
        data_bar.Set("negativeFill", TranslateCfColor(env, rule.data_bar_negative_fill));
      }
      if (rule.data_bar_border_engaged != 0) {
        data_bar.Set("border", TranslateCfColor(env, rule.data_bar_border));
      }
      if (rule.data_bar_negative_border_engaged != 0) {
        data_bar.Set("negativeBorder", TranslateCfColor(env, rule.data_bar_negative_border));
      }
      if (rule.data_bar_axis_color_engaged != 0) {
        data_bar.Set("axisColor", TranslateCfColor(env, rule.data_bar_axis_color));
      }
      data_bar.Set("direction", Napi::Number::New(env, static_cast<uint32_t>(rule.data_bar_direction)));
      item.Set("dataBar", data_bar);
    }
    if (rule.icon_set_engaged != 0) {
      Napi::Object icon_set = Napi::Object::New(env);
      icon_set.Set("name", Napi::Number::New(env, static_cast<uint32_t>(rule.icon_set_name)));
      Napi::Array thresholds = Napi::Array::New(env, rule.icon_set_threshold_count);
      for (uint32_t j = 0; j < rule.icon_set_threshold_count; ++j) {
        thresholds.Set(j, CfvoToJs(env, rule.icon_set_thresholds[j]));
      }
      icon_set.Set("thresholds", thresholds);
      icon_set.Set("reverse", Napi::Boolean::New(env, rule.icon_set_reverse != 0));
      icon_set.Set("showValue", Napi::Boolean::New(env, rule.icon_set_show_value != 0));
      icon_set.Set("percent", Napi::Boolean::New(env, rule.icon_set_percent != 0));
      if (rule.icon_set_floor_engaged != 0) {
        icon_set.Set("floor", CfvoToJs(env, rule.icon_set_floor));
      }
      item.Set("iconSet", icon_set);
    }
    arr.Set(static_cast<uint32_t>(emitted), item);
    ++emitted;
  }
  return FinishListResult(env, arr, 0);
}

Napi::Value Workbook::AddConditionalFormat(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "index", 0);
  }
  if (info.Length() < 2 || !info[0].IsNumber() || !info[1].IsObject()) {
    Napi::TypeError::New(env, "addConditionalFormat expects (sheet:number, rule:object)").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  Napi::Object v = info[1].As<Napi::Object>();

  // Pull every JS field into local storage first; the C ABI receives
  // borrowed `const char*` views that must stay valid until
  // `fm_sheet_cf_add_rule` returns. Threshold value strings go into an
  // arena rather than a vector: the colorScale, dataBar and iconSet
  // branches below all feed the same store, and a vector would relocate
  // the earlier strings as the later branches add theirs.
  std::vector<fm_cf_cell_range_t> ranges_buf;
  std::vector<fm_cfvo_t> color_scale_thresholds;
  std::vector<fm_cf_color_t> color_scale_colors;
  std::vector<fm_cfvo_t> icon_set_thresholds;
  formulon::c_api::BorrowedStringArena cfvo_strings;
  CheckedSpecReader reader(env);
  Napi::Array sqref_arr;
  if (reader.Array(v, "sqref", &sqref_arr)) {
    const uint32_t n = sqref_arr.Length();
    if (!reader.ok()) {
      return env.Undefined();
    }
    ranges_buf.reserve(n);
    for (uint32_t i = 0; i < n; ++i) {
      if (!reader.ok()) {
        break;
      }
      Napi::Value rng_v;
      if (!reader.ArrayElement(sqref_arr, i, &rng_v)) {
        if (!reader.ok()) {
          break;
        }
        continue;
      }
      if (!rng_v.IsObject()) {
        continue;
      }
      Napi::Object rng = rng_v.As<Napi::Object>();
      fm_cf_cell_range_t r{};
      r.first_row = reader.U32(rng, "firstRow", 0U);
      r.first_col = reader.U32(rng, "firstCol", 0U);
      r.last_row = reader.U32(rng, "lastRow", 0U);
      r.last_col = reader.U32(rng, "lastCol", 0U);
      ranges_buf.push_back(r);
    }
  }
  std::string id;
  std::string formula1;
  std::string formula2;
  std::string text;
  reader.String(v, "id", &id);
  reader.String(v, "formula1", &formula1);
  reader.String(v, "formula2", &formula2);
  reader.String(v, "text", &text);

  fm_cf_rule_t rule{};
  rule.id = id.empty() ? nullptr : id.c_str();
  rule.type = reader.U8(v, "type", 0U);
  rule.priority = reader.I32(v, "priority", 0);
  rule.stop_if_true = reader.Bool(v, "stopIfTrue", false) ? 1 : 0;
  bool dxf_id_present = false;
  rule.dxf_id = reader.U32(v, "dxfId", 0U, &dxf_id_present);
  if (dxf_id_present) {
    rule.dxf_id_engaged = 1;
  }
  rule.sqref = ranges_buf.empty() ? nullptr : ranges_buf.data();
  rule.sqref_count = static_cast<uint32_t>(ranges_buf.size());
  rule.formula1 = formula1.empty() ? nullptr : formula1.c_str();
  rule.formula2 = formula2.empty() ? nullptr : formula2.c_str();
  bool op_present = false;
  rule.op = reader.U8(v, "op", 0U, &op_present);
  if (op_present) {
    rule.op_engaged = 1;
  }
  bool rank_present = false;
  rule.rank = reader.I32(v, "rank", 0, &rank_present);
  if (rank_present) {
    rule.rank_engaged = 1;
  }
  rule.percent = reader.Bool(v, "percent", false) ? 1 : 0;
  rule.bottom = reader.Bool(v, "bottom", false) ? 1 : 0;
  rule.above_average = reader.Bool(v, "aboveAverage", true) ? 1 : 0;
  rule.equal_average = reader.Bool(v, "equalAverage", false) ? 1 : 0;
  bool std_dev_present = false;
  rule.std_dev = reader.Double(v, "stdDev", 0.0, &std_dev_present);
  if (std_dev_present) {
    rule.std_dev_engaged = 1;
  }
  rule.text = text.empty() ? nullptr : text.c_str();
  bool time_period_present = false;
  rule.time_period = reader.U8(v, "timePeriod", 0U, &time_period_present);
  if (time_period_present) {
    rule.time_period_engaged = 1;
  }
  Napi::Object cs;
  if (reader.Object(v, "colorScale", &cs)) {
    Napi::Array arr;
    if (reader.Array(cs, "thresholds", &arr)) {
      color_scale_thresholds.reserve(arr.Length());
      for (uint32_t i = 0; i < arr.Length(); ++i) {
        if (!reader.ok()) {
          break;
        }
        Napi::Value threshold;
        if (!reader.ArrayElement(arr, i, &threshold)) {
          if (!reader.ok()) {
            break;
          }
          continue;
        }
        if (threshold.IsObject()) {
          color_scale_thresholds.push_back(PullCfvo(reader, threshold.As<Napi::Object>(), &cfvo_strings));
        }
      }
    }
    if (reader.Array(cs, "colors", &arr)) {
      color_scale_colors.reserve(arr.Length());
      for (uint32_t i = 0; i < arr.Length(); ++i) {
        if (!reader.ok()) {
          break;
        }
        Napi::Value color;
        if (!reader.ArrayElement(arr, i, &color)) {
          if (!reader.ok()) {
            break;
          }
          continue;
        }
        if (color.IsObject()) {
          color_scale_colors.push_back(PullCfColor(reader, color.As<Napi::Object>()));
        }
      }
    }
    rule.color_scale_thresholds = color_scale_thresholds.empty() ? nullptr : color_scale_thresholds.data();
    rule.color_scale_colors = color_scale_colors.empty() ? nullptr : color_scale_colors.data();
    rule.color_scale_count = static_cast<uint32_t>(color_scale_thresholds.size());
  }
  Napi::Object db;
  if (reader.Object(v, "dataBar", &db)) {
    rule.data_bar_engaged = 1;
    Napi::Object min;
    if (reader.Object(db, "min", &min)) {
      rule.data_bar_min = PullCfvo(reader, min, &cfvo_strings);
    }
    Napi::Object max;
    if (reader.Object(db, "max", &max)) {
      rule.data_bar_max = PullCfvo(reader, max, &cfvo_strings);
    }
    Napi::Object fill;
    if (reader.Object(db, "fill", &fill)) {
      rule.data_bar_fill = PullCfColor(reader, fill);
    }
    rule.data_bar_show_value = reader.Bool(db, "showValue", true) ? 1 : 0;
    rule.data_bar_min_length_pct = reader.U8(db, "minLengthPct", 10U);
    rule.data_bar_max_length_pct = reader.U8(db, "maxLengthPct", 90U);
    // `x14` extension payload. An omitted key leaves the `*_engaged` flag
    // clear, which the C ABI reads as "keep the model default" (gradient
    // on, automatic axis, negative fill equal to the positive fill, no
    // border, black axis).
    bool gradient_present = false;
    const bool gradient = reader.Bool(db, "gradient", true, &gradient_present);
    if (gradient_present) {
      rule.data_bar_gradient_engaged = 1;
      rule.data_bar_gradient = gradient ? 1 : 0;
    }
    bool axis_position_present = false;
    rule.data_bar_axis_position = reader.U8(db, "axisPosition", 0U, &axis_position_present);
    if (axis_position_present) {
      rule.data_bar_axis_position_engaged = 1;
    }
    Napi::Object negative_fill;
    if (reader.Object(db, "negativeFill", &negative_fill)) {
      rule.data_bar_negative_fill_engaged = 1;
      rule.data_bar_negative_fill = PullCfColor(reader, negative_fill);
    }
    Napi::Object border;
    if (reader.Object(db, "border", &border)) {
      rule.data_bar_border_engaged = 1;
      rule.data_bar_border = PullCfColor(reader, border);
    }
    Napi::Object negative_border;
    if (reader.Object(db, "negativeBorder", &negative_border)) {
      rule.data_bar_negative_border_engaged = 1;
      rule.data_bar_negative_border = PullCfColor(reader, negative_border);
    }
    Napi::Object axis_color;
    if (reader.Object(db, "axisColor", &axis_color)) {
      rule.data_bar_axis_color_engaged = 1;
      rule.data_bar_axis_color = PullCfColor(reader, axis_color);
    }
    rule.data_bar_direction = reader.U8(db, "direction", 0U);
  }
  Napi::Object is;
  if (reader.Object(v, "iconSet", &is)) {
    rule.icon_set_engaged = 1;
    rule.icon_set_name = reader.U8(is, "name", 0U);
    Napi::Array arr;
    if (reader.Array(is, "thresholds", &arr)) {
      icon_set_thresholds.reserve(arr.Length());
      for (uint32_t i = 0; i < arr.Length(); ++i) {
        if (!reader.ok()) {
          break;
        }
        Napi::Value threshold;
        if (!reader.ArrayElement(arr, i, &threshold)) {
          if (!reader.ok()) {
            break;
          }
          continue;
        }
        if (threshold.IsObject()) {
          icon_set_thresholds.push_back(PullCfvo(reader, threshold.As<Napi::Object>(), &cfvo_strings));
        }
      }
    }
    rule.icon_set_thresholds = icon_set_thresholds.empty() ? nullptr : icon_set_thresholds.data();
    rule.icon_set_threshold_count = static_cast<uint32_t>(icon_set_thresholds.size());
    rule.icon_set_reverse = reader.Bool(is, "reverse", false) ? 1 : 0;
    rule.icon_set_show_value = reader.Bool(is, "showValue", true) ? 1 : 0;
    rule.icon_set_percent = reader.Bool(is, "percent", true) ? 1 : 0;
    // An omitted floor keeps Excel's default, `percent 0`.
    Napi::Object floor;
    if (reader.Object(is, "floor", &floor)) {
      rule.icon_set_floor_engaged = 1;
      rule.icon_set_floor = PullCfvo(reader, floor, &cfvo_strings);
    }
  }
  if (!reader.ok()) {
    // See the matching guard in AddFont. The return value is discarded
    // in favor of the pending exception either way.
    return env.Undefined();
  }
  std::size_t new_index = 0;
  fm_status_t rc = fm_sheet_cf_add_rule(handle_, sheet, rule, &new_index);
  return MakeNumberFieldResult(env, rc, "index", static_cast<double>(new_index));
}

Napi::Value Workbook::RemoveConditionalFormatAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t index = static_cast<std::size_t>(ArgU32(info, 1));
  fm_status_t rc = fm_sheet_cf_remove_at(handle_, sheet, index);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::ClearConditionalFormats(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  fm_status_t rc = fm_sheet_cf_clear(handle_, sheet);
  return MakeStatus(env, rc);
}
}  // namespace formulon_node
