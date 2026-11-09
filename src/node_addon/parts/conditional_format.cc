// Conditional-formatting bindings: the `evaluateCfRange` evaluator plus the
// read / mutate surface for per-sheet conditional-format rules.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/borrowed_arena.h"
#include "node_addon/parts/workbook_class.h"

namespace formulon_node {
namespace {

fm_cfvo_t PullCfvo(CheckedSpecReader& reader, const Napi::Object& spec, formulon::c_api::BorrowedStringArena* strings) {
  fm_cfvo_t out{};
  out.type = reader.I32(spec, "type", 0);
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
  out.Set("value", cfvo.value != nullptr ? Napi::Value(Napi::String::New(env, cfvo.value)) : env.Null());
  return out;
}

}  // namespace

// ---- Conditional formatting -----------------------------------------

Napi::Value Workbook::EvaluateCfRange(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  CheckedSpecReader reader(env);
  fm_cell_rect_t rect{};
  if (!ReadCellRect(reader, ArgObject(info, 1), &rect)) {
    return env.Undefined();
  }
  // NaN is "no today": `timePeriod` rules then match nothing.
  const double today_serial = ArgDouble(info, 2);
  const int32_t has_today_serial = std::isnan(today_serial) ? 0 : 1;
  fm_cf_results_t* results = nullptr;
  fm_status_t rc = fm_sheet_evaluate_cf_range(handle_, sheet, &rect, has_today_serial,
                                              has_today_serial != 0 ? today_serial : 0.0, &results);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  std::size_t cell_count = 0;
  rc = fm_cf_results_get_cell_count(results, &cell_count);
  if (rc != 0) {
    fm_cf_results_destroy(results);
    return ThrowStatus(env, rc);
  }
  Napi::Array cells = Napi::Array::New(env, cell_count);
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < cell_count; ++i) {
    uint32_t row = 0;
    uint32_t col = 0;
    std::size_t match_count = 0;
    if (fm_cf_results_get_cell_at(results, i, &row, &col, &match_count) != 0) {
      // Defensive: skip entries the C ABI declines to materialise.
      continue;
    }
    Napi::Array matches = Napi::Array::New(env, match_count);
    std::size_t mj = 0;
    for (std::size_t j = 0; j < match_count; ++j) {
      fm_cf_match_info_t m = FM_STRUCT_INIT(fm_cf_match_info_t);
      if (fm_cf_results_get_match_at(results, i, j, &m) != 0) {
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
  return cells;
}

// ---- Conditional formatting (read / mutate) -------------------------

namespace {

// Builds the `ConditionalFormatEntry` shape of `rule`.
Napi::Object CfRuleToJs(Napi::Env env, const fm_cf_rule_t& rule) {
  Napi::Object item = Napi::Object::New(env);
  item.Set("id", JsString(env, rule.id));
  item.Set("type", Napi::Number::New(env, static_cast<uint32_t>(rule.type)));
  item.Set("priority", Napi::Number::New(env, rule.priority));
  item.Set("stopIfTrue", Napi::Boolean::New(env, rule.stop_if_true != 0));
  Napi::Array ranges = Napi::Array::New(env, rule.range_count);
  for (size_t r = 0; r < rule.range_count; ++r) {
    ranges.Set(static_cast<uint32_t>(r), RangeToJs(env, rule.ranges[r]));
  }
  item.Set("ranges", ranges);
  item.Set("dxfIndex", JsNumberOrNull(env, rule.has_dxf_index, rule.dxf_index));
  item.Set("formula1", JsStringOrNull(env, rule.formula1));
  item.Set("formula2", JsStringOrNull(env, rule.formula2));
  item.Set("operator", JsNumberOrNull(env, rule.has_op, static_cast<uint32_t>(rule.op)));
  item.Set("rank", JsNumberOrNull(env, rule.has_rank, rule.rank));
  item.Set("percent", JsBoolOrNull(env, rule.has_rank, rule.percent));
  item.Set("bottom", JsBoolOrNull(env, rule.has_rank, rule.bottom));
  // The average flags are meaningful only for the AboveAverage rule type.
  const bool above_average_rule = rule.type == 6 /* AboveAverage */;
  item.Set("aboveAverage", JsBoolOrNull(env, above_average_rule ? 1 : 0, rule.above_average));
  item.Set("equalAverage", JsBoolOrNull(env, above_average_rule ? 1 : 0, rule.equal_average));
  item.Set("stdDev", JsNumberOrNull(env, above_average_rule ? rule.has_std_dev : 0, rule.std_dev));
  item.Set("text", JsStringOrNull(env, rule.text));
  item.Set("timePeriod", JsNumberOrNull(env, rule.has_time_period, static_cast<uint32_t>(rule.time_period)));
  item.Set("colorScale", env.Null());
  if (rule.color_scale_count > 0 && rule.color_scale_thresholds != nullptr && rule.color_scale_colors != nullptr) {
    Napi::Object color_scale = Napi::Object::New(env);
    Napi::Array thresholds = Napi::Array::New(env, rule.color_scale_count);
    Napi::Array colors = Napi::Array::New(env, rule.color_scale_count);
    for (uint32_t j = 0; j < rule.color_scale_count; ++j) {
      thresholds.Set(j, CfvoToJs(env, rule.color_scale_thresholds[j]));
      colors.Set(j, ColorSpecToJs(env, rule.color_scale_colors[j]));
    }
    color_scale.Set("thresholds", thresholds);
    color_scale.Set("colors", colors);
    item.Set("colorScale", color_scale);
  }
  item.Set("dataBar", env.Null());
  if (rule.has_data_bar != 0) {
    Napi::Object data_bar = Napi::Object::New(env);
    data_bar.Set("min", CfvoToJs(env, rule.data_bar_min));
    data_bar.Set("max", CfvoToJs(env, rule.data_bar_max));
    data_bar.Set("fill", ColorSpecToJs(env, rule.data_bar_fill));
    data_bar.Set("showValue", Napi::Boolean::New(env, rule.data_bar_show_value != 0));
    data_bar.Set("minLengthPct", Napi::Number::New(env, static_cast<uint32_t>(rule.data_bar_min_length_pct)));
    data_bar.Set("maxLengthPct", Napi::Number::New(env, static_cast<uint32_t>(rule.data_bar_max_length_pct)));
    // The `x14` extension keys read `null` when the rule keeps the model default.
    data_bar.Set("gradient", JsBoolOrNull(env, rule.has_data_bar_gradient, rule.data_bar_gradient));
    data_bar.Set("axisPosition", JsNumberOrNull(env, rule.has_data_bar_axis_position,
                                                static_cast<uint32_t>(rule.data_bar_axis_position)));
    const auto color_or_null = [&env](const fm_color_t& color) {
      return color.kind != FM_COLOR_KIND_NONE ? Napi::Value(ColorSpecToJs(env, color)) : env.Null();
    };
    data_bar.Set("negativeFill", color_or_null(rule.data_bar_negative_fill));
    data_bar.Set("border", color_or_null(rule.data_bar_border));
    data_bar.Set("negativeBorder", color_or_null(rule.data_bar_negative_border));
    data_bar.Set("axisColor", color_or_null(rule.data_bar_axis_color));
    data_bar.Set("direction", Napi::Number::New(env, static_cast<uint32_t>(rule.data_bar_direction)));
    item.Set("dataBar", data_bar);
  }
  item.Set("iconSet", env.Null());
  if (rule.has_icon_set != 0) {
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
    icon_set.Set("floor", rule.has_icon_set_floor != 0 ? Napi::Value(CfvoToJs(env, rule.icon_set_floor)) : env.Null());
    icon_set.Set("icons", env.Null());
    if (rule.icon_set_icon_count != 0U) {
      Napi::Array icons = Napi::Array::New(env, rule.icon_set_icon_count);
      for (uint32_t j = 0; j < rule.icon_set_icon_count; ++j) {
        Napi::Object icon = Napi::Object::New(env);
        icon.Set("iconSet", Napi::Number::New(env, rule.icon_set_icons[j].icon_set));
        icon.Set("iconIndex", Napi::Number::New(env, rule.icon_set_icons[j].icon_index));
        icons.Set(j, icon);
      }
      icon_set.Set("icons", icons);
    }
    item.Set("iconSet", icon_set);
  }
  return item;
}

}  // namespace

Napi::Value Workbook::GetCfRuleCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  std::size_t count = 0;
  const fm_status_t rc = fm_sheet_get_cf_rule_count(handle_, ArgU32(info, 0), &count);
  return ValueOrThrow(env, rc, Napi::Number::New(env, static_cast<double>(count)));
}

Napi::Value Workbook::GetCfRule(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  const fm_status_t rc = fm_sheet_get_cf_rule_at(handle_, ArgU32(info, 0), ArgU32(info, 1), &rule);
  return ValueOrThrow(env, rc, rc == 0 ? Napi::Value(CfRuleToJs(env, rule)) : env.Undefined());
}

Napi::Value Workbook::AddConditionalFormat(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  Napi::Object v = ArgObject(info, 1);

  // Pull every JS field into local storage first; the C ABI receives
  // borrowed `const char*` views that must stay valid until
  // `fm_sheet_add_cf_rule` returns. Threshold value strings go into an
  // arena rather than a vector: the colorScale, dataBar and iconSet
  // branches below all feed the same store, and a vector would relocate
  // the earlier strings as the later branches add theirs.
  std::vector<fm_cell_rect_t> ranges_buf;
  std::vector<fm_cfvo_t> color_scale_thresholds;
  std::vector<fm_color_t> color_scale_colors;
  std::vector<fm_cfvo_t> icon_set_thresholds;
  std::vector<fm_cf_icon_t> icon_set_icons;
  formulon::c_api::BorrowedStringArena cfvo_strings;
  CheckedSpecReader reader(env);
  Napi::Array ranges_arr;
  if (reader.RequiredArray(v, "ranges", &ranges_arr)) {
    const uint32_t n = ranges_arr.Length();
    if (!reader.ok()) {
      return env.Undefined();
    }
    ranges_buf.reserve(n);
    for (uint32_t i = 0; i < n; ++i) {
      if (!reader.ok()) {
        break;
      }
      Napi::Value rng_v;
      if (!reader.ArrayElement(ranges_arr, i, &rng_v)) {
        if (!reader.ok()) {
          break;
        }
        continue;
      }
      if (!rng_v.IsObject()) {
        continue;
      }
      Napi::Object rng = rng_v.As<Napi::Object>();
      fm_cell_rect_t r{};
      if (!ReadCellRect(reader, rng, &r)) {
        break;
      }
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

  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.id = id.empty() ? nullptr : id.c_str();
  rule.type = reader.I32(v, "type", 0);
  rule.priority = reader.I32(v, "priority", 0);
  rule.stop_if_true = reader.Bool(v, "stopIfTrue", false) ? 1 : 0;
  bool dxf_index_present = false;
  rule.dxf_index = reader.U32(v, "dxfIndex", 0U, &dxf_index_present);
  if (dxf_index_present) {
    rule.has_dxf_index = 1;
  }
  rule.ranges = ranges_buf.empty() ? nullptr : ranges_buf.data();
  rule.range_count = ranges_buf.size();
  rule.formula1 = formula1.empty() ? nullptr : formula1.c_str();
  rule.formula2 = formula2.empty() ? nullptr : formula2.c_str();
  bool op_present = false;
  rule.op = reader.I32(v, "operator", 0, &op_present);
  if (op_present) {
    rule.has_op = 1;
  }
  bool rank_present = false;
  rule.rank = reader.I32(v, "rank", 0, &rank_present);
  if (rank_present) {
    rule.has_rank = 1;
  }
  rule.percent = reader.Bool(v, "percent", false) ? 1 : 0;
  rule.bottom = reader.Bool(v, "bottom", false) ? 1 : 0;
  rule.above_average = reader.Bool(v, "aboveAverage", true) ? 1 : 0;
  rule.equal_average = reader.Bool(v, "equalAverage", false) ? 1 : 0;
  bool std_dev_present = false;
  rule.std_dev = reader.Double(v, "stdDev", 0.0, &std_dev_present);
  if (std_dev_present) {
    rule.has_std_dev = 1;
  }
  rule.text = text.empty() ? nullptr : text.c_str();
  bool time_period_present = false;
  rule.time_period = reader.I32(v, "timePeriod", 0, &time_period_present);
  if (time_period_present) {
    rule.has_time_period = 1;
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
          color_scale_colors.push_back(ReadColorSpec(reader, color, "colors"));
        }
      }
    }
    rule.color_scale_thresholds = color_scale_thresholds.empty() ? nullptr : color_scale_thresholds.data();
    rule.color_scale_colors = color_scale_colors.empty() ? nullptr : color_scale_colors.data();
    rule.color_scale_count = static_cast<uint32_t>(color_scale_thresholds.size());
  }
  Napi::Object db;
  if (reader.Object(v, "dataBar", &db)) {
    rule.has_data_bar = 1;
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
      rule.data_bar_fill = ReadColorSpec(reader, fill, "fill");
    }
    rule.data_bar_show_value = reader.Bool(db, "showValue", true) ? 1 : 0;
    rule.data_bar_min_length_pct = reader.U8(db, "minLengthPct", 10U);
    rule.data_bar_max_length_pct = reader.U8(db, "maxLengthPct", 90U);
    // `x14` extension payload. An omitted key leaves the `has_*` flag
    // clear, which the C ABI reads as "keep the model default" (gradient
    // on, automatic axis, negative fill equal to the positive fill, no
    // border, black axis).
    bool gradient_present = false;
    const bool gradient = reader.Bool(db, "gradient", true, &gradient_present);
    if (gradient_present) {
      rule.has_data_bar_gradient = 1;
      rule.data_bar_gradient = gradient ? 1 : 0;
    }
    bool axis_position_present = false;
    rule.data_bar_axis_position = reader.I32(db, "axisPosition", 0, &axis_position_present);
    if (axis_position_present) {
      rule.has_data_bar_axis_position = 1;
    }
    Napi::Object negative_fill;
    if (reader.Object(db, "negativeFill", &negative_fill)) {
      rule.data_bar_negative_fill = ReadColorSpec(reader, negative_fill, "negativeFill");
    }
    Napi::Object border;
    if (reader.Object(db, "border", &border)) {
      rule.data_bar_border = ReadColorSpec(reader, border, "border");
    }
    Napi::Object negative_border;
    if (reader.Object(db, "negativeBorder", &negative_border)) {
      rule.data_bar_negative_border = ReadColorSpec(reader, negative_border, "negativeBorder");
    }
    Napi::Object axis_color;
    if (reader.Object(db, "axisColor", &axis_color)) {
      rule.data_bar_axis_color = ReadColorSpec(reader, axis_color, "axisColor");
    }
    rule.data_bar_direction = reader.I32(db, "direction", 0);
  }
  Napi::Object is;
  if (reader.Object(v, "iconSet", &is)) {
    rule.has_icon_set = 1;
    rule.icon_set_name = reader.I32(is, "name", 0);
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
      rule.has_icon_set_floor = 1;
      rule.icon_set_floor = PullCfvo(reader, floor, &cfvo_strings);
    }
    Napi::Array icons;
    if (reader.Array(is, "icons", &icons)) {
      icon_set_icons.reserve(icons.Length());
      for (uint32_t i = 0; i < icons.Length(); ++i) {
        Napi::Value element;
        if (!reader.ArrayElement(icons, i, &element)) {
          if (!reader.ok()) {
            break;
          }
          continue;
        }
        Napi::Object icon;
        if (!reader.Object(element, "icons[]", &icon)) {
          break;
        }
        // A missing set stays out of range, which the C layer rejects.
        const int32_t icon_set = reader.I32(icon, "iconSet", -1);
        icon_set_icons.push_back(fm_cf_icon_t{icon_set, reader.I32(icon, "iconIndex", 0)});
      }
      rule.icon_set_icons = icon_set_icons.empty() ? nullptr : icon_set_icons.data();
      rule.icon_set_icon_count = static_cast<uint32_t>(icon_set_icons.size());
    }
  }
  if (!reader.ok()) {
    // See the matching guard in AddFont. The return value is discarded
    // in favor of the pending exception either way.
    return env.Undefined();
  }
  std::size_t new_index = 0;
  const fm_status_t rc = fm_sheet_add_cf_rule(handle_, sheet, &rule, &new_index);
  return ValueOrThrow(env, rc, Napi::Number::New(env, static_cast<double>(new_index)));
}

Napi::Value Workbook::RemoveConditionalFormatAt(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_cf_rule);
}

Napi::Value Workbook::ClearConditionalFormats(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  return ValueOrThrow(env, fm_sheet_clear_cf_rules(handle_, sheet), env.Undefined());
}
}  // namespace formulon_node
