// Copyright 2026 libraz. Licensed under the MIT License.
//
// Style-table bindings: read / write entries in the xf / font / fill /
// border / numFmt pools plus the per-cell xf-index getters, and the
// `evaluateCfRange` conditional-formatting evaluator. The cf
// translator lives here because it shares the styles compilation
// dependency on `fm_styles_*` typedefs declared in the same C ABI
// header block.

#include <cstddef>
#include <cstdint>
#include <string>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

// ---- Cell-XF index --------------------------------------------------

Napi::Value Workbook::GetCellXfIndex(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "xfIndex", 0);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  uint32_t xf = 0;
  fm_status_t rc = fm_cell_get_xf_index(handle_, sheet, row, col, &xf);
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "xfIndex", 0);
  }
  return MakeNumberFieldResult(env, MakeOkStatus(env), "xfIndex", xf);
}

Napi::Value Workbook::SetCellXfIndex(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  const uint32_t xf = ArgU32(info, 3);
  fm_status_t rc = fm_cell_set_xf_index(handle_, sheet, row, col, xf);
  return MakeStatus(env, rc);
}

// ---- Style getters --------------------------------------------------

Napi::Value Workbook::GetCellXf(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint32_t xf_index = ArgU32(info, 0);
  fm_cell_xf xf{};
  fm_status_t rc = fm_styles_get_cell_xf(handle_, xf_index, &xf);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("fontIndex", Napi::Number::New(env, xf.font_index));
  out.Set("fillIndex", Napi::Number::New(env, xf.fill_index));
  out.Set("borderIndex", Napi::Number::New(env, xf.border_index));
  out.Set("numFmtId", Napi::Number::New(env, static_cast<uint32_t>(xf.num_fmt_id)));
  out.Set("horizontalAlign", Napi::Number::New(env, static_cast<uint32_t>(xf.horizontal_align)));
  out.Set("verticalAlign", Napi::Number::New(env, static_cast<uint32_t>(xf.vertical_align)));
  out.Set("wrapText", Napi::Boolean::New(env, xf.wrap_text != 0));
  return out;
}

Napi::Value Workbook::GetFont(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint32_t font_index = ArgU32(info, 0);
  fm_font_record f{};
  fm_status_t rc = fm_styles_get_font(handle_, font_index, &f);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("name", Napi::String::New(env, f.name != nullptr ? f.name : ""));
  out.Set("size", Napi::Number::New(env, f.size));
  out.Set("colorArgb", Napi::Number::New(env, f.color_argb));
  out.Set("bold", Napi::Boolean::New(env, f.bold != 0));
  out.Set("italic", Napi::Boolean::New(env, f.italic != 0));
  out.Set("strike", Napi::Boolean::New(env, f.strike != 0));
  out.Set("underline", Napi::Number::New(env, static_cast<uint32_t>(f.underline)));
  return out;
}

Napi::Value Workbook::GetFill(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint32_t fill_index = ArgU32(info, 0);
  fm_fill_record f{};
  fm_status_t rc = fm_styles_get_fill(handle_, fill_index, &f);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("pattern", Napi::Number::New(env, static_cast<uint32_t>(f.pattern)));
  out.Set("fgArgb", Napi::Number::New(env, f.fg_argb));
  out.Set("bgArgb", Napi::Number::New(env, f.bg_argb));
  return out;
}

Napi::Value Workbook::GetBorder(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint32_t border_index = ArgU32(info, 0);
  fm_border_record b{};
  fm_status_t rc = fm_styles_get_border(handle_, border_index, &b);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  auto side_obj = [&](const fm_border_side& s) {
    Napi::Object so = Napi::Object::New(env);
    so.Set("style", Napi::Number::New(env, static_cast<uint32_t>(s.style)));
    so.Set("colorArgb", Napi::Number::New(env, s.color_argb));
    return so;
  };
  out.Set("status", MakeOkStatus(env));
  out.Set("left", side_obj(b.left));
  out.Set("right", side_obj(b.right));
  out.Set("top", side_obj(b.top));
  out.Set("bottom", side_obj(b.bottom));
  out.Set("diagonal", side_obj(b.diagonal));
  out.Set("diagonalUp", Napi::Boolean::New(env, b.diagonal_up != 0));
  out.Set("diagonalDown", Napi::Boolean::New(env, b.diagonal_down != 0));
  return out;
}

Napi::Value Workbook::GetNumFmt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint32_t num_fmt_id = ArgU32(info, 0);
  const char* s = nullptr;
  fm_status_t rc = fm_styles_get_num_fmt_string(handle_, static_cast<uint16_t>(num_fmt_id), &s);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("numFmtId", Napi::Number::New(env, num_fmt_id));
  out.Set("formatCode", Napi::String::New(env, s != nullptr ? s : ""));
  return out;
}

// ---- Style adders ---------------------------------------------------

Napi::Value Workbook::AddFont(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "index", 0);
  }
  Napi::Object record = (info.Length() > 0 && info[0].IsObject()) ? info[0].As<Napi::Object>() : Napi::Object::New(env);
  std::string name;
  if (record.Has("name") && !record.Get("name").IsUndefined() && !record.Get("name").IsNull()) {
    name = record.Get("name").ToString().Utf8Value();
  }
  fm_font_record fr{};
  fr.name = name.c_str();
  fr.size = record.Has("size") ? record.Get("size").ToNumber().DoubleValue() : 11.0;
  fr.bold = (record.Has("bold") && record.Get("bold").ToBoolean().Value()) ? 1 : 0;
  fr.italic = (record.Has("italic") && record.Get("italic").ToBoolean().Value()) ? 1 : 0;
  fr.strike = (record.Has("strike") && record.Get("strike").ToBoolean().Value()) ? 1 : 0;
  fr.underline = record.Has("underline") ? static_cast<uint8_t>(record.Get("underline").ToNumber().Uint32Value()) : 0U;
  fr.color_argb = record.Has("colorArgb") ? record.Get("colorArgb").ToNumber().Uint32Value() : 0xFF000000U;
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_font(handle_, fr, &idx);
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "index", 0);
  }
  return MakeNumberFieldResult(env, MakeOkStatus(env), "index", idx);
}

Napi::Value Workbook::AddFill(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "index", 0);
  }
  Napi::Object record = (info.Length() > 0 && info[0].IsObject()) ? info[0].As<Napi::Object>() : Napi::Object::New(env);
  fm_fill_record fr{};
  fr.pattern = record.Has("pattern") ? static_cast<uint8_t>(record.Get("pattern").ToNumber().Uint32Value()) : 0U;
  fr.fg_argb = record.Has("fgArgb") ? record.Get("fgArgb").ToNumber().Uint32Value() : 0U;
  fr.bg_argb = record.Has("bgArgb") ? record.Get("bgArgb").ToNumber().Uint32Value() : 0U;
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_fill(handle_, fr, &idx);
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "index", 0);
  }
  return MakeNumberFieldResult(env, MakeOkStatus(env), "index", idx);
}

Napi::Value Workbook::AddBorder(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "index", 0);
  }
  Napi::Object record = (info.Length() > 0 && info[0].IsObject()) ? info[0].As<Napi::Object>() : Napi::Object::New(env);
  auto pull_side = [&](const char* key) {
    fm_border_side s{};
    if (record.Has(key) && record.Get(key).IsObject()) {
      Napi::Object so = record.Get(key).As<Napi::Object>();
      if (so.Has("style")) {
        s.style = static_cast<uint8_t>(so.Get("style").ToNumber().Uint32Value());
      }
      if (so.Has("colorArgb")) {
        s.color_argb = so.Get("colorArgb").ToNumber().Uint32Value();
      }
    }
    return s;
  };
  fm_border_record br{};
  br.left = pull_side("left");
  br.right = pull_side("right");
  br.top = pull_side("top");
  br.bottom = pull_side("bottom");
  br.diagonal = pull_side("diagonal");
  br.diagonal_up = (record.Has("diagonalUp") && record.Get("diagonalUp").ToBoolean().Value()) ? 1 : 0;
  br.diagonal_down = (record.Has("diagonalDown") && record.Get("diagonalDown").ToBoolean().Value()) ? 1 : 0;
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_border(handle_, br, &idx);
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "index", 0);
  }
  return MakeNumberFieldResult(env, MakeOkStatus(env), "index", idx);
}

Napi::Value Workbook::AddNumFmt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "numFmtId", 0);
  }
  const std::string code = ArgString(info, 0);
  uint16_t id = 0;
  fm_status_t rc = fm_styles_add_num_fmt(handle_, code.c_str(), &id);
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "numFmtId", 0);
  }
  return MakeNumberFieldResult(env, MakeOkStatus(env), "numFmtId", static_cast<uint32_t>(id));
}

Napi::Value Workbook::AddXf(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "index", 0);
  }
  Napi::Object record = (info.Length() > 0 && info[0].IsObject()) ? info[0].As<Napi::Object>() : Napi::Object::New(env);
  fm_cell_xf xf{};
  xf.font_index = record.Has("fontIndex") ? record.Get("fontIndex").ToNumber().Uint32Value() : 0U;
  xf.fill_index = record.Has("fillIndex") ? record.Get("fillIndex").ToNumber().Uint32Value() : 0U;
  xf.border_index = record.Has("borderIndex") ? record.Get("borderIndex").ToNumber().Uint32Value() : 0U;
  xf.num_fmt_id = record.Has("numFmtId") ? static_cast<uint16_t>(record.Get("numFmtId").ToNumber().Uint32Value()) : 0U;
  xf.horizontal_align =
      record.Has("horizontalAlign") ? static_cast<uint8_t>(record.Get("horizontalAlign").ToNumber().Uint32Value()) : 0U;
  xf.vertical_align =
      record.Has("verticalAlign") ? static_cast<uint8_t>(record.Get("verticalAlign").ToNumber().Uint32Value()) : 0U;
  xf.wrap_text = (record.Has("wrapText") && record.Get("wrapText").ToBoolean().Value()) ? 1 : 0;
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_cell_xf(handle_, xf, &idx);
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "index", 0);
  }
  return MakeNumberFieldResult(env, MakeOkStatus(env), "index", idx);
}

// ---- Style pool counts ----------------------------------------------
//
// `FontCount` / `FillCount` / `BorderCount` / `XfCount` are now emitted
// by the binding codegen (see `src/node_addon/generated/styles_counts.cc`).

// ---- Conditional formatting -----------------------------------------

Napi::Value Workbook::EvaluateCfRange(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    out.Set("cells", Napi::Array::New(env));
    return out;
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t first_row = ArgU32(info, 1);
  const uint32_t first_col = ArgU32(info, 2);
  const uint32_t last_row = ArgU32(info, 3);
  const uint32_t last_col = ArgU32(info, 4);
  const double today_serial = ArgDouble(info, 5);
  fm_cf_results_t* results = nullptr;
  fm_status_t rc =
      fm_workbook_cf_evaluate_range(handle_, sheet, first_row, first_col, last_row, last_col, today_serial, &results);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    out.Set("cells", Napi::Array::New(env));
    return out;
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
  out.Set("status", MakeOkStatus(env));
  out.Set("cells", cells);
  return out;
}

}  // namespace formulon_node
