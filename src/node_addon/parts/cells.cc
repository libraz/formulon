// Workbook cell bindings: cell mutation and read (including phonetic
// metadata), ad-hoc formula evaluation, and lambda text.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

// ---- Cell mutation --------------------------------------------------

Napi::Value Workbook::SetNumber(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  const double value = ArgDouble(info, 3);
  fm_status_t rc = fm_workbook_set_number(handle_, sheet, row, col, value);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetBool(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  bool value = false;
  if (info.Length() > 3) {
    value = info[3].ToBoolean().Value();
  }
  fm_status_t rc = fm_workbook_set_bool(handle_, sheet, row, col, value ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetError(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (info.Length() < 4) {
    // Unlike the other cell-mutation setters, a missing 4th argument
    // here has no sane zero-value default: `error_code = 0` silently
    // writes `#NULL!`, masking a caller bug instead of surfacing it.
    // Reject it the same way the WASM (embind arity check) and Python
    // (required positional parameter) bindings already do.
    Napi::TypeError::New(env, "setError requires 4 arguments (sheet, row, col, errorCode)")
        .ThrowAsJavaScriptException();
    return env.Undefined();
  }
  if (!info[3].IsNumber()) {
    // `.As<Napi::Number>()` below is an unchecked cast: on a non-number
    // 4th argument it does not itself throw, so without this check the
    // mutation below would run first and any resulting exception would
    // only surface afterwards, on return to JS.
    Napi::TypeError::New(env, "setError: `errorCode` must be a number").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  const int32_t error_code = info[3].As<Napi::Number>().Int32Value();
  fm_status_t rc = fm_workbook_set_error(handle_, sheet, row, col, static_cast<fm_error_code_t>(error_code));
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetText(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  const std::string text = ArgString(info, 3);
  fm_status_t rc = fm_workbook_set_text(handle_, sheet, row, col, text.c_str());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetCellPhonetic(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  const std::string phonetic = ArgString(info, 3);
  fm_status_t rc = fm_workbook_set_cell_phonetic(handle_, sheet, row, col, phonetic.c_str());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetCellPhoneticRuns(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  if (info.Length() <= 3 || !info[3].IsArray()) {
    return MakeBindingArgumentError(env, "setCellPhoneticRuns: `runs` must be an array of { sb, eb, text }");
  }
  const Napi::Array runs = info[3].As<Napi::Array>();
  const uint32_t count = runs.Length();
  // Two passes so no `c_str()` is taken before `texts` has finished growing.
  std::vector<std::string> texts;
  texts.reserve(count);
  std::vector<fm_phonetic_run_t> records;
  records.reserve(count);
  for (uint32_t i = 0; i < count; ++i) {
    Napi::Value element = runs.Get(i);
    if (!element.IsObject()) {
      return MakeBindingArgumentError(env, "setCellPhoneticRuns: each run must be an object { sb, eb, text }");
    }
    const Napi::Object run = element.As<Napi::Object>();
    const Napi::Value text = run.Get("text");
    texts.push_back(text.IsString() ? text.As<Napi::String>().Utf8Value() : std::string());
    records.push_back(fm_phonetic_run_t{SpecPullU32(run, "sb", 0U), SpecPullU32(run, "eb", 0U), nullptr});
  }
  for (uint32_t i = 0; i < count; ++i) {
    records[i].text = texts[i].c_str();
  }
  if (env.IsExceptionPending()) {
    // A malformed `sb`/`eb` left a pending JS exception (see
    // SpecPullU32): stop before the C ABI call commits a default value
    // for it.
    return env.Undefined();
  }
  fm_status_t rc = fm_workbook_set_cell_phonetic_runs(handle_, sheet, row, col, records.data(), records.size());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetCellPhoneticProperties(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  if (info.Length() <= 3 || !info[3].IsObject()) {
    return MakeBindingArgumentError(
        env, "setCellPhoneticProperties: `properties` must be an object { fontId, type, alignment }");
  }
  const Napi::Object props = info[3].As<Napi::Object>();
  const uint32_t font_id = SpecPullU32(props, "fontId", 0U);
  const uint32_t type = SpecPullU32(props, "type", 0U);
  const uint32_t alignment = SpecPullU32(props, "alignment", 0U);
  if (env.IsExceptionPending()) {
    // A malformed field left a pending JS exception (see SpecPullU32):
    // stop before the C ABI call commits a default value for it.
    return env.Undefined();
  }
  fm_status_t rc = fm_workbook_set_cell_phonetic_properties(handle_, sheet, row, col, font_id, type, alignment);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetBlank(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  fm_status_t rc = fm_workbook_set_blank(handle_, sheet, row, col);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetFormula(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  const std::string formula = ArgString(info, 3);
  fm_status_t rc = fm_workbook_set_formula(handle_, sheet, row, col, formula.c_str());
  return MakeStatus(env, rc);
}

// ---- Cell read ------------------------------------------------------

Napi::Value Workbook::GetValue(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeEmptyValueResult(env, NullHandleError(env));
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  fm_value_t v{};
  fm_status_t rc = fm_workbook_get_value(handle_, sheet, row, col, &v);
  return MakeValueResult(env, rc, v);
}

Napi::Value Workbook::GetCellPhonetic(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeStringFieldResult(env, NullHandleError(env), "value", "");
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  const char* text = nullptr;
  fm_status_t rc = fm_workbook_get_cell_phonetic(handle_, sheet, row, col, &text);
  if (rc != 0) {
    return MakeStringFieldResult(env, MakeErrorStatus(env, rc), "value", "");
  }
  return MakeStringFieldResult(env, MakeOkStatus(env), "value", text);
}

Napi::Value Workbook::GetCellPhoneticRuns(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeFieldResult(env, NullHandleError(env), "runs", Napi::Array::New(env, 0));
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  uint32_t count = 0;
  fm_status_t rc = fm_workbook_get_cell_phonetic_run_count(handle_, sheet, row, col, &count);
  Napi::Array out = Napi::Array::New(env, rc == 0 ? count : 0);
  for (uint32_t i = 0; rc == 0 && i < count; ++i) {
    fm_phonetic_run_t run{};
    rc = fm_workbook_get_cell_phonetic_run(handle_, sheet, row, col, i, &run);
    if (rc != 0) {
      break;
    }
    Napi::Object entry = Napi::Object::New(env);
    entry.Set("sb", Napi::Number::New(env, run.sb));
    entry.Set("eb", Napi::Number::New(env, run.eb));
    // Copied immediately: each read refreshes the handle's scratch, so the
    // previous run's pointer is dead by the time the next one lands.
    entry.Set("text", Napi::String::New(env, run.text != nullptr ? run.text : ""));
    out.Set(i, entry);
  }
  if (rc != 0) {
    return MakeFieldResult(env, MakeErrorStatus(env, rc), "runs", Napi::Array::New(env, 0));
  }
  return MakeFieldResult(env, MakeOkStatus(env), "runs", out);
}

Napi::Value Workbook::GetCellPhoneticProperties(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  uint32_t font_id = 0;
  uint32_t type = 0;
  uint32_t alignment = 0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_get_cell_phonetic_properties(handle_, sheet, row, col, &font_id, &type,
                                                                    &alignment)
                         : kBindingInvalidHandle;
  // The payload keys are declared unconditionally, so a failure reports the
  // defaults beside the status rather than dropping them.
  if (rc != 0) {
    font_id = 0;
    type = 0;
    alignment = 0;
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("fontId", Napi::Number::New(env, font_id));
  out.Set("type", Napi::Number::New(env, type));
  out.Set("alignment", Napi::Number::New(env, alignment));
  out.Set("status", MakeStatus(env, rc));
  return out;
}

Napi::Value Workbook::EvaluateFormulaText(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeEmptyValueResult(env, NullHandleError(env));
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  const std::string formula = ArgString(info, 3);
  fm_value_t v{};
  fm_status_t rc = fm_workbook_evaluate_formula(handle_, sheet, row, col, formula.c_str(), &v);
  return MakeValueResult(env, rc, v);
}

Napi::Value Workbook::EvaluateFormulaArray(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    out.Set("rows", Napi::Number::New(env, 0));
    out.Set("cols", Napi::Number::New(env, 0));
    out.Set("cells", Napi::Array::New(env, 0));
    return out;
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  const std::string formula = ArgString(info, 3);

  uint32_t rows = 0;
  uint32_t cols = 0;
  fm_status_t rc = fm_workbook_evaluate_formula_array(handle_, sheet, row, col, formula.c_str(), &rows, &cols);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    out.Set("rows", Napi::Number::New(env, 0));
    out.Set("cols", Napi::Number::New(env, 0));
    out.Set("cells", Napi::Array::New(env, 0));
    return out;
  }

  // Build a rows x cols nested array of Value objects, reading each stashed
  // cell by its row-major index (r * cols + c).
  Napi::Array cells = Napi::Array::New(env, rows);
  for (uint32_t r = 0; r < rows; ++r) {
    Napi::Array js_row = Napi::Array::New(env, cols);
    for (uint32_t c = 0; c < cols; ++c) {
      const std::size_t index = static_cast<std::size_t>(r) * cols + c;
      fm_value_t v{};
      fm_status_t cell_rc = fm_workbook_evaluate_formula_array_cell(handle_, index, &v);
      if (cell_rc != 0) {
        out.Set("status", MakeErrorStatus(env, cell_rc));
        out.Set("rows", Napi::Number::New(env, 0));
        out.Set("cols", Napi::Number::New(env, 0));
        out.Set("cells", Napi::Array::New(env, 0));
        return out;
      }
      js_row.Set(c, TranslateValue(env, v));
    }
    cells.Set(r, js_row);
  }

  out.Set("status", MakeOkStatus(env));
  out.Set("rows", Napi::Number::New(env, rows));
  out.Set("cols", Napi::Number::New(env, cols));
  out.Set("cells", cells);
  return out;
}

Napi::Value Workbook::EvaluateConditionalFormula(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeEmptyValueResult(env, NullHandleError(env));
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  const uint32_t anchor_row = ArgU32(info, 3);
  const uint32_t anchor_col = ArgU32(info, 4);
  const std::string formula = ArgString(info, 5);
  fm_value_t v{};
  fm_status_t rc =
      fm_workbook_evaluate_cf_formula(handle_, sheet, row, col, anchor_row, anchor_col, formula.c_str(), &v);
  return MakeValueResult(env, rc, v);
}

Napi::Value Workbook::GetLambdaText(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeStringFieldResult(env, NullHandleError(env), "text", "");
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  const char* text = nullptr;
  fm_status_t rc = fm_workbook_lambda_text_at(handle_, sheet, row, col, &text);
  if (rc != 0) {
    return MakeStringFieldResult(env, MakeErrorStatus(env, rc), "text", "");
  }
  return MakeStringFieldResult(env, MakeOkStatus(env), "text", text);
}

}  // namespace formulon_node
