// Workbook lifecycle bindings: cell mutation / read, recalc & save,
// iterative-solver registration, and the trivial `isValid` predicate.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>

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
  if (rc != 0) {
    return MakeEmptyValueResult(env, MakeErrorStatus(env, rc));
  }
  return MakeValueResult(env, MakeOkStatus(env), v);
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
  if (rc != 0) {
    return MakeEmptyValueResult(env, MakeErrorStatus(env, rc));
  }
  return MakeValueResult(env, MakeOkStatus(env), v);
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
  if (rc != 0) {
    return MakeEmptyValueResult(env, MakeErrorStatus(env, rc));
  }
  return MakeValueResult(env, MakeOkStatus(env), v);
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

// ---- Calc policy / behaviour profile --------------------------------

Napi::Value Workbook::CalcMode(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, static_cast<int32_t>(FM_CALC_MODE_AUTO));
  }
  fm_calc_mode_t mode = FM_CALC_MODE_AUTO;
  fm_workbook_calc_mode(handle_, &mode);
  return Napi::Number::New(env, static_cast<int32_t>(mode));
}

Napi::Value Workbook::SetCalcMode(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t mode = ArgU32(info, 0);
  fm_status_t rc = fm_workbook_set_calc_mode(handle_, static_cast<fm_calc_mode_t>(mode));
  return MakeStatus(env, rc);
}

Napi::Value Workbook::ExcelProfileId(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::String::New(env, "win-365-ja_JP");
  }
  const char* id = nullptr;
  fm_workbook_excel_profile_id(handle_, &id);
  return Napi::String::New(env, id != nullptr ? id : "win-365-ja_JP");
}

Napi::Value Workbook::SetExcelProfileId(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::string profile_id = ArgString(info, 0);
  fm_status_t rc = fm_workbook_set_excel_profile_id(handle_, profile_id.c_str());
  return MakeStatus(env, rc);
}

// ---- Recalc + save --------------------------------------------------

Napi::Value Workbook::Recalc(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  fm_status_t rc = fm_workbook_recalc(handle_);
  // A full recalc is the coarsest boundary the binding has and the one
  // after which the footprint has most likely moved (spilled arrays,
  // newly cached text), so the external-memory figure is refreshed here
  // rather than on every cell write.
  SyncExternalMemory(env);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PartialRecalc(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "recomputed", 0);
  }
  // Embind takes the viewport as a JS object; we mirror that shape.
  fm_viewport vp{};
  if (info.Length() > 0 && info[0].IsObject()) {
    Napi::Object vpobj = info[0].As<Napi::Object>();
    vp.sheet = vpobj.Get("sheet").ToNumber().Uint32Value();
    vp.first_row = vpobj.Get("firstRow").ToNumber().Uint32Value();
    vp.last_row = vpobj.Get("lastRow").ToNumber().Uint32Value();
    vp.first_col = vpobj.Get("firstCol").ToNumber().Uint32Value();
    vp.last_col = vpobj.Get("lastCol").ToNumber().Uint32Value();
  }
  uint32_t recomputed = 0;
  fm_status_t rc = fm_workbook_partial_recalc(handle_, &vp, &recomputed);
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "recomputed", 0);
  }
  return MakeNumberFieldResult(env, MakeOkStatus(env), "recomputed", recomputed);
}

Napi::Value Workbook::SetIterative(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const bool enabled = ArgBool(info, 0);
  const uint32_t max_iter = ArgU32(info, 1);
  const double max_change = ArgDouble(info, 2);
  fm_status_t rc = fm_workbook_set_iterative(handle_, enabled ? 1 : 0, static_cast<int32_t>(max_iter), max_change);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetIterativeProgress(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  // Passing null / undefined clears the callback. Anything else MUST be
  // a JS function -- we surface a 7000-band error if it is not.
  if (info.Length() < 1 || info[0].IsNull() || info[0].IsUndefined()) {
    iterative_progress_callback_.Reset();
    fm_status_t rc = fm_workbook_set_iterative_progress(handle_, nullptr, nullptr);
    return MakeStatus(env, rc);
  }
  if (!info[0].IsFunction()) {
    return MakeErrorStatus(env, kBindingInvalidHandle);
  }
  // Persist the function on this wrapper, then let the C ABI give the
  // trampoline this wrapper as user-data. Replacing a callback on another
  // Workbook never changes this instance's callback.
  iterative_progress_callback_.Reset();
  iterative_progress_callback_ = Napi::Persistent(info[0].As<Napi::Function>());
  fm_status_t rc = fm_workbook_set_iterative_progress(handle_, &Workbook::IterativeProgressTrampoline, this);
  return MakeStatus(env, rc);
}

bool Workbook::IterativeProgressTrampoline(uint32_t iteration, double max_residual, uint32_t max_iterations,
                                           void* user_data) {
  auto* const workbook = static_cast<Workbook*>(user_data);
  if (workbook == nullptr || workbook->iterative_progress_callback_.IsEmpty()) {
    return true;
  }
  Napi::Env env = workbook->iterative_progress_callback_.Env();
  Napi::HandleScope scope(env);
  workbook->in_iterative_progress_callback_ = true;
  Napi::Value ret = workbook->iterative_progress_callback_.Call({
      Napi::Number::New(env, iteration),
      Napi::Number::New(env, max_residual),
      Napi::Number::New(env, max_iterations),
  });
  workbook->in_iterative_progress_callback_ = false;
  if (env.IsExceptionPending()) {
    (void)env.GetAndClearPendingException();
    return false;
  }
  return ret.IsUndefined() || ret.IsNull() || ret.ToBoolean().Value();
}

Napi::Value Workbook::Dispose(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (in_iterative_progress_callback_) {
    Napi::Error::New(env, "cannot dispose a Workbook from its iterative progress callback")
        .ThrowAsJavaScriptException();
    return env.Undefined();
  }
  DestroyHandle(env);
  return env.Undefined();
}

Napi::Value Workbook::IsValid(const Napi::CallbackInfo& info) {
  return Napi::Boolean::New(info.Env(), handle_ != nullptr);
}

Napi::Value Workbook::MemoryUsage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  size_t bytes = 0;
  if (fm_workbook_memory_usage(handle_, &bytes) != 0) {
    return Napi::Number::New(env, 0);
  }
  // Re-report while the figure is in hand: a script that has been
  // filling cells since the last sync has grown the workbook without V8
  // hearing about it, and this is the natural moment to correct that.
  SyncExternalMemory(env);
  return Napi::Number::New(env, static_cast<double>(bytes));
}

Napi::Value Workbook::Save(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    out.Set("bytes", env.Null());
    return out;
  }
  uint8_t* buf = nullptr;
  std::size_t len = 0;
  fm_status_t rc = fm_workbook_save(handle_, &buf, &len);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    out.Set("bytes", env.Null());
    return out;
  }
  // Copy into a fresh Uint8Array on the JS heap; the C-side buffer is
  // owned by the engine and must be released with `fm_buffer_free`.
  Napi::Uint8Array dst = Napi::Uint8Array::New(env, len);
  if (len != 0 && buf != nullptr) {
    std::memcpy(dst.Data(), buf, len);
  }
  fm_buffer_free(buf);
  out.Set("status", MakeOkStatus(env));
  out.Set("bytes", dst);
  return out;
}

Napi::Value Workbook::SaveEx(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (info.Length() < 1) {
    // `format` has no sane default (unlike the other setters' 0-valued
    // fallbacks): a silent default would pick a container format the
    // caller never asked for. Reject like the WASM binding (embind
    // arity check) and the Python binding (required positional arg).
    Napi::TypeError::New(env, "saveEx requires 1 argument (format)").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    out.Set("bytes", env.Null());
    return out;
  }
  const auto format = static_cast<fm_workbook_format_t>(info[0].As<Napi::Number>().Int32Value());
  uint8_t* buf = nullptr;
  std::size_t len = 0;
  fm_status_t rc = fm_workbook_save_ex(handle_, format, &buf, &len);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    out.Set("bytes", env.Null());
    return out;
  }
  Napi::Uint8Array dst = Napi::Uint8Array::New(env, len);
  if (len != 0 && buf != nullptr) {
    std::memcpy(dst.Data(), buf, len);
  }
  fm_buffer_free(buf);
  out.Set("status", MakeOkStatus(env));
  out.Set("bytes", dst);
  return out;
}

Napi::Value Workbook::SaveExWithDiagnostics(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (info.Length() < 1) {
    Napi::TypeError::New(env, "saveExWithDiagnostics requires 1 argument (format)").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("bytes", env.Null());
  out.Set("downgradedFormulaCount", Napi::Number::New(env, 0));
  out.Set("deferredFeatureCount", Napi::Number::New(env, 0));
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const auto format = static_cast<fm_workbook_format_t>(info[0].As<Napi::Number>().Int32Value());
  uint8_t* buf = nullptr;
  std::size_t len = 0;
  std::size_t downgraded_formula_count = 0;
  std::size_t deferred_feature_count = 0;
  fm_status_t rc = fm_workbook_save_ex_with_diagnostics(handle_, format, &buf, &len, &downgraded_formula_count,
                                                        &deferred_feature_count);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  Napi::Uint8Array dst = Napi::Uint8Array::New(env, len);
  if (len != 0 && buf != nullptr) {
    std::memcpy(dst.Data(), buf, len);
  }
  fm_buffer_free(buf);
  out.Set("status", MakeOkStatus(env));
  out.Set("bytes", dst);
  out.Set("downgradedFormulaCount", Napi::Number::New(env, static_cast<double>(downgraded_formula_count)));
  out.Set("deferredFeatureCount", Napi::Number::New(env, static_cast<double>(deferred_feature_count)));
  return out;
}

Napi::Value Workbook::XlsbReadDiagnostics(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("undecodedFormulaCount", Napi::Number::New(env, 0));
  out.Set("undecodedDefinedNameCount", Napi::Number::New(env, 0));
  out.Set("droppedPartCount", Napi::Number::New(env, 0));
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  std::size_t undecoded_formula_count = 0;
  std::size_t undecoded_defined_name_count = 0;
  std::size_t dropped_part_count = 0;
  fm_status_t rc = fm_workbook_xlsb_read_diagnostics_ex(handle_, &undecoded_formula_count,
                                                        &undecoded_defined_name_count, &dropped_part_count);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("undecodedFormulaCount", Napi::Number::New(env, static_cast<double>(undecoded_formula_count)));
  out.Set("undecodedDefinedNameCount", Napi::Number::New(env, static_cast<double>(undecoded_defined_name_count)));
  out.Set("droppedPartCount", Napi::Number::New(env, static_cast<double>(dropped_part_count)));
  return out;
}

}  // namespace formulon_node
