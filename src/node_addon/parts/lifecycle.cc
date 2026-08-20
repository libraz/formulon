// Workbook lifecycle bindings: cell mutation / read, recalc & save,
// iterative-solver registration, and the trivial `isValid` predicate.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

namespace {

Napi::Object MakeParallelRecalcResult(Napi::Env env, Napi::Object status, const fm_parallel_recalc_stats& stats) {
  Napi::Object stats_obj = Napi::Object::New(env);
  // The C ABI counters are uint64_t. The engine's counters are bounded well
  // below Number.MAX_SAFE_INTEGER for a single call, so this conversion is
  // exact while preserving the package's existing JS-number surface.
  stats_obj.Set("cellsEvaluated", Napi::Number::New(env, static_cast<double>(stats.cells_evaluated)));
  stats_obj.Set("sccsProcessed", Napi::Number::New(env, static_cast<double>(stats.sccs_processed)));
  stats_obj.Set("parallelSteps", Napi::Number::New(env, static_cast<double>(stats.parallel_steps)));
  stats_obj.Set("serialFallbackSteps", Napi::Number::New(env, static_cast<double>(stats.serial_fallback_steps)));
  stats_obj.Set("cycleRecoveries", Napi::Number::New(env, static_cast<double>(stats.cycle_recoveries)));
  stats_obj.Set("workerThreadsStarted", Napi::Number::New(env, static_cast<double>(stats.worker_threads_started)));
  stats_obj.Set("workerThreadsUsed", Napi::Number::New(env, static_cast<double>(stats.worker_threads_used)));

  Napi::Object result = Napi::Object::New(env);
  result.Set("status", status);
  result.Set("stats", stats_obj);
  return result;
}

bool ReadThreadCount(const Napi::CallbackInfo& info, uint32_t& thread_count) {
  if (info.Length() == 0 || !info[0].IsNumber()) {
    return false;
  }
  const double value = info[0].As<Napi::Number>().DoubleValue();
  if (!std::isfinite(value) || std::trunc(value) != value || value < 0.0 || value > 8.0) {
    return false;
  }
  thread_count = static_cast<uint32_t>(value);
  return true;
}

}  // namespace

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
  const std::int32_t mode = info.Length() > 0 ? info[0].ToNumber().Int32Value() : 0;
  fm_status_t rc = fm_workbook_set_calc_mode(handle_, mode);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PinnedNow(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return env.Null();
  }
  fm_civil_time_t now{};
  std::int32_t pinned = 0;
  if (fm_workbook_pinned_now(handle_, &now, &pinned) != 0 || pinned == 0) {
    return env.Null();
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("year", Napi::Number::New(env, now.year));
  out.Set("month", Napi::Number::New(env, now.month));
  out.Set("day", Napi::Number::New(env, now.day));
  out.Set("hour", Napi::Number::New(env, now.hour));
  out.Set("minute", Napi::Number::New(env, now.minute));
  out.Set("second", Napi::Number::New(env, now.second));
  return out;
}

Napi::Value Workbook::SetPinnedNow(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  // Missing arguments fall through as 0, which the C layer rejects: the
  // calendar domain is validated in exactly one place.
  fm_civil_time_t now{};
  std::int32_t* const fields[] = {&now.year, &now.month, &now.day, &now.hour, &now.minute, &now.second};
  for (std::size_t index = 0; index < 6; ++index) {
    *fields[index] = info.Length() > index ? info[index].ToNumber().Int32Value() : 0;
  }
  return MakeStatus(env, fm_workbook_set_pinned_now(handle_, &now));
}

Napi::Value Workbook::ClearPinnedNow(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  return MakeStatus(env, fm_workbook_clear_pinned_now(handle_));
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
  (void)TakeIterativeProgressThrew();
  fm_status_t rc = fm_workbook_recalc(handle_);
  const bool callback_threw = TakeIterativeProgressThrew();
  // A full recalc is the coarsest boundary the binding has and the one
  // after which the footprint has most likely moved (spilled arrays,
  // newly cached text), so the external-memory figure is refreshed here
  // rather than on every cell write.
  SyncExternalMemory(env);
  // An engine failure wins: it carries its own diagnostic and is the more
  // specific answer. Otherwise a throwing progress callback, which the
  // engine only saw as a cancellation, becomes the reported failure.
  if (rc == 0 && callback_threw) {
    return MakeErrorStatus(env, kBindingCallbackException);
  }
  return MakeStatus(env, rc);
}

Napi::Value Workbook::RecalcParallel(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  fm_parallel_recalc_stats stats{};
  if (handle_ == nullptr) {
    return MakeParallelRecalcResult(env, NullHandleError(env), stats);
  }

  uint32_t thread_count = 0;
  if (!ReadThreadCount(info, thread_count)) {
    // Reuse the C ABI's canonical invalid-thread-count diagnostic. The C
    // entry point validates before touching the workbook, and zeroes stats
    // on this failure path just as it does for a direct caller.
    const fm_status_t rc = fm_workbook_recalc_parallel(handle_, 9U, &stats);
    Napi::Object status = MakeStatus(env, rc);
    SyncExternalMemory(env);
    return MakeParallelRecalcResult(env, status, stats);
  }

  (void)TakeIterativeProgressThrew();
  const fm_status_t rc = fm_workbook_recalc_parallel(handle_, thread_count, &stats);
  const bool callback_threw = TakeIterativeProgressThrew();
  if (rc == 0 && callback_threw) {
    // Same shape as an engine failure: the pass did not finish, so the
    // counters are not reported.
    stats = fm_parallel_recalc_stats{};
    SyncExternalMemory(env);
    return MakeParallelRecalcResult(env, MakeErrorStatus(env, kBindingCallbackException), stats);
  }
  Napi::Object status = MakeStatus(env, rc);
  // Parallel recalc can change cached values and spill geometry just like the
  // serial entry point, so keep V8's external-memory estimate in sync before
  // returning the result envelope.
  SyncExternalMemory(env);
  return MakeParallelRecalcResult(env, status, stats);
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
  (void)TakeIterativeProgressThrew();
  fm_status_t rc = fm_workbook_partial_recalc(handle_, &vp, &recomputed);
  const bool callback_threw = TakeIterativeProgressThrew();
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "recomputed", 0);
  }
  if (callback_threw) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, kBindingCallbackException), "recomputed", 0);
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

int32_t Workbook::IterativeProgressTrampoline(uint32_t iteration, double max_residual, uint32_t max_iterations,
                                              void* user_data) {
  auto* const workbook = static_cast<Workbook*>(user_data);
  if (workbook == nullptr || workbook->iterative_progress_callback_.IsEmpty()) {
    return 1;
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
    // Abort the solve and leave the reason behind for the recalc entry
    // point: without it the throw is indistinguishable from a deliberate
    // cancel and `recalc()` reports success over a half-solved workbook.
    (void)env.GetAndClearPendingException();
    workbook->iterative_progress_threw_ = true;
    return 0;
  }
  return (ret.IsUndefined() || ret.IsNull() || ret.ToBoolean().Value()) ? 1 : 0;
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

Napi::Value Workbook::SaveAs(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (info.Length() < 1) {
    // `format` has no sane default (unlike the other setters' 0-valued
    // fallbacks): a silent default would pick a container format the
    // caller never asked for. Reject like the WASM binding (embind
    // arity check) and the Python binding (required positional arg).
    Napi::TypeError::New(env, "saveAs requires 1 argument (format)").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    out.Set("bytes", env.Null());
    return out;
  }
  const std::int32_t format = info[0].As<Napi::Number>().Int32Value();
  uint8_t* buf = nullptr;
  std::size_t len = 0;
  fm_status_t rc = fm_workbook_save_as(handle_, format, &buf, &len);
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

Napi::Value Workbook::SaveWithDiagnostics(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (info.Length() < 1) {
    Napi::TypeError::New(env, "saveWithDiagnostics requires 1 argument (format)").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("bytes", env.Null());
  out.Set("downgradedFormulaCount", Napi::Number::New(env, 0));
  out.Set("deferredFeatureCount", Napi::Number::New(env, 0));
  out.Set("droppedPartCount", Napi::Number::New(env, 0));
  out.Set("droppedRelationshipCount", Napi::Number::New(env, 0));
  out.Set("renumberedPartCount", Napi::Number::New(env, 0));
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const std::int32_t format = info[0].As<Napi::Number>().Int32Value();
  uint8_t* buf = nullptr;
  std::size_t len = 0;
  fm_save_diagnostics_t d{};
  fm_status_t rc = fm_workbook_save_with_diagnostics(handle_, format, &buf, &len, &d);
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
  out.Set("downgradedFormulaCount", Napi::Number::New(env, static_cast<double>(d.downgraded_formula_count)));
  out.Set("deferredFeatureCount", Napi::Number::New(env, static_cast<double>(d.deferred_feature_count)));
  out.Set("droppedPartCount", Napi::Number::New(env, static_cast<double>(d.dropped_part_count)));
  out.Set("droppedRelationshipCount", Napi::Number::New(env, static_cast<double>(d.dropped_relationship_count)));
  out.Set("renumberedPartCount", Napi::Number::New(env, static_cast<double>(d.renumbered_part_count)));
  return out;
}

Napi::Value Workbook::ReadDiagnostics(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("undecodedFormulaCount", Napi::Number::New(env, 0));
  out.Set("undecodedDefinedNameCount", Napi::Number::New(env, 0));
  out.Set("undecodedPartCount", Napi::Number::New(env, 0));
  out.Set("skippedFeatureCount", Napi::Number::New(env, 0));
  out.Set("unknownContentTypeCount", Napi::Number::New(env, 0));
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  fm_read_diagnostics_t d{};
  fm_status_t rc = fm_workbook_read_diagnostics(handle_, &d);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("undecodedFormulaCount", Napi::Number::New(env, static_cast<double>(d.undecoded_formula_count)));
  out.Set("undecodedDefinedNameCount", Napi::Number::New(env, static_cast<double>(d.undecoded_defined_name_count)));
  out.Set("undecodedPartCount", Napi::Number::New(env, static_cast<double>(d.undecoded_part_count)));
  out.Set("skippedFeatureCount", Napi::Number::New(env, static_cast<double>(d.skipped_feature_count)));
  out.Set("unknownContentTypeCount", Napi::Number::New(env, static_cast<double>(d.unknown_content_type_count)));
  return out;
}

}  // namespace formulon_node
