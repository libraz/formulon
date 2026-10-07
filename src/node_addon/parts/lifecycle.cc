// Workbook lifecycle bindings: calc policy, recalc & save, iterative-solver
// registration, diagnostics, and the trivial `isValid` predicate.

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

// Copies an engine-owned save buffer onto the JS heap and releases it.
Napi::Uint8Array TakeSavedBytes(Napi::Env env, uint8_t* buf, std::size_t len) {
  Napi::Uint8Array dst = Napi::Uint8Array::New(env, len);
  if (len != 0 && buf != nullptr) {
    std::memcpy(dst.Data(), buf, len);
  }
  fm_buffer_free(buf);
  return dst;
}

// Builds `{ status, bytes }` for a save call; `bytes` is null on failure.
Napi::Object MakeSaveResult(Napi::Env env, fm_status_t rc, uint8_t* buf, std::size_t len) {
  if (rc != 0) {
    return MakeFieldResult(env, MakeErrorStatus(env, rc), "bytes", env.Null());
  }
  Napi::Uint8Array dst = TakeSavedBytes(env, buf, len);
  return MakeFieldResult(env, MakeOkStatus(env), "bytes", dst);
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

// ---- Calc policy / behaviour profile --------------------------------

Napi::Value Workbook::CalcMode(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberResult(env, kBindingInvalidHandle, 0);
  }
  fm_calc_mode_t mode = FM_CALC_MODE_AUTO;
  const fm_status_t rc = fm_workbook_calc_mode(handle_, &mode);
  return MakeNumberResult(env, rc, static_cast<int32_t>(mode));
}

Napi::Value Workbook::SetCalcMode(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::int32_t mode = ArgI32(info, 0);
  fm_status_t rc = fm_workbook_set_calc_mode(handle_, mode);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PinnedNow(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeFieldResult(env, NullHandleError(env), "now", env.Null());
  }
  fm_civil_time_t now{};
  std::int32_t pinned = 0;
  const fm_status_t rc = fm_workbook_pinned_now(handle_, &now, &pinned);
  if (rc != 0 || pinned == 0) {
    return MakeFieldResult(env, MakeStatus(env, rc), "now", env.Null());
  }
  Napi::Object civil = Napi::Object::New(env);
  civil.Set("year", Napi::Number::New(env, now.year));
  civil.Set("month", Napi::Number::New(env, now.month));
  civil.Set("day", Napi::Number::New(env, now.day));
  civil.Set("hour", Napi::Number::New(env, now.hour));
  civil.Set("minute", Napi::Number::New(env, now.minute));
  civil.Set("second", Napi::Number::New(env, now.second));
  return MakeFieldResult(env, MakeOkStatus(env), "now", civil);
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
    return MakeStringResult(env, kBindingInvalidHandle, nullptr);
  }
  const char* id = nullptr;
  const fm_status_t rc = fm_workbook_excel_profile_id(handle_, &id);
  return MakeStringResult(env, rc, id);
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
    return MakeCallbackThrewStatus(env);
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
    // Reject here instead of forwarding a sentinel to the C ABI: its own
    // out-of-range diagnostic would otherwise report a thread_count value
    // the caller never passed, violating the Status contract that a
    // binding-raised failure carries an empty context.
    Napi::Object status =
        MakeBindingError(env, kInvalidArgument, "recalcParallel: `threadCount` must be an integer in 0..8");
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
    return MakeParallelRecalcResult(env, MakeCallbackThrewStatus(env), stats);
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
    CheckedSpecReader reader(env);
    vp.sheet = reader.U32(vpobj, "sheet", 0U);
    vp.first_row = reader.U32(vpobj, "firstRow", 0U);
    vp.last_row = reader.U32(vpobj, "lastRow", 0U);
    vp.first_col = reader.U32(vpobj, "firstCol", 0U);
    vp.last_col = reader.U32(vpobj, "lastCol", 0U);
    if (!reader.ok()) {
      return env.Undefined();
    }
  }
  uint32_t recomputed = 0;
  (void)TakeIterativeProgressThrew();
  fm_status_t rc = fm_workbook_partial_recalc(handle_, &vp, &recomputed);
  const bool callback_threw = TakeIterativeProgressThrew();
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "recomputed", 0);
  }
  if (callback_threw) {
    return MakeNumberFieldResult(env, MakeCallbackThrewStatus(env), "recomputed", 0);
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

Napi::Value Workbook::GetIterative(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  int32_t enabled = 0;
  uint32_t max_iterations = 0;
  double max_change = 0.0;
  const fm_status_t rc = handle_ != nullptr ? fm_workbook_get_iterative(handle_, &enabled, &max_iterations, &max_change)
                                            : kBindingInvalidHandle;
  if (rc != 0) {
    enabled = 0;
    max_iterations = 0;
    max_change = 0.0;
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", MakeStatus(env, rc));
  out.Set("enabled", Napi::Boolean::New(env, enabled != 0));
  out.Set("maxIterations", Napi::Number::New(env, max_iterations));
  out.Set("maxChange", Napi::Number::New(env, max_change));
  return out;
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
    return MakeBindingArgumentError(env, "setIterativeProgress: `callback` must be a function or null");
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
  uint8_t* buf = nullptr;
  std::size_t len = 0;
  const fm_status_t rc = handle_ != nullptr ? fm_workbook_save(handle_, &buf, &len) : kBindingInvalidHandle;
  return MakeSaveResult(info.Env(), rc, buf, len);
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
  uint8_t* buf = nullptr;
  std::size_t len = 0;
  fm_status_t rc = kBindingInvalidHandle;
  if (handle_ != nullptr) {
    const std::int32_t format = info[0].As<Napi::Number>().Int32Value();
    rc = fm_workbook_save_as(handle_, format, &buf, &len);
  }
  return MakeSaveResult(env, rc, buf, len);
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
  Napi::Uint8Array dst = TakeSavedBytes(env, buf, len);
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
