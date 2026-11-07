// Workbook lifecycle bindings: calc policy, recalc & save, iterative-solver
// registration, diagnostics, and the `disposed` property.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

namespace {

Napi::Object ParallelRecalcStatsToJs(Napi::Env env, const fm_parallel_recalc_stats_t& stats) {
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
  return stats_obj;
}

// The saved bytes, or the save's failure; the buffer is released either way.
Napi::Value SavedBytesOrThrow(Napi::Env env, fm_status_t rc, uint8_t* buf, std::size_t len) {
  if (rc != 0) {
    fm_buffer_destroy(buf);
    return ThrowStatus(env, rc);
  }
  return TakeSavedBytes(env, buf, len);
}

// True when the optional `format` at `info[0]` was omitted.
bool FormatOmitted(const Napi::CallbackInfo& info) {
  return info.Length() == 0 || info[0].IsUndefined();
}

// Reads `threadCount`: a non-number is a TypeError and a non-integer or a
// value outside uint32 a RangeError. The C ABI rejects a count past its
// scheduler limit with the value the caller passed.
bool ReadThreadCount(const Napi::CallbackInfo& info, uint32_t& thread_count) {
  const Napi::Value value = info[0];
  if (!value.IsNumber()) {
    Napi::TypeError::New(info.Env(), "positional argument must be a number").ThrowAsJavaScriptException();
    return false;
  }
  const double number = NumberValue(value, 0.0);
  if (!std::isfinite(number) || std::trunc(number) != number || number < 0.0 || number > kUint32MaxNumber) {
    Napi::RangeError::New(info.Env(), "positional argument is outside uint32 range").ThrowAsJavaScriptException();
    return false;
  }
  thread_count = static_cast<uint32_t>(number);
  return true;
}

}  // namespace

// ---- Calc policy / behaviour profile --------------------------------

Napi::Value Workbook::CalcMode(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  std::int32_t mode = FM_CALC_MODE_AUTO;
  const fm_status_t rc = fm_workbook_get_calc_mode(handle_, &mode);
  return ValueOrThrow(env, rc, Napi::Number::New(env, mode));
}

Napi::Value Workbook::SetCalcMode(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::int32_t mode = ArgI32(info, 0);
  const fm_status_t rc = fm_workbook_set_calc_mode(handle_, mode);
  return ValueOrThrow(env, rc, env.Undefined());
}

namespace {

/// Reads one boolean workbook flag.
Napi::Value FlagResult(Napi::Env env, fm_workbook_t* handle, fm_status_t (*read)(const fm_workbook_t*, int32_t*)) {
  if (handle == nullptr) {
    return ThrowInvalidHandle(env);
  }
  int32_t value = 0;
  const fm_status_t rc = read(handle, &value);
  return ValueOrThrow(env, rc, JsBool(env, value));
}

}  // namespace

Napi::Value Workbook::FullPrecision(const Napi::CallbackInfo& info) {
  return FlagResult(info.Env(), handle_, &fm_workbook_get_full_precision);
}

Napi::Value Workbook::SetFullPrecision(const Napi::CallbackInfo& info) {
  return StatusCall(info,
                    [&](fm_workbook_t* handle) { return fm_workbook_set_full_precision(handle, ArgFlag(info, 0)); });
}

Napi::Value Workbook::CalcOnSave(const Napi::CallbackInfo& info) {
  return FlagResult(info.Env(), handle_, &fm_workbook_get_calc_on_save);
}

Napi::Value Workbook::SetCalcOnSave(const Napi::CallbackInfo& info) {
  return StatusCall(info,
                    [&](fm_workbook_t* handle) { return fm_workbook_set_calc_on_save(handle, ArgFlag(info, 0)); });
}

Napi::Value Workbook::RefMode(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  std::int32_t mode = FM_REF_MODE_A1;
  const fm_status_t rc = fm_workbook_get_ref_mode(handle_, &mode);
  return ValueOrThrow(env, rc, Napi::Number::New(env, mode));
}

Napi::Value Workbook::SetRefMode(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) { return fm_workbook_set_ref_mode(handle, ArgI32(info, 0)); });
}

Napi::Value Workbook::Date1904(const Napi::CallbackInfo& info) {
  return FlagResult(info.Env(), handle_, &fm_workbook_get_date1904);
}

Napi::Value Workbook::SetDate1904(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) { return fm_workbook_set_date1904(handle, ArgFlag(info, 0)); });
}

Napi::Value Workbook::PinnedNow(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_civil_time_t now{};
  std::int32_t pinned = 0;
  const fm_status_t rc = fm_workbook_get_pinned_now(handle_, &now, &pinned);
  if (rc != 0 || pinned == 0) {
    return OptionalOrNull(env, rc, 0, env.Null());
  }
  Napi::Object civil = Napi::Object::New(env);
  civil.Set("year", Napi::Number::New(env, now.year));
  civil.Set("month", Napi::Number::New(env, now.month));
  civil.Set("day", Napi::Number::New(env, now.day));
  civil.Set("hour", Napi::Number::New(env, now.hour));
  civil.Set("minute", Napi::Number::New(env, now.minute));
  civil.Set("second", Napi::Number::New(env, now.second));
  return civil;
}

Napi::Value Workbook::SetPinnedNow(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  // The registration guard requires all six as int32; the calendar domain
  // is validated in exactly one place, the C layer.
  fm_civil_time_t now{};
  std::int32_t* const fields[] = {&now.year, &now.month, &now.day, &now.hour, &now.minute, &now.second};
  for (std::size_t index = 0; index < 6; ++index) {
    *fields[index] = ArgI32(info, index);
  }
  return ValueOrThrow(env, fm_workbook_set_pinned_now(handle_, &now), env.Undefined());
}

Napi::Value Workbook::ClearPinnedNow(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  return ValueOrThrow(env, fm_workbook_clear_pinned_now(handle_), env.Undefined());
}

Napi::Value Workbook::ProfileId(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const char* id = nullptr;
  const fm_status_t rc = fm_workbook_get_profile_id(handle_, &id);
  return ValueOrThrow(env, rc, JsString(env, id));
}

Napi::Value Workbook::SetProfileId(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::string profile_id = ArgString(info, 0);
  const fm_status_t rc = fm_workbook_set_profile_id(handle_, profile_id.c_str());
  return ValueOrThrow(env, rc, env.Undefined());
}

Napi::Value Workbook::FileName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const char* name = nullptr;
  const fm_status_t rc = fm_workbook_get_file_name(handle_, &name);
  return ValueOrThrow(env, rc, JsStringOrNull(env, name));
}

Napi::Value Workbook::SetFileName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::string name = ArgString(info, 0);
  return ValueOrThrow(env, fm_workbook_set_file_name(handle_, name.c_str()), env.Undefined());
}

// ---- Recalc + save --------------------------------------------------

Napi::Value Workbook::Recalc(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const fm_status_t rc = fm_workbook_recalc(handle_);
  // Raise the failure before the memory query below clears the thread-local
  // diagnostics. An exception the progress callback raised stays pending,
  // and that exception is what the caller sees.
  const Napi::Value result = ValueOrThrow(env, rc, env.Undefined());
  // A full recalc is the coarsest boundary the binding has and the one
  // after which the footprint has most likely moved (spilled arrays,
  // newly cached text), so the external-memory figure is refreshed here
  // rather than on every cell write.
  SyncExternalMemory(env);
  return result;
}

Napi::Value Workbook::RecalcParallel(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  uint32_t thread_count = 0;
  if (!ReadThreadCount(info, thread_count)) {
    return env.Undefined();
  }
  fm_parallel_recalc_stats_t stats = FM_STRUCT_INIT(fm_parallel_recalc_stats_t);
  const fm_status_t rc = fm_workbook_recalc_parallel(handle_, thread_count, &stats);
  const Napi::Value result =
      rc == 0 && !env.IsExceptionPending() ? Napi::Value(ParallelRecalcStatsToJs(env, stats)) : ThrowStatus(env, rc);
  // Parallel recalc can change cached values and spill geometry just like the
  // serial entry point, so keep V8's external-memory estimate in sync.
  SyncExternalMemory(env);
  return result;
}

Napi::Value Workbook::PartialRecalc(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = ArgU32(info, 0);
  fm_cell_rect_t vp{};
  CheckedSpecReader reader(env);
  if (!ReadCellRect(reader, ArgObject(info, 1), &vp)) {
    return env.Undefined();
  }
  uint32_t recomputed = 0;
  const fm_status_t rc = fm_sheet_recalc_viewport(handle_, sheet, &vp, &recomputed);
  return ValueOrThrow(env, rc, Napi::Number::New(env, recomputed));
}

Napi::Value Workbook::HasDirtyCells(const Napi::CallbackInfo& info) {
  return FlagResult(info.Env(), handle_, &fm_workbook_has_dirty_cells);
}

Napi::Value Workbook::SetIterative(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  // Every field is non-nullable: an omitted or `undefined` key keeps the
  // stored value, `null` and a non-integer cap are type errors.
  const Napi::Object patch = ArgObject(info, 0);
  CheckedSpecReader reader(env);
  fm_iterative_settings_patch_t p = FM_STRUCT_INIT(fm_iterative_settings_patch_t);
  p.enabled = -1;
  p.max_iterations = -1;
  p.max_change = -1.0;
  const Napi::Value enabled = reader.Get(patch, "enabled");
  if (reader.ok() && !enabled.IsUndefined()) {
    if (!enabled.IsBoolean()) {
      reader.FailType("enabled", "boolean");
    } else {
      p.enabled = enabled.As<Napi::Boolean>().Value() ? 1 : 0;
    }
  }
  const Napi::Value cap = reader.Get(patch, "maxIterations");
  if (reader.ok() && !cap.IsUndefined()) {
    const double requested = cap.IsNumber() ? cap.As<Napi::Number>().DoubleValue() : 0.0;
    if (!cap.IsNumber() || !std::isfinite(requested) || std::trunc(requested) != requested) {
      reader.FailType("maxIterations", "whole number");
    } else if (requested < 0.0 || requested > static_cast<double>(std::numeric_limits<int32_t>::max())) {
      reader.FailRange("maxIterations", "[0, 2147483647]");
    } else {
      p.max_iterations = static_cast<int32_t>(requested);
    }
  }
  const Napi::Value change = reader.Get(patch, "maxChange");
  if (reader.ok() && !change.IsUndefined()) {
    if (!change.IsNumber()) {
      reader.FailType("maxChange", "number");
    } else {
      p.max_change = change.As<Napi::Number>().DoubleValue();
    }
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_workbook_set_iterative(handle_, &p), env.Undefined());
}

Napi::Value Workbook::GetIterative(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  int32_t enabled = 0;
  uint32_t max_iterations = 0;
  double max_change = 0.0;
  const fm_status_t rc = fm_workbook_get_iterative(handle_, &enabled, &max_iterations, &max_change);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("enabled", Napi::Boolean::New(env, enabled != 0));
  out.Set("maxIterations", Napi::Number::New(env, max_iterations));
  out.Set("maxChange", Napi::Number::New(env, max_change));
  return out;
}

Napi::Value Workbook::SetIterativeProgress(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  // Passing null clears the callback. Anything else MUST be a JS function;
  // a rejected value leaves the installed callback in place.
  Napi::Function callback;
  if (info[0].IsNull()) {
    (void)Value().Delete(IterativeProgressKey(env));
    return ValueOrThrow(env, fm_workbook_set_iterative_progress(handle_, nullptr, nullptr), env.Undefined());
  }
  if (!FunctionArg(info, 0, &callback)) {
    Napi::TypeError::New(env, "callback must be a function or null").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  // The function lives on this wrapper, and the C ABI hands the trampoline
  // this wrapper as user-data, so replacing a callback on another Workbook
  // never changes this instance's callback.
  (void)Value().Set(IterativeProgressKey(env), callback);
  return ValueOrThrow(env, fm_workbook_set_iterative_progress(handle_, &Workbook::IterativeProgressTrampoline, this),
                      env.Undefined());
}

int32_t Workbook::IterativeProgressTrampoline(uint32_t iteration, double max_residual, uint32_t max_iterations,
                                              void* user_data) {
  auto* const workbook = static_cast<Workbook*>(user_data);
  if (workbook == nullptr) {
    return 1;
  }
  Napi::Env env = workbook->Env();
  // A pending exception was raised by the callback in an earlier sweep:
  // keep aborting, and leave it pending so the recalc entry point returns it
  // to JS unchanged.
  if (env.IsExceptionPending()) {
    return 0;
  }
  Napi::HandleScope scope(env);
  const Napi::Value callback = workbook->Value().Get(IterativeProgressKey(env));
  if (env.IsExceptionPending()) {
    return 0;
  }
  if (!callback.IsFunction()) {
    return 1;
  }
  workbook->in_iterative_progress_callback_ = true;
  const Napi::Value ret = callback.As<Napi::Function>().Call({
      Napi::Number::New(env, iteration),
      Napi::Number::New(env, max_residual),
      Napi::Number::New(env, max_iterations),
  });
  workbook->in_iterative_progress_callback_ = false;
  if (env.IsExceptionPending()) {
    return 0;
  }
  return (ret.IsUndefined() || ret.IsNull() || ret.ToBoolean().Value()) ? 1 : 0;
}

Napi::Value Workbook::Dispose(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  // This call is one of the active ones; any other means a method of this
  // workbook is still on the stack, reached here through a getter or a
  // callback, and is about to use the handle again.
  if (active_calls_ > 1) {
    return ThrowReentrant(env, "cannot dispose Workbook during an active method call (kGraphRecalcReentrant, 4005)");
  }
  (void)Value().Delete(IterativeProgressKey(env));
  DestroyHandle(env);
  return env.Undefined();
}

Napi::Value Workbook::Disposed(const Napi::CallbackInfo& info) {
  return Napi::Boolean::New(info.Env(), handle_ == nullptr);
}

Napi::Value Workbook::MemoryUsage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  size_t bytes = 0;
  if (fm_workbook_get_memory_usage(handle_, &bytes) != 0) {
    return Napi::Number::New(env, 0);
  }
  // Re-report the measurement already in hand: a script that has been
  // filling cells since the last sync has grown the workbook without V8
  // hearing about it, and this is the natural moment to correct that.
  ApplyExternalMemoryEstimate(env, bytes);
  return Napi::Number::New(env, static_cast<double>(bytes));
}

Napi::Value Workbook::Save(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  uint8_t* buf = nullptr;
  std::size_t len = 0;
  const fm_status_t rc = FormatOmitted(info) ? fm_workbook_save(handle_, &buf, &len)
                                             : fm_workbook_save_as(handle_, ArgI32(info, 0), &buf, &len);
  return SavedBytesOrThrow(env, rc, buf, len);
}

Napi::Value Workbook::SaveWithDiagnostics(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const int32_t format = FormatOmitted(info) ? static_cast<int32_t>(FM_WORKBOOK_FORMAT_XLSX) : ArgI32(info, 0);
  uint8_t* buf = nullptr;
  std::size_t len = 0;
  fm_save_diagnostics_t d = FM_STRUCT_INIT(fm_save_diagnostics_t);
  const fm_status_t rc = fm_workbook_save_with_diagnostics(handle_, format, &buf, &len, &d);
  if (rc != 0) {
    fm_buffer_destroy(buf);
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("bytes", TakeSavedBytes(env, buf, len));
  out.Set("downgradedFormulaCount", Napi::Number::New(env, static_cast<double>(d.downgraded_formula_count)));
  out.Set("deferredFeatureCount", Napi::Number::New(env, static_cast<double>(d.deferred_feature_count)));
  out.Set("droppedPartCount", Napi::Number::New(env, static_cast<double>(d.dropped_part_count)));
  out.Set("droppedRelationshipCount", Napi::Number::New(env, static_cast<double>(d.dropped_relationship_count)));
  out.Set("renumberedPartCount", Napi::Number::New(env, static_cast<double>(d.renumbered_part_count)));
  return out;
}

Napi::Value Workbook::ReadDiagnostics(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_read_diagnostics_t d = FM_STRUCT_INIT(fm_read_diagnostics_t);
  const fm_status_t rc = fm_workbook_get_diagnostics(handle_, &d);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("undecodedFormulaCount", Napi::Number::New(env, static_cast<double>(d.undecoded_formula_count)));
  out.Set("undecodedDefinedNameCount", Napi::Number::New(env, static_cast<double>(d.undecoded_defined_name_count)));
  out.Set("undecodedPartCount", Napi::Number::New(env, static_cast<double>(d.undecoded_part_count)));
  out.Set("skippedFeatureCount", Napi::Number::New(env, static_cast<double>(d.skipped_feature_count)));
  out.Set("unknownContentTypeCount", Napi::Number::New(env, static_cast<double>(d.unknown_content_type_count)));
  return out;
}

}  // namespace formulon_node
