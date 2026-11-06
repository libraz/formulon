// Implementation of the module-level free functions exported alongside
// the `Workbook` class. See `free_funcs.h` for the contract.

#include "node_addon/parts/free_funcs.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <mutex>
#include <string>
#include <vector>

namespace formulon_node {
namespace {

// Structured-log delivery is process-wide, so the sink state is too. The
// engine may emit a record from any thread it owns, which rules out
// calling the JS function directly: everything goes through a
// ThreadSafeFunction. Delivery is therefore asynchronous even though the
// C ABI invokes the adapter synchronously -- a record produced while the
// JS thread is blocked inside a native call is queued and drained when
// that call returns.
std::mutex& log_sink_mutex() {
  static std::mutex m;
  return m;
}

Napi::ThreadSafeFunction& log_sink_tsfn() {
  static Napi::ThreadSafeFunction tsfn;
  return tsfn;
}

void ReleaseLogSinkLocked() {
  Napi::ThreadSafeFunction& tsfn = log_sink_tsfn();
  if (tsfn) {
    tsfn.Release();
    tsfn = Napi::ThreadSafeFunction();
  }
}

void LogSinkTrampoline(const char* record, std::size_t len, void* /*user_data*/) {
  const std::lock_guard<std::mutex> lock(log_sink_mutex());
  Napi::ThreadSafeFunction& tsfn = log_sink_tsfn();
  if (!tsfn) {
    return;
  }
  // The record view dies with this call, so copy before queueing.
  auto* payload = new std::vector<std::uint8_t>(record, record + len);
  const napi_status rc =
      tsfn.NonBlockingCall(payload, [](Napi::Env env, Napi::Function cb, std::vector<std::uint8_t>* bytes) {
        Napi::Buffer<std::uint8_t> view = Napi::Buffer<std::uint8_t>::Copy(env, bytes->data(), bytes->size());
        delete bytes;
        cb.Call({view});
        if (env.IsExceptionPending()) {
          // A throwing sink must not leave a pending exception at this
          // ThreadSafeFunction drain point: an uncaught exception here is
          // already fatal to the process under current Node's
          // unhandled-exception default. Records are best-effort
          // telemetry with no return path to report a failure on, so the
          // call that produced this one simply completes -- matching the
          // WASM sink's contract (a throw is intercepted, never
          // propagated).
          (void)env.GetAndClearPendingException();
        }
      });
  if (rc != napi_ok) {
    delete payload;
  }
}

// Same contract as the Workbook registration guard: required, an integer in
// int32 range.
bool ReadInt32(const Napi::CallbackInfo& info, std::size_t idx, std::int32_t* out) {
  const Napi::Value value = info[idx];
  if (!value.IsNumber()) {
    Napi::TypeError::New(info.Env(), "positional argument must be a number").ThrowAsJavaScriptException();
    return false;
  }
  const double number = NumberValue(value, 0.0);
  if (!std::isfinite(number) || std::trunc(number) != number ||
      number < static_cast<double>(std::numeric_limits<std::int32_t>::min()) ||
      number > static_cast<double>(std::numeric_limits<std::int32_t>::max())) {
    Napi::RangeError::New(info.Env(), "positional argument is outside int32 range").ThrowAsJavaScriptException();
    return false;
  }
  *out = static_cast<std::int32_t>(number);
  return true;
}

}  // namespace

// Read-only ad-hoc evaluation anchored at `Sheet1!A1`, matching the embind
// and Python surfaces byte for byte. Writing the formula into A1 and
// recalcing instead would make every A1-referencing formula (`=A1`,
// `=COUNTA(A1)`) a self-reference and return `#REF!`, which is neither
// what a one-shot helper should mean nor what the other two surfaces do.
Napi::Value EvalFormula(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (!info[0].IsString()) {
    Napi::TypeError::New(env, "formula must be a string").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  const std::string formula = StringArg(info, 0);

  fm_workbook_t* wb = nullptr;
  fm_status_t rc = fm_workbook_create_empty(&wb);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  std::size_t sheet_index = 0;
  rc = fm_workbook_add_sheet(wb, "Sheet1", &sheet_index);
  if (rc != 0) {
    const Napi::Value failed = ThrowStatus(env, rc);
    fm_workbook_destroy(wb);
    return failed;
  }
  fm_value_t v{};
  rc = fm_sheet_evaluate_formula(wb, 0, 0, 0, formula.c_str(), &v);
  // Translate before the destroy: a text result borrows the workbook's scratch.
  const Napi::Value out = rc == 0 ? ValueToJs(env, v) : ThrowStatus(env, rc);
  fm_workbook_destroy(wb);
  return out;
}

Napi::Value Version(const Napi::CallbackInfo& info) {
  return JsString(info.Env(), fm_get_version());
}

Napi::Value StatusString(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  int32_t code = 0;
  if (!ReadInt32(info, 0, &code)) {
    return env.Undefined();
  }
  const char* s = fm_get_status_name(static_cast<fm_status_t>(code));
  return JsString(env, s);
}

Napi::Value ErrorDisplayName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  int32_t code = 0;
  if (!ReadInt32(info, 0, &code)) {
    return env.Undefined();
  }
  const char* s = fm_get_error_display_name(static_cast<fm_error_code_t>(code));
  return JsString(env, s);
}

Napi::Value SetLogMinLevel(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  int32_t level = 0;
  if (!ReadInt32(info, 0, &level)) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_set_log_min_level(level), env.Undefined());
}

Napi::Value SetLogSink(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const bool clearing = info[0].IsNull() || info[0].IsUndefined();
  Napi::Function sink;
  if (!clearing && !FunctionArg(info, 0, &sink)) {
    Napi::TypeError::New(env, "sink must be a function or null").ThrowAsJavaScriptException();
    return env.Undefined();
  }

  if (clearing) {
    fm_status_t rc = fm_set_log_sink(nullptr, nullptr);
    const std::lock_guard<std::mutex> lock(log_sink_mutex());
    ReleaseLogSinkLocked();
    return ValueOrThrow(env, rc, env.Undefined());
  }

  Napi::ThreadSafeFunction fresh = Napi::ThreadSafeFunction::New(env, sink, "formulon.logSink", 0, 1);
  // A registered sink must not by itself keep the process alive.
  fresh.Unref(env);
  // A refused registration keeps the previous sink, so the swap follows it.
  const fm_status_t rc = fm_set_log_sink(&LogSinkTrampoline, nullptr);
  {
    const std::lock_guard<std::mutex> lock(log_sink_mutex());
    if (rc == 0) {
      ReleaseLogSinkLocked();
      log_sink_tsfn() = fresh;
    } else {
      fresh.Release();
    }
  }
  return ValueOrThrow(env, rc, env.Undefined());
}

Napi::Value SetErrorClass(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Function error_class;
  if (!FunctionArg(info, 0, &error_class)) {
    Napi::TypeError::New(env, "the error class must be a constructor").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  GetAddonData(env).error_class = Napi::Persistent(error_class);
  return env.Undefined();
}

}  // namespace formulon_node
