// Implementation of the module-level free functions exported alongside
// the `Workbook` class. See `free_funcs.h` for the contract.

#include "node_addon/parts/free_funcs.h"

#include <cstddef>
#include <cstdint>
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
      });
  if (rc != napi_ok) {
    delete payload;
  }
}

}  // namespace

// Read-only ad-hoc evaluation anchored at `Sheet1!A1`, matching the embind
// and Python surfaces byte for byte. Writing the formula into A1 and
// recalcing instead would make every A1-referencing formula (`=A1`,
// `=COUNTA(A1)`) a self-reference and return `#REF!`, which is neither
// what a one-shot helper should mean nor what the other two surfaces do.
Napi::Value EvalFormula(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  std::string formula;
  if (info.Length() > 0) {
    formula = info[0].ToString().Utf8Value();
  }

  fm_workbook_t* wb = nullptr;
  fm_status_t rc = fm_workbook_create_empty(&wb);
  if (rc != 0) {
    return MakeEmptyValueResult(env, MakeErrorStatus(env, rc));
  }
  rc = fm_workbook_add_sheet(wb, "Sheet1");
  if (rc != 0) {
    Napi::Object out = MakeEmptyValueResult(env, MakeErrorStatus(env, rc));
    fm_workbook_destroy(wb);
    return out;
  }
  fm_value_t v{};
  rc = fm_workbook_evaluate_formula(wb, 0, 0, 0, formula.c_str(), &v);
  if (rc != 0) {
    Napi::Object out = MakeEmptyValueResult(env, MakeErrorStatus(env, rc));
    fm_workbook_destroy(wb);
    return out;
  }
  Napi::Object out = MakeValueResult(env, MakeOkStatus(env), v);
  fm_workbook_destroy(wb);
  return out;
}

Napi::Value Version(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const char* s = fm_version_string();
  return Napi::String::New(env, s != nullptr ? s : "");
}

Napi::Value LastErrorMessage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const char* s = fm_last_error_message();
  return Napi::String::New(env, s != nullptr ? s : "");
}

Napi::Value LastErrorContext(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const char* s = fm_last_error_context();
  return Napi::String::New(env, s != nullptr ? s : "");
}

Napi::Value StatusString(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  int32_t code = 0;
  if (info.Length() > 0) {
    code = info[0].ToNumber().Int32Value();
  }
  const char* s = fm_status_string(static_cast<fm_status_t>(code));
  return Napi::String::New(env, s != nullptr ? s : "");
}

Napi::Value ErrorDisplayName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  int32_t code = 0;
  if (info.Length() > 0) {
    code = info[0].ToNumber().Int32Value();
  }
  const char* s = fm_error_display_name(static_cast<fm_error_code_t>(code));
  return Napi::String::New(env, s != nullptr ? s : "");
}

Napi::Value SetLogMinLevel(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  int32_t level = 0;
  if (info.Length() > 0) {
    level = info[0].ToNumber().Int32Value();
  }
  return MakeStatus(env, fm_set_log_min_level(level));
}

Napi::Value SetLogSink(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const bool clearing = info.Length() == 0 || info[0].IsNull() || info[0].IsUndefined();
  if (!clearing && !info[0].IsFunction()) {
    return MakeBindingArgumentError(env, "setLogSink: `sink` must be a function or null");
  }

  if (clearing) {
    fm_status_t rc = fm_set_log_sink(nullptr, nullptr);
    const std::lock_guard<std::mutex> lock(log_sink_mutex());
    ReleaseLogSinkLocked();
    return MakeStatus(env, rc);
  }

  Napi::ThreadSafeFunction fresh =
      Napi::ThreadSafeFunction::New(env, info[0].As<Napi::Function>(), "formulon.logSink", 0, 1);
  // A registered sink must not by itself keep the process alive.
  fresh.Unref(env);
  {
    const std::lock_guard<std::mutex> lock(log_sink_mutex());
    ReleaseLogSinkLocked();
    log_sink_tsfn() = fresh;
  }
  return MakeStatus(env, fm_set_log_sink(&LogSinkTrampoline, nullptr));
}

}  // namespace formulon_node
