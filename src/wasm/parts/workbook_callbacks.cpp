//
// JS callbacks the engine calls back into: the per-workbook iterative-progress
// bridge and the process-wide structured-log sink.

#include <emscripten/em_js.h>
#include <emscripten/val.h>

#include <cstddef>
#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

// Publishes which workbooks are inside their iterative progress callback, by
// native address (the embind handle's `$$.ptr`). The pre-js method wrapper
// refuses every call on such a workbook, as the Node addon does.
// clang-format off
EM_JS(void, fm_wasm_mark_progress_callback, (uintptr_t workbook, int32_t active), {
  var key = Symbol.for('formulon.progressCallbackWorkbooks');
  var busy = Module[key];
  if (!busy) {
    busy = new Set();
    Module[key] = busy;
  }
  if (active) {
    busy.add(workbook);
  } else {
    busy.delete(workbook);
  }
});
// clang-format on

namespace formulon {
namespace wasm {
namespace parts {

namespace {

/// `typeof value === 'function'`, which never runs user code.
bool is_js_function(const emscripten::val& value) {
  return value.typeOf().as<std::string>() == "function";
}

}  // namespace

// ---- Iterative-progress callback bridge --------------------------------
//
// The callback belongs to the wrapper, not to the module: the C ABI
// carries a `user_data` pointer and this binding passes `this` through it,
// so two workbooks driven from one page keep their own callbacks. The Node
// addon has always worked this way, and `packages/npm-native/README.md`
// claims the two surfaces behave alike for every shared method.

int32_t JsWorkbook::iterativeProgressTrampoline(uint32_t iteration, double max_residual, uint32_t max_iterations,
                                                void* user_data) {
  auto* const self = static_cast<JsWorkbook*>(user_data);
  if (self == nullptr || self->progress_callback_.isNull() || self->progress_callback_.isUndefined()) {
    return 1;
  }
  emscripten::val args = emscripten::val::array();
  args.set(0, iteration);
  args.set(1, max_residual);
  args.set(2, max_iterations);
  self->in_iterative_progress_callback_ = true;
  fm_wasm_mark_progress_callback(reinterpret_cast<uintptr_t>(self), 1);
  const JsCallbackOutcome outcome = call_js_callback(self->progress_callback_, args);
  fm_wasm_mark_progress_callback(reinterpret_cast<uintptr_t>(self), 0);
  self->in_iterative_progress_callback_ = false;
  if (outcome == JsCallbackOutcome::kThrew) {
    // Abort the solve so every frame between here and the recalc entry
    // point unwinds the ordinary way, and leave the reason behind for it.
    self->progress_callback_threw_ = true;
    return 0;
  }
  // Returning nothing means "keep going"; anything else is read the way
  // JavaScript itself would read it.
  return outcome == JsCallbackOutcome::kFalsy ? 0 : 1;
}

void JsWorkbook::rebind_progress_callback() {
  if (handle_ == nullptr || progress_callback_.isNull() || progress_callback_.isUndefined()) {
    return;
  }
  (void)fm_workbook_set_iterative_progress(handle_, &JsWorkbook::iterativeProgressTrampoline, this);
}

void JsWorkbook::setIterativeProgress(emscripten::val cb) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  if (cb.isNull() || cb.isUndefined()) {
    progress_callback_ = emscripten::val::null();
    succeeded(fm_workbook_set_iterative_progress(handle_, nullptr, nullptr));
    return;
  }
  // Rejected at install, as on the Node surface; the installed callback stays.
  if (!is_js_function(cb)) {
    record_type_error("Workbook.setIterativeProgress callback must be a function or null");
    return;
  }
  progress_callback_ = cb;
  succeeded(fm_workbook_set_iterative_progress(handle_, &JsWorkbook::iterativeProgressTrampoline, this));
}

// ---- Structured logging (process-wide, not per handle) --------------------

namespace {

emscripten::val& js_log_sink() {
  static emscripten::val sink = emscripten::val::null();
  return sink;
}

void log_sink_trampoline(const char* record, std::size_t len, void* /*user_data*/) {
  emscripten::val& sink = js_log_sink();
  if (sink.isNull() || sink.isUndefined()) {
    return;
  }
  // The record is a length-delimited byte range, never NUL-terminated. A
  // `bytes_to_val` copy, not a live view into the module's own memory: the
  // `/threads` build backs that memory with a `SharedArrayBuffer`, and a
  // view over it fails `formulon.d.ts`'s other documented read path --
  // most browsers' `TextDecoder#decode` rejects such a backing buffer.
  emscripten::val args = emscripten::val::array();
  args.set(0, bytes_to_val(reinterpret_cast<const std::uint8_t*>(record), len));
  // A log write has no return path to report on, so the engine carries on;
  // a throwing sink is rethrown by the enclosing call once it returns.
  (void)call_js_callback(sink, args);
}

}  // namespace

void set_log_min_level(int32_t level) {
  succeeded(fm_set_log_min_level(level));
}

void set_log_sink(emscripten::val sink) {
  if (sink.isNull() || sink.isUndefined()) {
    js_log_sink() = emscripten::val::null();
    succeeded(fm_set_log_sink(nullptr, nullptr));
    return;
  }
  if (!is_js_function(sink)) {
    record_type_error("Formulon.setLogSink sink must be a function");
    return;
  }
  js_log_sink() = sink;
  succeeded(fm_set_log_sink(&log_sink_trampoline, nullptr));
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
