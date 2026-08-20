//
// JsWorkbook lifecycle, sheet management, recalc / calc-mode plumbing,
// and the iterative-progress callback bridge. The split is by API
// surface area: cells / styles / pivot / etc. live in sibling files
// under `parts/`. The single `EMSCRIPTEN_BINDINGS` block is in
// `parts/bindings_register.cpp`.

#include <emscripten/bind.h>
#include <emscripten/em_js.h>
#include <emscripten/val.h>

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

// The JS half of `call_js_callback`. The returned ordinals are the
// `JsCallbackOutcome` enumerators, spelled as literals because the body is
// a string handed to the JS glue generator rather than compiled C++.
// `Reflect.apply` rather than `fn.apply` so an object that shadows `apply`
// cannot redirect the call, and a non-callable argument surfaces as a
// TypeError this handler classifies like any other throw. The body carries
// no `//` comments: it is stringified onto a single line. `clang-format`
// is off over it for the same reason -- it reads the body as C++ and
// splits JavaScript's `===` into `== =`.
//
// FORMULON-ALLOW: the try / catch below is JavaScript source, not C++. It
// is the only place a JS exception can be intercepted in a binary linked
// -fno-exceptions -sDISABLE_EXCEPTION_CATCHING=1, where an escaping
// exception would tear down C++ frames without running their destructors.
// clang-format off
EM_JS(int32_t, fm_wasm_invoke_js_callback, (emscripten::EM_VAL fn_handle, emscripten::EM_VAL args_handle), {
  var fn = Emval.toValue(fn_handle);
  var args = Emval.toValue(args_handle);
  try {
    var result = Reflect.apply(fn, undefined, args);
    if (result === undefined || result === null) {
      return 2;
    }
    return result ? 1 : 0;
  } catch (e) {
    return -1;
  }
});
// clang-format on

EM_JS_DEPS(fm_wasm_js_callback, "$Emval");

namespace formulon {
namespace wasm {
namespace parts {

JsCallbackOutcome call_js_callback(const emscripten::val& fn, const emscripten::val& args) {
  // Both handles stay owned by the arguments; the JS side only reads them.
  return static_cast<JsCallbackOutcome>(fm_wasm_invoke_js_callback(fn.as_handle(), args.as_handle()));
}

namespace {

/// The envelope a throwing progress callback produces. Built by hand
/// rather than through `error_status`: the engine never saw a failure, so
/// there is no thread-local C-ABI diagnostic to copy out.
JsStatus progress_callback_threw_status() {
  JsStatus s;
  s.ok = false;
  s.status = kBindingCallbackException;
  s.message = "the iterative progress callback threw; the solve was aborted";
  s.context = "Workbook.setIterativeProgress";
  return s;
}

}  // namespace

JsWorkbook::~JsWorkbook() {
  if (handle_ != nullptr) {
    fm_workbook_destroy(handle_);
    handle_ = nullptr;
  }
}

JsWorkbook::JsWorkbook(JsWorkbook&& other) noexcept
    : handle_(other.handle_),
      progress_callback_(std::move(other.progress_callback_)),
      progress_callback_threw_(other.progress_callback_threw_) {
  other.handle_ = nullptr;
  other.progress_callback_ = emscripten::val::null();
  other.progress_callback_threw_ = false;
  // The engine holds the address the callback was registered from.
  rebind_progress_callback();
}

JsWorkbook& JsWorkbook::operator=(JsWorkbook&& other) noexcept {
  if (this != &other) {
    if (handle_ != nullptr) {
      fm_workbook_destroy(handle_);
    }
    handle_ = other.handle_;
    progress_callback_ = std::move(other.progress_callback_);
    progress_callback_threw_ = other.progress_callback_threw_;
    other.handle_ = nullptr;
    other.progress_callback_ = emscripten::val::null();
    other.progress_callback_threw_ = false;
    rebind_progress_callback();
  }
  return *this;
}

JsWorkbook* JsWorkbook::createDefault() {
  auto wb = std::unique_ptr<JsWorkbook>(new JsWorkbook());
  fm_status_t rc = fm_workbook_create(&wb->handle_);
  if (rc != 0) {
    // The handle remains null; the caller can read lastErrorMessage().
    // We still return the object so JS can inspect status via a
    // subsequent isValid() check.
  }
  return wb.release();
}

JsWorkbook* JsWorkbook::createEmpty() {
  auto wb = std::unique_ptr<JsWorkbook>(new JsWorkbook());
  (void)fm_workbook_create_empty(&wb->handle_);
  return wb.release();
}

JsWorkbook* JsWorkbook::loadBytes(emscripten::val bytes) {
  auto wb = std::unique_ptr<JsWorkbook>(new JsWorkbook());
  const std::vector<uint8_t> buf = val_to_bytes(bytes);
  if (buf.empty()) {
    (void)fm_workbook_load(nullptr, 0, &wb->handle_);
    return wb.release();
  }
  (void)fm_workbook_load(buf.data(), buf.size(), &wb->handle_);
  return wb.release();
}

JsSaveResult JsWorkbook::save() const {
  JsSaveResult r;
  if (handle_ == nullptr) {
    r.status = error_status(kBindingInvalidHandle);
    return r;
  }
  uint8_t* out = nullptr;
  std::size_t len = 0;
  fm_status_t rc = fm_workbook_save(handle_, &out, &len);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.bytes = bytes_to_val(out, len);
  fm_buffer_free(out);
  r.status = ok_status();
  return r;
}

JsSaveResult JsWorkbook::saveAs(int32_t format) const {
  JsSaveResult r;
  if (handle_ == nullptr) {
    r.status = error_status(kBindingInvalidHandle);
    return r;
  }
  uint8_t* out = nullptr;
  std::size_t len = 0;
  fm_status_t rc = fm_workbook_save_as(handle_, format, &out, &len);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.bytes = bytes_to_val(out, len);
  fm_buffer_free(out);
  r.status = ok_status();
  return r;
}

JsSaveDiagnosticsResult JsWorkbook::saveWithDiagnostics(int32_t format) const {
  JsSaveDiagnosticsResult r;
  if (handle_ == nullptr) {
    r.status = error_status(kBindingInvalidHandle);
    return r;
  }
  uint8_t* out = nullptr;
  std::size_t len = 0;
  fm_save_diagnostics_t d{};
  fm_status_t rc = fm_workbook_save_with_diagnostics(handle_, format, &out, &len, &d);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.bytes = bytes_to_val(out, len);
  fm_buffer_free(out);
  r.downgradedFormulaCount = d.downgraded_formula_count;
  r.deferredFeatureCount = d.deferred_feature_count;
  r.droppedPartCount = d.dropped_part_count;
  r.droppedRelationshipCount = d.dropped_relationship_count;
  r.renumberedPartCount = d.renumbered_part_count;
  r.status = ok_status();
  return r;
}

JsReadDiagnosticsResult JsWorkbook::readDiagnostics() const {
  JsReadDiagnosticsResult r;
  if (handle_ == nullptr) {
    r.status = error_status(kBindingInvalidHandle);
    return r;
  }
  fm_read_diagnostics_t d{};
  fm_status_t rc = fm_workbook_read_diagnostics(handle_, &d);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.undecodedFormulaCount = d.undecoded_formula_count;
  r.undecodedDefinedNameCount = d.undecoded_defined_name_count;
  r.undecodedPartCount = d.undecoded_part_count;
  r.skippedFeatureCount = d.skipped_feature_count;
  r.unknownContentTypeCount = d.unknown_content_type_count;
  r.status = ok_status();
  return r;
}

// ---- Sheet management ----------------------------------------------------

JsStatus JsWorkbook::addSheet(const std::string& name) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_add_sheet(handle_, name.c_str());
  return status_from_rc(rc);
}

JsStatus JsWorkbook::moveSheet(uint32_t fromIdx, uint32_t toIdx) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_move_sheet(handle_, fromIdx, toIdx);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::removeSheet(uint32_t index) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_remove_sheet(handle_, index);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::renameSheet(uint32_t index, const std::string& newName) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_rename_sheet(handle_, index, newName.c_str());
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setDefinedName(const std::string& name, const std::string& formula) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_set_defined_name(handle_, name.c_str(), formula.c_str());
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setDefinedNameScoped(const std::string& name, const std::string& formula, int32_t localSheetId) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_set_defined_name_scoped(handle_, name.c_str(), formula.c_str(), localSheetId);
  return status_from_rc(rc);
}

// ---- Row / column edits --------------------------------------------------

JsStatus JsWorkbook::invoke_row_col_edit(RowColEditFn fn, uint32_t sheet, uint32_t index, uint32_t count) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  return status_from_rc(fn(handle_, sheet, index, count));
}

JsStatus JsWorkbook::insertRows(uint32_t sheet, uint32_t row, uint32_t count) {
  return invoke_row_col_edit(fm_workbook_insert_rows, sheet, row, count);
}

JsStatus JsWorkbook::deleteRows(uint32_t sheet, uint32_t row, uint32_t count) {
  return invoke_row_col_edit(fm_workbook_delete_rows, sheet, row, count);
}

JsStatus JsWorkbook::insertCols(uint32_t sheet, uint32_t col, uint32_t count) {
  return invoke_row_col_edit(fm_workbook_insert_cols, sheet, col, count);
}

JsStatus JsWorkbook::deleteCols(uint32_t sheet, uint32_t col, uint32_t count) {
  return invoke_row_col_edit(fm_workbook_delete_cols, sheet, col, count);
}

// `JsWorkbook::sheetCount` is now emitted by the binding codegen (see
// `src/wasm/generated/workbook_counts.cpp`).

JsStringResult JsWorkbook::sheetName(uint32_t idx) const {
  JsStringResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  const char* name = nullptr;
  fm_status_t rc = fm_workbook_sheet_name(handle_, idx, &name);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.value = (name != nullptr) ? std::string(name) : std::string();
  r.status = ok_status();
  return r;
}

// ---- Recalc / calc mode --------------------------------------------------

JsStatus JsWorkbook::recalc() {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  progress_callback_threw_ = false;
  fm_status_t rc = fm_workbook_recalc(handle_);
  const bool callback_threw = progress_callback_threw_;
  progress_callback_threw_ = false;
  // An engine failure wins: it carries its own diagnostic and is the more
  // specific answer. Otherwise a throwing progress callback, which the
  // engine only saw as a cancellation, becomes the reported failure.
  if (rc == 0 && callback_threw) {
    return progress_callback_threw_status();
  }
  return status_from_rc(rc);
}

JsParallelRecalcResult JsWorkbook::recalcParallel(emscripten::val threadCount) {
  JsParallelRecalcResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }

  // Do not bind this parameter as uint32_t: embind would otherwise coerce
  // e.g. -0.5, 1.5, NaN, Infinity, null, and a missing argument before this
  // method could apply the public 0..8 integer contract. Reuse the C ABI's
  // invalid-argument path so its status message/context remain canonical.
  const emscripten::val number = emscripten::val::global("Number");
  const bool is_finite = number.call<bool>("isFinite", threadCount);
  const bool is_integer = is_finite && number.call<bool>("isInteger", threadCount);
  const double requested = is_integer ? threadCount.as<double>() : -1.0;
  const bool is_valid = is_integer && requested >= 0.0 && requested <= 8.0;

  fm_parallel_recalc_stats stats{};
  const uint32_t native_thread_count = is_valid ? static_cast<uint32_t>(requested) : 9U;
  progress_callback_threw_ = false;
  const fm_status_t rc = fm_workbook_recalc_parallel(handle_, native_thread_count, &stats);
  const bool callback_threw = progress_callback_threw_;
  progress_callback_threw_ = false;
  if (rc != 0) {
    // The C ABI guarantees zeroed stats on every failure. Keep the result's
    // default-zero payload as the binding-level equivalent of that contract.
    r.status = error_status(rc);
    return r;
  }
  if (callback_threw) {
    // Same shape as an engine failure: the pass did not finish, so the
    // counters are not reported.
    r.status = progress_callback_threw_status();
    return r;
  }

  // Keep these counters as JS numbers. Values through 2^53-1 are represented
  // exactly by a double, and the public TypeScript contract documents that
  // precision boundary rather than leaking a runtime-dependent BigInt.
  r.stats.cellsEvaluated = static_cast<double>(stats.cells_evaluated);
  r.stats.sccsProcessed = static_cast<double>(stats.sccs_processed);
  r.stats.parallelSteps = static_cast<double>(stats.parallel_steps);
  r.stats.serialFallbackSteps = static_cast<double>(stats.serial_fallback_steps);
  r.stats.cycleRecoveries = static_cast<double>(stats.cycle_recoveries);
  r.stats.workerThreadsStarted = stats.worker_threads_started;
  r.stats.workerThreadsUsed = stats.worker_threads_used;
  r.status = ok_status();
  return r;
}

JsStatus JsWorkbook::setIterative(bool enabled, uint32_t max_iterations, double max_change) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc =
      fm_workbook_set_iterative(handle_, enabled ? 1 : 0, static_cast<int32_t>(max_iterations), max_change);
  return status_from_rc(rc);
}

uint32_t JsWorkbook::calcMode() const {
  if (handle_ == nullptr) {
    return static_cast<uint32_t>(FM_CALC_MODE_AUTO);
  }
  fm_calc_mode_t mode = FM_CALC_MODE_AUTO;
  fm_workbook_calc_mode(handle_, &mode);
  return static_cast<uint32_t>(mode);
}

JsStatus JsWorkbook::setCalcMode(uint32_t mode) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_set_calc_mode(handle_, static_cast<std::int32_t>(mode));
  return status_from_rc(rc);
}

emscripten::val JsWorkbook::pinnedNow() const {
  if (handle_ == nullptr) {
    return emscripten::val::null();
  }
  fm_civil_time_t now{};
  std::int32_t pinned = 0;
  if (fm_workbook_pinned_now(handle_, &now, &pinned) != 0 || pinned == 0) {
    return emscripten::val::null();
  }
  emscripten::val out = emscripten::val::object();
  out.set("year", now.year);
  out.set("month", now.month);
  out.set("day", now.day);
  out.set("hour", now.hour);
  out.set("minute", now.minute);
  out.set("second", now.second);
  return out;
}

JsStatus JsWorkbook::setPinnedNow(uint32_t year, uint32_t month, uint32_t day, uint32_t hour, uint32_t minute,
                                  uint32_t second) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  // Each field is widened rather than range-checked here: the C layer owns
  // the calendar domain, and a JS-side duplicate would be a second place to
  // keep in step with it.
  const fm_civil_time_t now{static_cast<std::int32_t>(year),   static_cast<std::int32_t>(month),
                            static_cast<std::int32_t>(day),    static_cast<std::int32_t>(hour),
                            static_cast<std::int32_t>(minute), static_cast<std::int32_t>(second)};
  return status_from_rc(fm_workbook_set_pinned_now(handle_, &now));
}

JsStatus JsWorkbook::clearPinnedNow() {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  return status_from_rc(fm_workbook_clear_pinned_now(handle_));
}

std::string JsWorkbook::excelProfileId() const {
  if (handle_ == nullptr) {
    return "win-365-ja_JP";
  }
  const char* id = nullptr;
  fm_workbook_excel_profile_id(handle_, &id);
  return id == nullptr ? std::string("win-365-ja_JP") : std::string(id);
}

JsStatus JsWorkbook::setExcelProfileId(const std::string& profile_id) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_set_excel_profile_id(handle_, profile_id.c_str());
  return status_from_rc(rc);
}

emscripten::val JsWorkbook::partialRecalc(emscripten::val viewport) {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    o.set("recomputed", static_cast<uint32_t>(0));
    return o;
  }
  fm_viewport vp{};
  vp.sheet = viewport["sheet"].as<uint32_t>();
  vp.first_row = viewport["firstRow"].as<uint32_t>();
  vp.last_row = viewport["lastRow"].as<uint32_t>();
  vp.first_col = viewport["firstCol"].as<uint32_t>();
  vp.last_col = viewport["lastCol"].as<uint32_t>();
  uint32_t recomputed = 0;
  progress_callback_threw_ = false;
  fm_status_t rc = fm_workbook_partial_recalc(handle_, &vp, &recomputed);
  const bool callback_threw = progress_callback_threw_;
  progress_callback_threw_ = false;
  if (rc != 0) {
    o.set("status", error_status(rc));
    o.set("recomputed", static_cast<uint32_t>(0));
    return o;
  }
  if (callback_threw) {
    o.set("status", progress_callback_threw_status());
    o.set("recomputed", static_cast<uint32_t>(0));
    return o;
  }
  o.set("status", ok_status());
  o.set("recomputed", recomputed);
  return o;
}

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
  const JsCallbackOutcome outcome = call_js_callback(self->progress_callback_, args);
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

JsStatus JsWorkbook::setIterativeProgress(emscripten::val cb) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  if (cb.isNull() || cb.isUndefined()) {
    progress_callback_ = emscripten::val::null();
    fm_status_t rc = fm_workbook_set_iterative_progress(handle_, nullptr, nullptr);
    return status_from_rc(rc);
  }
  progress_callback_ = cb;
  fm_status_t rc = fm_workbook_set_iterative_progress(handle_, &JsWorkbook::iterativeProgressTrampoline, this);
  return status_from_rc(rc);
}

// ---- Free helpers ------------------------------------------------------

JsEvalResult eval_formula(const std::string& formula) {
  JsEvalResult r;
  fm_workbook_t* wb = nullptr;
  fm_status_t rc = fm_workbook_create_empty(&wb);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  rc = fm_workbook_add_sheet(wb, "Sheet1");
  if (rc != 0) {
    r.status = error_status(rc);
    fm_workbook_destroy(wb);
    return r;
  }
  fm_value_t v{};
  rc = fm_workbook_evaluate_formula(wb, 0, 0, 0, formula.c_str(), &v);
  if (rc != 0) {
    r.status = error_status(rc);
    fm_workbook_destroy(wb);
    return r;
  }
  r.value = translate_value(v);
  r.status = ok_status();
  fm_workbook_destroy(wb);
  return r;
}

std::string version_string() {
  const char* s = fm_version_string();
  return s != nullptr ? std::string(s) : std::string();
}

std::string status_string(int32_t status) {
  const char* s = fm_status_string(static_cast<fm_status_t>(status));
  return s != nullptr ? std::string(s) : std::string();
}

std::string error_display_name(int32_t error_code) {
  const char* s = fm_error_display_name(static_cast<fm_error_code_t>(error_code));
  return s != nullptr ? std::string(s) : std::string();
}

std::string last_error_message() {
  const char* s = fm_last_error_message();
  return s != nullptr ? std::string(s) : std::string();
}

std::string last_error_context() {
  const char* s = fm_last_error_context();
  return s != nullptr ? std::string(s) : std::string();
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
  // The record is a length-delimited byte range, never NUL-terminated;
  // hand the JS side a view over exactly `len` bytes.
  emscripten::val args = emscripten::val::array();
  args.set(0, emscripten::val(emscripten::typed_memory_view(len, reinterpret_cast<const std::uint8_t*>(record))));
  // A throwing sink is dropped. Records are emitted from arbitrary depth
  // inside the engine and a log write has no return path to report on, so
  // the only contract worth keeping is that it cannot damage the caller.
  (void)call_js_callback(sink, args);
}

}  // namespace

JsStatus set_log_min_level(int32_t level) {
  return status_from_rc(fm_set_log_min_level(level));
}

JsStatus set_log_sink(emscripten::val sink) {
  if (sink.isNull() || sink.isUndefined()) {
    js_log_sink() = emscripten::val::null();
    return status_from_rc(fm_set_log_sink(nullptr, nullptr));
  }
  js_log_sink() = sink;
  return status_from_rc(fm_set_log_sink(&log_sink_trampoline, nullptr));
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
