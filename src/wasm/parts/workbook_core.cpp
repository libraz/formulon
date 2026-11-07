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
// TypeError this handler classifies like any other throw. A throw is also
// recorded (kind 2) in the enclosing wrapped call's slot, displacing a status
// failure but not an argument error, so that call rethrows the same object.
// The body carries no `//` comments: it is stringified onto a single line.
// `clang-format` is off over it for the same reason -- it reads the body as
// C++ and splits JavaScript's `===` into `== =`.
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
    var errors = Module[Symbol.for('formulon.argumentErrors')];
    if (errors && errors.depth > 0 && (errors.pending === undefined || errors.pending.kind === 3)) {
      errors.pending = { kind: 2, message: '', thrown: { value: e } };
    }
    return -1;
  }
});
// clang-format on

EM_JS_DEPS(fm_wasm_js_callback, "$Emval");

// Aborts the in-flight `delete()` call with a JS-visible error instead of
// letting `~JsWorkbook()` return normally into `fm_workbook_destroy`. The
// throw unwinds back through embind's `delete()` (which never reaches the
// lines that null out its pointer) and, if the caller is the progress
// callback itself, on into the `try`/`catch` in
// `fm_wasm_invoke_js_callback` above -- the same path an ordinarily
// throwing callback takes. Error and message match the pre-js lifecycle
// guard and the Node addon's `Workbook::Dispose`, so every path refuses alike.
//
// FORMULON-ALLOW: JavaScript source again, for the same reason as the
// callback trampoline above -- it is the one place a JS-visible exception
// can be raised from a binary built with -fno-exceptions.
// clang-format off
EM_JS(void, fm_wasm_throw_dispose_during_callback, (), {
  var message = 'cannot dispose Workbook during an active method call (kGraphRecalcReentrant, 4005)';
  if (typeof Module.FormulonError !== 'function') throw new TypeError(message);
  throw new Module.FormulonError(4005, message, '');
});
// clang-format on

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

JsCallbackOutcome call_js_callback(const emscripten::val& fn, const emscripten::val& args) {
  // Both handles stay owned by the arguments; the JS side only reads them.
  return static_cast<JsCallbackOutcome>(fm_wasm_invoke_js_callback(fn.as_handle(), args.as_handle()));
}

namespace {

/// Records the failure of a solve a throwing progress callback aborted. The
/// callback's own exception, recorded by the trampoline, takes precedence;
/// this stands only when that record could not be made.
void record_progress_callback_threw() {
  record_binding_failure(kBindingCallbackException, "the iterative progress callback threw; the solve was aborted");
}

/// `typeof value === 'function'`, which never runs user code.
bool is_js_function(const emscripten::val& value) {
  return value.typeOf().as<std::string>() == "function";
}

}  // namespace

JsWorkbook::~JsWorkbook() {
  if (in_iterative_progress_callback_) {
    // Never reached past this call: it throws into the JS caller of
    // `delete()`, which is `progress_callback_` itself here.
    fm_wasm_throw_dispose_during_callback();
  }
  if (handle_ != nullptr) {
    fm_workbook_destroy(handle_);
    handle_ = nullptr;
  }
}

void JsWorkbook::dispose() {
  if (in_iterative_progress_callback_) {
    // Throws into the JS caller, which is `progress_callback_` itself here.
    fm_wasm_throw_dispose_during_callback();
  }
  if (handle_ != nullptr) {
    fm_workbook_destroy(handle_);
    handle_ = nullptr;
  }
  progress_callback_ = emscripten::val::null();
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

namespace {

// Hands a factory's workbook to JS, which owns it from here on; `undefined`
// after recording the failure.
emscripten::val factory_result(fm_status_t rc, std::unique_ptr<JsWorkbook> wb) {
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  return emscripten::val(wb.release(), emscripten::allow_raw_pointers());
}

}  // namespace

emscripten::val JsWorkbook::createDefault() {
  auto wb = std::unique_ptr<JsWorkbook>(new JsWorkbook());
  const fm_status_t rc = fm_workbook_create_default(&wb->handle_);
  return factory_result(rc, std::move(wb));
}

emscripten::val JsWorkbook::createEmpty() {
  auto wb = std::unique_ptr<JsWorkbook>(new JsWorkbook());
  const fm_status_t rc = fm_workbook_create_empty(&wb->handle_);
  return factory_result(rc, std::move(wb));
}

emscripten::val JsWorkbook::createForProfile(const std::string& profileId) {
  auto wb = std::unique_ptr<JsWorkbook>(new JsWorkbook());
  const fm_status_t rc = fm_workbook_create_for_profile(profileId.c_str(), &wb->handle_);
  return factory_result(rc, std::move(wb));
}

emscripten::val JsWorkbook::loadBytes(emscripten::val bytes) {
  std::vector<uint8_t> data;
  if (!val_to_bytes_checked(bytes, &data)) {
    return emscripten::val::undefined();
  }
  auto wb = std::unique_ptr<JsWorkbook>(new JsWorkbook());
  // An empty array still reaches C, which rejects it with its own diagnostic.
  static const uint8_t kEmpty = 0;
  const fm_status_t rc = fm_workbook_load(data.empty() ? &kEmpty : data.data(), data.size(), &wb->handle_);
  return factory_result(rc, std::move(wb));
}

namespace {

// The `fm_workbook_format_t` ordinal a save's optional `format` argument
// names: xlsx when absent. A value past int32 names no format, so it maps to
// the unknown ordinal for the C layer to reject.
int32_t save_format(const emscripten::val& format) {
  if (!js_value_present(format)) {
    return FM_WORKBOOK_FORMAT_XLSX;
  }
  const double value = format.as<double>();
  return value >= 0.0 && value <= 2147483647.0 ? static_cast<int32_t>(value) : FM_WORKBOOK_FORMAT_UNKNOWN;
}

}  // namespace

emscripten::val JsWorkbook::save(emscripten::val format) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  uint8_t* out = nullptr;
  std::size_t len = 0;
  const fm_status_t rc = fm_workbook_save_as(handle_, save_format(format), &out, &len);
  return snapshot_bytes(rc, out, len);
}

JsSaveDiagnostics JsWorkbook::saveWithDiagnostics(emscripten::val format) const {
  JsSaveDiagnostics r;
  if (handle_ == nullptr) {
    record_invalid_handle();
    return r;
  }
  uint8_t* out = nullptr;
  std::size_t len = 0;
  fm_save_diagnostics_t d = FM_STRUCT_INIT(fm_save_diagnostics_t);
  const fm_status_t rc = fm_workbook_save_with_diagnostics(handle_, save_format(format), &out, &len, &d);
  r.bytes = snapshot_bytes(rc, out, len);
  if (rc != 0) {
    return r;
  }
  r.downgradedFormulaCount = d.downgraded_formula_count;
  r.deferredFeatureCount = d.deferred_feature_count;
  r.droppedPartCount = d.dropped_part_count;
  r.droppedRelationshipCount = d.dropped_relationship_count;
  r.renumberedPartCount = d.renumbered_part_count;
  return r;
}

JsReadDiagnostics JsWorkbook::readDiagnostics() const {
  JsReadDiagnostics r;
  if (handle_ == nullptr) {
    record_invalid_handle();
    return r;
  }
  fm_read_diagnostics_t d = FM_STRUCT_INIT(fm_read_diagnostics_t);
  if (!succeeded(fm_workbook_get_diagnostics(handle_, &d))) {
    return r;
  }
  r.undecodedFormulaCount = d.undecoded_formula_count;
  r.undecodedDefinedNameCount = d.undecoded_defined_name_count;
  r.undecodedPartCount = d.undecoded_part_count;
  r.skippedFeatureCount = d.skipped_feature_count;
  r.unknownContentTypeCount = d.unknown_content_type_count;
  return r;
}

// ---- Sheet management ----------------------------------------------------

uint32_t JsWorkbook::addSheet(const std::string& name) {
  std::size_t out = 0;
  if (!succeeded(handle_ != nullptr ? fm_workbook_add_sheet(handle_, name.c_str(), &out) : kBindingInvalidHandle)) {
    return 0U;
  }
  return static_cast<uint32_t>(out);
}

void JsWorkbook::moveSheet(uint32_t fromIndex, uint32_t toIndex) {
  succeeded(handle_ != nullptr ? fm_workbook_move_sheet(handle_, fromIndex, toIndex) : kBindingInvalidHandle);
}

uint32_t JsWorkbook::copySheet(uint32_t sourceIndex, uint32_t insertBefore, emscripten::val newName) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0U;
  }
  std::string name;
  const bool has_name = js_value_present(newName);
  if (has_name) {
    name = newName.as<std::string>();
  }
  std::size_t out = 0;
  if (!succeeded(fm_workbook_copy_sheet(handle_, sourceIndex, insertBefore, has_name ? name.c_str() : nullptr, &out))) {
    return 0U;
  }
  return static_cast<uint32_t>(out);
}

uint32_t JsWorkbook::activeSheet() const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0U;
  }
  std::size_t out = 0;
  if (!succeeded(fm_workbook_get_active_sheet(handle_, &out))) {
    return 0U;
  }
  return static_cast<uint32_t>(out);
}

void JsWorkbook::setActiveSheet(uint32_t sheet) {
  succeeded(handle_ != nullptr ? fm_workbook_set_active_sheet(handle_, sheet) : kBindingInvalidHandle);
}

void JsWorkbook::removeSheet(uint32_t index) {
  succeeded(handle_ != nullptr ? fm_workbook_remove_sheet(handle_, index) : kBindingInvalidHandle);
}

void JsWorkbook::renameSheet(uint32_t index, const std::string& newName) {
  succeeded(handle_ != nullptr ? fm_workbook_rename_sheet(handle_, index, newName.c_str()) : kBindingInvalidHandle);
}

void JsWorkbook::setDefinedName(emscripten::val input) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Workbook.setDefinedName");
  if (!reader.require_object(input, "definedName")) {
    return;
  }
  const std::string name = reader.required_string(input, "name", "name");
  const std::string formula = reader.required_string(input, "formula", "formula");
  const emscripten::val scope = reader.value(input, "localSheetId", "localSheetId");
  int32_t local_sheet_id = -1;
  if (js_value_present(scope)) {
    const uint32_t sheet = reader.u32_value(scope, 0U, "localSheetId");
    if (reader.ok() && sheet > 2147483647U) {
      reader.fail_range("localSheetId", "[0, 2147483647]");
    }
    local_sheet_id = static_cast<int32_t>(sheet);
  }
  std::string comment_store;
  const char* comment = reader.optional_string(input, "comment", comment_store, "comment");
  const bool hidden = reader.boolean(input, "hidden", false, "hidden");
  if (!reader.ok()) {
    return;
  }
  fm_defined_name_input_t in = FM_STRUCT_INIT(fm_defined_name_input_t);
  in.name = name.c_str();
  in.formula = formula.c_str();
  in.local_sheet_id = local_sheet_id;
  in.comment = comment;
  in.hidden = hidden ? 1 : 0;
  succeeded(fm_workbook_set_defined_name(handle_, &in));
}

// ---- Row / column edits --------------------------------------------------

void JsWorkbook::invoke_row_col_edit(RowColEditFn fn, uint32_t sheet, uint32_t index, uint32_t count) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  succeeded(fn(handle_, sheet, index, count));
}

void JsWorkbook::insertRows(uint32_t sheet, uint32_t row, uint32_t count) {
  invoke_row_col_edit(fm_sheet_insert_rows, sheet, row, count);
}

void JsWorkbook::deleteRows(uint32_t sheet, uint32_t row, uint32_t count) {
  invoke_row_col_edit(fm_sheet_delete_rows, sheet, row, count);
}

void JsWorkbook::insertColumns(uint32_t sheet, uint32_t col, uint32_t count) {
  invoke_row_col_edit(fm_sheet_insert_columns, sheet, col, count);
}

void JsWorkbook::deleteColumns(uint32_t sheet, uint32_t col, uint32_t count) {
  invoke_row_col_edit(fm_sheet_delete_columns, sheet, col, count);
}

// `JsWorkbook::sheetCount` is now emitted by the binding codegen (see
// `src/wasm/generated/workbook_counts.cpp`).

std::string JsWorkbook::sheetName(uint32_t idx) const {
  const char* name = nullptr;
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_sheet_name(handle_, idx, &name) : kBindingInvalidHandle)) {
    return std::string();
  }
  return string_from_cstr(name);
}

// ---- Recalc / calc mode --------------------------------------------------

void JsWorkbook::recalc() {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  progress_callback_threw_ = false;
  fm_status_t rc = fm_workbook_recalc(handle_);
  const bool callback_threw = progress_callback_threw_;
  progress_callback_threw_ = false;
  // An engine failure wins: it carries its own diagnostic and is the more
  // specific answer. Otherwise a throwing progress callback, which the
  // engine only saw as a cancellation, becomes the reported failure.
  if (rc == 0 && callback_threw) {
    record_progress_callback_threw();
    return;
  }
  succeeded(rc);
}

JsParallelRecalcStats JsWorkbook::recalcParallel(emscripten::val threadCount) {
  JsParallelRecalcStats r;
  if (handle_ == nullptr) {
    record_invalid_handle();
    return r;
  }

  // Do not bind this parameter as uint32_t: embind would otherwise coerce
  // e.g. -0.5, 1.5, NaN, Infinity, null, and a missing argument before this
  // method could apply the public 0..8 integer contract.
  // Same contract as the Node guard: a non-number is a TypeError, a number
  // that is not a uint32 is a RangeError, and the 0..8 limit is the C ABI's.
  if (!threadCount.isNumber()) {
    record_type_error("positional argument must be a number");
    return r;
  }
  const emscripten::val number = emscripten::val::global("Number");
  const double requested = threadCount.as<double>();
  if (!number.call<bool>("isInteger", threadCount) || requested < 0.0 || requested > 4294967295.0) {
    record_range_error("positional argument is outside uint32 range");
    return r;
  }

  fm_parallel_recalc_stats_t stats = FM_STRUCT_INIT(fm_parallel_recalc_stats_t);
  const uint32_t native_thread_count = static_cast<uint32_t>(requested);
  progress_callback_threw_ = false;
  const fm_status_t rc = fm_workbook_recalc_parallel(handle_, native_thread_count, &stats);
  const bool callback_threw = progress_callback_threw_;
  progress_callback_threw_ = false;
  if (!succeeded(rc)) {
    return r;
  }
  if (callback_threw) {
    // Same shape as an engine failure: the pass did not finish, so the
    // counters are not reported.
    record_progress_callback_threw();
    return r;
  }

  // Keep these counters as JS numbers. Values through 2^53-1 are represented
  // exactly by a double, and the public TypeScript contract documents that
  // precision boundary rather than leaking a runtime-dependent BigInt.
  r.cellsEvaluated = static_cast<double>(stats.cells_evaluated);
  r.sccsProcessed = static_cast<double>(stats.sccs_processed);
  r.parallelSteps = static_cast<double>(stats.parallel_steps);
  r.serialFallbackSteps = static_cast<double>(stats.serial_fallback_steps);
  r.cycleRecoveries = static_cast<double>(stats.cycle_recoveries);
  r.workerThreadsStarted = stats.worker_threads_started;
  r.workerThreadsUsed = stats.worker_threads_used;
  return r;
}

bool JsWorkbook::hasDirtyCells() const {
  std::int32_t value = 0;
  return succeeded(handle_ != nullptr ? fm_workbook_has_dirty_cells(handle_, &value) : kBindingInvalidHandle) &&
         value != 0;
}

void JsWorkbook::setIterative(emscripten::val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  // Every field is non-nullable: an omitted or `undefined` key keeps the stored
  // value, `null` and a non-integer cap are type errors.
  JsNarrowNumericReader reader("Workbook.setIterative");
  if (!reader.require_object(patch, "patch")) {
    return;
  }
  fm_iterative_settings_patch_t p = FM_STRUCT_INIT(fm_iterative_settings_patch_t);
  p.enabled = -1;
  p.max_iterations = -1;
  p.max_change = -1.0;
  const emscripten::val enabled = reader.value(patch, "enabled", "enabled");
  if (reader.ok() && !enabled.isUndefined()) {
    if (enabled.isNull() || !(enabled.isTrue() || enabled.isFalse())) {
      reader.fail_type("enabled", "a boolean");
    }
    p.enabled = enabled.isTrue() ? 1 : 0;
  }
  const emscripten::val cap = reader.value(patch, "maxIterations", "maxIterations");
  if (reader.ok() && !cap.isUndefined()) {
    const emscripten::val number = emscripten::val::global("Number");
    if (!cap.isNumber() || !number.call<bool>("isInteger", cap)) {
      reader.fail_type("maxIterations", "an integer Number");
    } else {
      const double requested = cap.as<double>();
      if (requested < 0.0 || requested > 2147483647.0) {
        reader.fail_range("maxIterations", "[0, 2147483647]");
      }
      p.max_iterations = static_cast<int32_t>(requested);
    }
  }
  const emscripten::val change = reader.value(patch, "maxChange", "maxChange");
  if (reader.ok() && !change.isUndefined()) {
    if (!change.isNumber()) {
      reader.fail_type("maxChange", "a Number");
    } else {
      p.max_change = change.as<double>();
    }
  }
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_workbook_set_iterative(handle_, &p));
}

emscripten::val JsWorkbook::getIterative() const {
  int32_t enabled = 0;
  uint32_t max_iterations = 0;
  double max_change = 0.0;
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_iterative(handle_, &enabled, &max_iterations, &max_change)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  o.set("enabled", enabled != 0);
  o.set("maxIterations", max_iterations);
  o.set("maxChange", max_change);
  return o;
}

double JsWorkbook::calcMode() const {
  std::int32_t mode = FM_CALC_MODE_AUTO;
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_calc_mode(handle_, &mode) : kBindingInvalidHandle)) {
    return 0.0;
  }
  return static_cast<double>(mode);
}

void JsWorkbook::setCalcMode(uint32_t mode) {
  succeeded(handle_ != nullptr ? fm_workbook_set_calc_mode(handle_, static_cast<std::int32_t>(mode))
                               : kBindingInvalidHandle);
}

bool JsWorkbook::fullPrecision() const {
  std::int32_t value = 1;
  return succeeded(handle_ != nullptr ? fm_workbook_get_full_precision(handle_, &value) : kBindingInvalidHandle) &&
         value != 0;
}

void JsWorkbook::setFullPrecision(bool fullPrecision) {
  succeeded(handle_ != nullptr ? fm_workbook_set_full_precision(handle_, fullPrecision ? 1 : 0)
                               : kBindingInvalidHandle);
}

bool JsWorkbook::calcOnSave() const {
  std::int32_t value = 1;
  return succeeded(handle_ != nullptr ? fm_workbook_get_calc_on_save(handle_, &value) : kBindingInvalidHandle) &&
         value != 0;
}

void JsWorkbook::setCalcOnSave(bool calcOnSave) {
  succeeded(handle_ != nullptr ? fm_workbook_set_calc_on_save(handle_, calcOnSave ? 1 : 0) : kBindingInvalidHandle);
}

double JsWorkbook::refMode() const {
  std::int32_t mode = FM_REF_MODE_A1;
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_ref_mode(handle_, &mode) : kBindingInvalidHandle)) {
    return 0.0;
  }
  return static_cast<double>(mode);
}

void JsWorkbook::setRefMode(uint32_t mode) {
  succeeded(handle_ != nullptr ? fm_workbook_set_ref_mode(handle_, static_cast<std::int32_t>(mode))
                               : kBindingInvalidHandle);
}

bool JsWorkbook::date1904() const {
  std::int32_t value = 0;
  return succeeded(handle_ != nullptr ? fm_workbook_get_date1904(handle_, &value) : kBindingInvalidHandle) &&
         value != 0;
}

void JsWorkbook::setDate1904(bool date1904) {
  succeeded(handle_ != nullptr ? fm_workbook_set_date1904(handle_, date1904 ? 1 : 0) : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::pinnedNow() const {
  fm_civil_time_t now{};
  std::int32_t pinned = 0;
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_pinned_now(handle_, &now, &pinned) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  if (pinned == 0) {
    return emscripten::val::null();
  }
  emscripten::val civil = emscripten::val::object();
  civil.set("year", now.year);
  civil.set("month", now.month);
  civil.set("day", now.day);
  civil.set("hour", now.hour);
  civil.set("minute", now.minute);
  civil.set("second", now.second);
  return civil;
}

void JsWorkbook::setPinnedNow(uint32_t year, uint32_t month, uint32_t day, uint32_t hour, uint32_t minute,
                              uint32_t second) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  // Each field is widened rather than range-checked here: the C layer owns
  // the calendar domain, and a JS-side duplicate would be a second place to
  // keep in step with it.
  const fm_civil_time_t now{static_cast<std::int32_t>(year),   static_cast<std::int32_t>(month),
                            static_cast<std::int32_t>(day),    static_cast<std::int32_t>(hour),
                            static_cast<std::int32_t>(minute), static_cast<std::int32_t>(second)};
  succeeded(fm_workbook_set_pinned_now(handle_, &now));
}

void JsWorkbook::clearPinnedNow() {
  succeeded(handle_ != nullptr ? fm_workbook_clear_pinned_now(handle_) : kBindingInvalidHandle);
}

std::string JsWorkbook::profileId() const {
  const char* id = nullptr;
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_profile_id(handle_, &id) : kBindingInvalidHandle)) {
    return std::string();
  }
  return string_from_cstr(id);
}

void JsWorkbook::setProfileId(const std::string& profile_id) {
  succeeded(handle_ != nullptr ? fm_workbook_set_profile_id(handle_, profile_id.c_str()) : kBindingInvalidHandle);
}

double JsWorkbook::memoryUsage() const {
  std::size_t bytes = 0;
  if (handle_ == nullptr || fm_workbook_get_memory_usage(handle_, &bytes) != 0) {
    return 0.0;
  }
  return static_cast<double>(bytes);
}

emscripten::val JsWorkbook::fileName() const {
  const char* name = nullptr;
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_file_name(handle_, &name) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return cstr_or_null(name);
}

void JsWorkbook::setFileName(const std::string& name) {
  succeeded(handle_ != nullptr ? fm_workbook_set_file_name(handle_, name.c_str()) : kBindingInvalidHandle);
}

double JsWorkbook::partialRecalc(uint32_t sheet, emscripten::val viewport) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0.0;
  }
  JsNarrowNumericReader reader("recalcViewport");
  const fm_cell_rect_t vp = js_pull_range(viewport, &reader);
  if (!reader.ok()) {
    return 0.0;
  }
  uint32_t recomputed = 0;
  progress_callback_threw_ = false;
  fm_status_t rc = fm_sheet_recalc_viewport(handle_, sheet, &vp, &recomputed);
  const bool callback_threw = progress_callback_threw_;
  progress_callback_threw_ = false;
  if (rc == 0 && callback_threw) {
    record_progress_callback_threw();
    return 0.0;
  }
  if (!succeeded(rc)) {
    return 0.0;
  }
  return static_cast<double>(recomputed);
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
    record_type_error("setIterativeProgress: `callback` must be a function or null");
    return;
  }
  progress_callback_ = cb;
  succeeded(fm_workbook_set_iterative_progress(handle_, &JsWorkbook::iterativeProgressTrampoline, this));
}

// ---- Free helpers ------------------------------------------------------

emscripten::val evaluate_formula(const std::string& formula) {
  fm_workbook_t* wb = nullptr;
  if (!succeeded(fm_workbook_create_empty(&wb))) {
    return emscripten::val::undefined();
  }
  std::size_t sheet_index = 0;
  fm_value_array_t* array = nullptr;
  const bool ok = succeeded(fm_workbook_add_sheet(wb, "Sheet1", &sheet_index)) &&
                  succeeded(fm_sheet_evaluate_formula_array(wb, 0, 0, 0, formula.c_str(), &array));
  // A text payload borrows the array, so convert before destroying it.
  emscripten::val out = ok ? value_array_to_val(array) : emscripten::val::undefined();
  fm_value_array_destroy(array);
  fm_workbook_destroy(wb);
  return out;
}

std::string version_string() {
  return string_from_cstr(fm_get_version());
}

std::string status_string(int32_t status) {
  return string_from_cstr(fm_get_status_name(static_cast<fm_status_t>(status)));
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
    record_type_error("setLogSink: `sink` must be a function or null");
    return;
  }
  js_log_sink() = sink;
  succeeded(fm_set_log_sink(&log_sink_trampoline, nullptr));
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
