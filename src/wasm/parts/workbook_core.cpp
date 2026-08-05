//
// JsWorkbook lifecycle, sheet management, recalc / calc-mode plumbing,
// and the iterative-progress callback bridge. The split is by API
// surface area: cells / styles / pivot / etc. live in sibling files
// under `parts/`. The single `EMSCRIPTEN_BINDINGS` block is in
// `parts/bindings_register.cpp`.

#include <emscripten/bind.h>
#include <emscripten/val.h>

#include <cstdint>
#include <memory>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

JsWorkbook::~JsWorkbook() {
  if (handle_ != nullptr) {
    fm_workbook_destroy(handle_);
    handle_ = nullptr;
  }
}

JsWorkbook::JsWorkbook(JsWorkbook&& other) noexcept : handle_(other.handle_) {
  other.handle_ = nullptr;
}

JsWorkbook& JsWorkbook::operator=(JsWorkbook&& other) noexcept {
  if (this != &other) {
    if (handle_ != nullptr) {
      fm_workbook_destroy(handle_);
    }
    handle_ = other.handle_;
    other.handle_ = nullptr;
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

JsSaveResult JsWorkbook::saveEx(int32_t format) const {
  JsSaveResult r;
  if (handle_ == nullptr) {
    r.status = error_status(kBindingInvalidHandle);
    return r;
  }
  uint8_t* out = nullptr;
  std::size_t len = 0;
  fm_status_t rc = fm_workbook_save_ex(handle_, static_cast<fm_workbook_format_t>(format), &out, &len);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.bytes = bytes_to_val(out, len);
  fm_buffer_free(out);
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
  fm_status_t rc = fm_workbook_recalc(handle_);
  return status_from_rc(rc);
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
  fm_status_t rc = fm_workbook_set_calc_mode(handle_, static_cast<fm_calc_mode_t>(mode));
  return status_from_rc(rc);
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
  fm_status_t rc = fm_workbook_partial_recalc(handle_, &vp, &recomputed);
  if (rc != 0) {
    o.set("status", error_status(rc));
    o.set("recomputed", static_cast<uint32_t>(0));
    return o;
  }
  o.set("status", ok_status());
  o.set("recomputed", recomputed);
  return o;
}

// ---- Iterative-progress callback bridge --------------------------------
//
// The wrapper holds the JS callable in a function-local static slot for
// the lifetime of this WASM module. We do not pass `user_data` through
// the C ABI: the JS layer does not need it because the closure captures
// whatever state the JS caller wants. As a consequence, only ONE JS
// progress callback can be active at a time across all workbook handles
// in this WASM instance -- installing a new one displaces any previous
// registration. This matches the typical UI workflow (a single
// "calculation in progress" dialog) without requiring per-handle
// thread-local storage.

emscripten::val& JsWorkbook::js_progress_callback() {
  static emscripten::val cb = emscripten::val::null();
  return cb;
}

bool JsWorkbook::iterativeProgressTrampoline(uint32_t iteration, double max_residual, uint32_t max_iterations,
                                             void* /*user_data*/) {
  emscripten::val& cb = js_progress_callback();
  if (cb.isNull() || cb.isUndefined()) {
    return true;
  }
  emscripten::val ret = cb(iteration, max_residual, max_iterations);
  if (ret.isUndefined() || ret.isNull()) {
    return true;
  }
  return ret.as<bool>();
}

JsStatus JsWorkbook::setIterativeProgress(emscripten::val cb) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  if (cb.isNull() || cb.isUndefined()) {
    js_progress_callback() = emscripten::val::null();
    fm_status_t rc = fm_workbook_set_iterative_progress(handle_, nullptr, nullptr);
    return status_from_rc(rc);
  }
  js_progress_callback() = cb;
  fm_status_t rc = fm_workbook_set_iterative_progress(handle_, &JsWorkbook::iterativeProgressTrampoline, nullptr);
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

std::string last_error_message() {
  const char* s = fm_last_error_message();
  return s != nullptr ? std::string(s) : std::string();
}

std::string last_error_context() {
  const char* s = fm_last_error_context();
  return s != nullptr ? std::string(s) : std::string();
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
