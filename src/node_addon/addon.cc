// Copyright 2026 libraz. Licensed under the MIT License.
//
// Node.js N-API addon for the Formulon engine.
//
// This translation unit is compiled only when `FM_BUILD_NODE_ADDON=ON`.
// It is a thin C++ wrapper around the stable C ABI declared in
// `c_api/formulon_c.h`; the JavaScript surface NEVER touches
// `formulon::Workbook`, `formulon::Value`, or any other internal
// symbol directly -- exactly mirroring the WASM/embind binding's
// architectural stance.
//
// ## Design notes
//
//   * The whole engine is built `-fno-exceptions -fno-rtti`. node-addon-api
//     would normally throw C++ exceptions across the boundary; we
//     disable that via `NODE_ADDON_API_DISABLE_CPP_EXCEPTIONS` (and
//     `NAPI_DISABLE_CPP_EXCEPTIONS`) before any napi header include,
//     which the FormulonNodeAddon.cmake target propagates.
//
//   * Every fallible binding entry returns the same JS shape as the
//     embind binding:
//        Status = { ok: boolean, status: number,
//                   message: string, context: string }
//        Value  = { kind: number, number: number, boolean: number,
//                   text: string, errorCode: number }
//     The thread-local `fm_last_error_*` strings are snapshotted into
//     the Status envelope on every error path.
//
//   * `Workbook` is wrapped in `Napi::ObjectWrap<Workbook>`. The wrapper
//     owns the `fm_workbook_t*` and frees it in its destructor, which
//     N-API invokes when the JS object is garbage-collected.
//
//   * The full method surface mirrors `src/wasm/embind.cpp` field-for-
//     field so JS callers can swap between the WASM and native packages
//     without code changes. Field names on returned objects are kept
//     IDENTICAL to the embind shape.
//
//   * `setIterativeProgress` registers a JS callback through a static
//     `Napi::FunctionReference` slot. The slot is module-global (one
//     callback at a time across all workbook handles in the process),
//     mirroring the embind binding's single-slot policy. The C ABI's
//     iterative solver is synchronous within `recalc()` so the JS
//     callback always runs on the same thread that invoked recalc;
//     no thread-safe-function plumbing is required.

// NOTE: `NODE_ADDON_API_DISABLE_CPP_EXCEPTIONS` and
// `NAPI_DISABLE_CPP_EXCEPTIONS` are defined on the command line by
// `cmake/FormulonNodeAddon.cmake` so they apply uniformly to every TU
// that includes `napi.h`. They are required for the addon to compile
// under the project's `-fno-exceptions` policy.
//
// The compiler driver assigns the implicit replacement list `1` to a
// `-D X` flag, but `napi.h` later does an unconditional `#define X`
// (empty replacement list). To keep the build `-Werror`-clean we undef
// the command-line versions first; the napi.h `#ifdef NAPI_DISABLE_*`
// blocks immediately afterwards re-establish the same macros with the
// expected empty replacement list.
#ifdef NAPI_DISABLE_CPP_EXCEPTIONS
#undef NAPI_DISABLE_CPP_EXCEPTIONS
#endif
#ifdef NODE_ADDON_API_DISABLE_CPP_EXCEPTIONS
#undef NODE_ADDON_API_DISABLE_CPP_EXCEPTIONS
#endif
#define NAPI_DISABLE_CPP_EXCEPTIONS
#define NODE_ADDON_API_DISABLE_CPP_EXCEPTIONS

// NOLINTNEXTLINE(misc-include-cleaner): napi.h is the canonical entry point.
#include <napi.h>

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"

namespace {

// ---------------------------------------------------------------------
// Status / Value helpers
// ---------------------------------------------------------------------

/// Builds an `ok` Status envelope:
///   { ok: true, status: 0, message: "", context: "" }
Napi::Object MakeOkStatus(Napi::Env env) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("ok", Napi::Boolean::New(env, true));
  o.Set("status", Napi::Number::New(env, 0));
  o.Set("message", Napi::String::New(env, ""));
  o.Set("context", Napi::String::New(env, ""));
  return o;
}

/// Builds an error Status envelope, copying the thread-local
/// diagnostics surfaced by the most recent C-ABI call.
Napi::Object MakeErrorStatus(Napi::Env env, fm_status_t code) {
  const char* msg = fm_last_error_message();
  const char* ctx = fm_last_error_context();
  Napi::Object o = Napi::Object::New(env);
  o.Set("ok", Napi::Boolean::New(env, false));
  o.Set("status", Napi::Number::New(env, static_cast<int32_t>(code)));
  o.Set("message", Napi::String::New(env, msg != nullptr ? msg : ""));
  o.Set("context", Napi::String::New(env, ctx != nullptr ? ctx : ""));
  return o;
}

/// Converts a C-ABI status code into the shared JS Status envelope.
Napi::Object MakeStatus(Napi::Env env, fm_status_t code) {
  return code == 0 ? MakeOkStatus(env) : MakeErrorStatus(env, code);
}

/// Translates an `fm_value_t` into the JS Value shape.
Napi::Object TranslateValue(Napi::Env env, const fm_value_t& v) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("kind", Napi::Number::New(env, static_cast<int32_t>(v.kind)));
  // Default-zero all fields so consumers can read any field without
  // checking kind first (matches the embind shape).
  double number_field = 0.0;
  int32_t boolean_field = 0;
  std::string text_field;
  int32_t error_code_field = 0;
  switch (v.kind) {
    case FM_VAL_NUMBER:
      number_field = v.u.number;
      break;
    case FM_VAL_BOOL:
      boolean_field = v.u.boolean;
      break;
    case FM_VAL_TEXT:
      text_field = (v.u.text != nullptr) ? std::string(v.u.text) : std::string();
      break;
    case FM_VAL_ERROR:
      error_code_field = v.u.error_code;
      break;
    case FM_VAL_BLANK:
    case FM_VAL_ARRAY:
    case FM_VAL_REF:
    case FM_VAL_LAMBDA:
    default:
      break;
  }
  o.Set("number", Napi::Number::New(env, number_field));
  o.Set("boolean", Napi::Number::New(env, boolean_field));
  o.Set("text", Napi::String::New(env, text_field));
  o.Set("errorCode", Napi::Number::New(env, error_code_field));
  return o;
}

Napi::Object MakeFieldResult(Napi::Env env, Napi::Object status, const char* field, Napi::Value value) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", status);
  out.Set(field, value);
  return out;
}

Napi::Object MakeNumberFieldResult(Napi::Env env, Napi::Object status, const char* field, double value) {
  return MakeFieldResult(env, status, field, Napi::Number::New(env, value));
}

Napi::Object MakeStringFieldResult(Napi::Env env, Napi::Object status, const char* field, const char* value) {
  return MakeFieldResult(env, status, field, Napi::String::New(env, value != nullptr ? value : ""));
}

Napi::Object MakeValueResult(Napi::Env env, Napi::Object status, const fm_value_t& value) {
  return MakeFieldResult(env, status, "value", TranslateValue(env, value));
}

Napi::Object MakeEmptyValueResult(Napi::Env env, Napi::Object status) {
  fm_value_t empty{};
  return MakeValueResult(env, status, empty);
}

Napi::Object EmptyPivotLayoutResult(Napi::Env env, Napi::Object status) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", status);
  out.Set("top", Napi::Number::New(env, 0));
  out.Set("left", Napi::Number::New(env, 0));
  out.Set("rows", Napi::Number::New(env, 0));
  out.Set("cols", Napi::Number::New(env, 0));
  out.Set("cells", Napi::Array::New(env));
  return out;
}

Napi::Object TranslatePivotCell(Napi::Env env, const fm_pivot_cell_t& cell) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("row", Napi::Number::New(env, cell.row));
  out.Set("col", Napi::Number::New(env, cell.col));
  out.Set("value", TranslateValue(env, cell.value));
  out.Set("kind", Napi::Number::New(env, static_cast<int32_t>(cell.kind)));
  out.Set("depth", Napi::Number::New(env, cell.depth));
  out.Set("fieldName", Napi::String::New(env, cell.field_name != nullptr ? cell.field_name : ""));
  out.Set("numberFormat", Napi::String::New(env, cell.number_format != nullptr ? cell.number_format : ""));
  return out;
}

/// `kBindingNullPointer` ordinal mirrors `formulon::FormulonErrorCode`
/// in the 7000-7999 range allocated to bindings (see CLAUDE.md error
/// code table). The C ABI itself returns this code when a NULL pointer
/// crosses the boundary; we emit the same code from the JS side when
/// the wrapper is asked to operate on a destroyed handle.
constexpr fm_status_t kBindingNullPointer = 7000;

/// Translates an `fm_cf_color_t` into the JS shape used by embind.
Napi::Object TranslateCfColor(Napi::Env env, const fm_cf_color_t& c) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("r", Napi::Number::New(env, static_cast<int32_t>(c.r)));
  o.Set("g", Napi::Number::New(env, static_cast<int32_t>(c.g)));
  o.Set("b", Napi::Number::New(env, static_cast<int32_t>(c.b)));
  o.Set("a", Napi::Number::New(env, static_cast<int32_t>(c.a)));
  return o;
}

/// Translates an `fm_cf_match_t` into the JS shape used by embind.
Napi::Object TranslateCfMatch(Napi::Env env, const fm_cf_match_t& m) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("kind", Napi::Number::New(env, static_cast<int32_t>(m.kind)));
  o.Set("priority", Napi::Number::New(env, m.priority));
  o.Set("dxfIdEngaged", Napi::Number::New(env, m.dxf_id_engaged));
  o.Set("dxfId", Napi::Number::New(env, m.dxf_id));
  o.Set("color", TranslateCfColor(env, m.color));
  o.Set("barLengthPct", Napi::Number::New(env, m.bar_length_pct));
  o.Set("barAxisPositionPct", Napi::Number::New(env, m.bar_axis_position_pct));
  o.Set("barIsNegative", Napi::Number::New(env, m.bar_is_negative));
  o.Set("barFill", TranslateCfColor(env, m.bar_fill));
  o.Set("barBorderEngaged", Napi::Number::New(env, m.bar_border_engaged));
  o.Set("barBorder", TranslateCfColor(env, m.bar_border));
  o.Set("barGradient", Napi::Number::New(env, m.bar_gradient));
  o.Set("iconSetName", Napi::Number::New(env, m.icon_set_name));
  o.Set("iconIndex", Napi::Number::New(env, static_cast<int32_t>(m.icon_index)));
  return o;
}

// ---------------------------------------------------------------------
// JS-side iterative-progress callback slot
// ---------------------------------------------------------------------
//
// The C ABI's iterative solver is synchronous: it invokes the
// registered C callback inline from `fm_workbook_recalc` /
// `fm_workbook_partial_recalc` on the calling thread. That means the
// JS function we hold here is always invoked on the same thread that
// drove recalc, and we can safely call it through the standard
// `Napi::FunctionReference::Call` API (no thread-safe-function plumbing
// is required).
//
// Mirrors the embind binding's single-slot policy: there is one JS
// callback for the whole module, installing a new one displaces the
// previous, and clearing it (passing `null`) reverts to the default
// "always continue" behaviour.
struct ProgressSlot {
  Napi::FunctionReference fn;
  bool installed = false;
};

// Function-local static keeps the slot alive for the addon's lifetime
// without needing eager static initialisation of a Napi::Reference.
ProgressSlot& js_progress_slot() {
  static ProgressSlot slot;
  return slot;
}

// C-ABI compatible trampoline that forwards into the held JS callback.
// Returning `false` from the JS side aborts the iterative solve.
bool IterativeProgressTrampoline(uint32_t iteration, double max_residual, uint32_t max_iterations,
                                 void* /*user_data*/) {
  ProgressSlot& slot = js_progress_slot();
  if (!slot.installed || slot.fn.IsEmpty()) {
    return true;
  }
  Napi::Env env = slot.fn.Env();
  Napi::HandleScope scope(env);
  // The iteration index, max residual, and max iteration cap are all
  // delivered to JS as plain numbers; embind does the same.
  Napi::Value ret = slot.fn.Call({
      Napi::Number::New(env, iteration),
      Napi::Number::New(env, max_residual),
      Napi::Number::New(env, max_iterations),
  });
  if (env.IsExceptionPending()) {
    // Under -fno-exceptions, node-addon-api still routes JS exceptions
    // through `env.GetAndClearPendingException()`. Treat any pending
    // exception from the callback as "abort the solve" and clear it
    // so we don't propagate into the engine.
    (void)env.GetAndClearPendingException();
    return false;
  }
  if (ret.IsUndefined() || ret.IsNull()) {
    return true;
  }
  return ret.ToBoolean().Value();
}

// ---------------------------------------------------------------------
// Workbook ObjectWrap
// ---------------------------------------------------------------------

class Workbook : public Napi::ObjectWrap<Workbook> {
 public:
  static Napi::Function GetClass(Napi::Env env);

  explicit Workbook(const Napi::CallbackInfo& info);
  ~Workbook() override;

  Workbook(const Workbook&) = delete;
  Workbook& operator=(const Workbook&) = delete;
  Workbook(Workbook&&) = delete;
  Workbook& operator=(Workbook&&) = delete;

  // Static factories.
  static Napi::Value CreateDefault(const Napi::CallbackInfo& info);
  static Napi::Value CreateEmpty(const Napi::CallbackInfo& info);
  static Napi::Value LoadBytes(const Napi::CallbackInfo& info);

  // Cell mutation.
  Napi::Value SetNumber(const Napi::CallbackInfo& info);
  Napi::Value SetBool(const Napi::CallbackInfo& info);
  Napi::Value SetText(const Napi::CallbackInfo& info);
  Napi::Value SetBlank(const Napi::CallbackInfo& info);
  Napi::Value SetFormula(const Napi::CallbackInfo& info);

  // Cell read.
  Napi::Value GetValue(const Napi::CallbackInfo& info);

  // Lifecycle.
  Napi::Value IsValid(const Napi::CallbackInfo& info);

  // Recalc + save.
  Napi::Value Recalc(const Napi::CallbackInfo& info);
  Napi::Value PartialRecalc(const Napi::CallbackInfo& info);
  Napi::Value SetIterative(const Napi::CallbackInfo& info);
  Napi::Value SetIterativeProgress(const Napi::CallbackInfo& info);
  Napi::Value Save(const Napi::CallbackInfo& info);

  // Sheet operations.
  Napi::Value AddSheet(const Napi::CallbackInfo& info);
  Napi::Value RemoveSheet(const Napi::CallbackInfo& info);
  Napi::Value RenameSheet(const Napi::CallbackInfo& info);
  Napi::Value MoveSheet(const Napi::CallbackInfo& info);
  Napi::Value SheetCount(const Napi::CallbackInfo& info);
  Napi::Value SheetName(const Napi::CallbackInfo& info);

  // Row / column structural edits.
  Napi::Value InsertRows(const Napi::CallbackInfo& info);
  Napi::Value DeleteRows(const Napi::CallbackInfo& info);
  Napi::Value InsertCols(const Napi::CallbackInfo& info);
  Napi::Value DeleteCols(const Napi::CallbackInfo& info);

  // Iteration / metadata.
  Napi::Value CellCount(const Napi::CallbackInfo& info);
  Napi::Value CellAt(const Napi::CallbackInfo& info);
  Napi::Value DefinedNameCount(const Napi::CallbackInfo& info);
  Napi::Value DefinedNameAt(const Napi::CallbackInfo& info);
  Napi::Value TableCount(const Napi::CallbackInfo& info);
  Napi::Value TableAt(const Napi::CallbackInfo& info);
  Napi::Value PassthroughCount(const Napi::CallbackInfo& info);
  Napi::Value PassthroughAt(const Napi::CallbackInfo& info);
  Napi::Value PivotCount(const Napi::CallbackInfo& info);
  Napi::Value PivotLayout(const Napi::CallbackInfo& info);

  // PivotCache mutation.
  Napi::Value PivotCacheCount(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheIdAt(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheCreate(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRemove(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldCount(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldName(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldAdd(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldClear(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldSharedItemCount(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldAddSharedItemNumber(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldAddSharedItemText(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldAddSharedItemBool(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldAddSharedItemBlank(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheFieldClearSharedItems(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordCount(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordAdd(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordClear(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordSetNumber(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordSetText(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordSetBool(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordSetBlank(const Napi::CallbackInfo& info);
  Napi::Value PivotCacheRecordSetError(const Napi::CallbackInfo& info);

  // PivotTable mutation.
  Napi::Value PivotCreate(const Napi::CallbackInfo& info);
  Napi::Value PivotRemove(const Napi::CallbackInfo& info);
  Napi::Value PivotSetName(const Napi::CallbackInfo& info);
  Napi::Value PivotSetAnchor(const Napi::CallbackInfo& info);
  Napi::Value PivotSetGrandTotals(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldCount(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldAdd(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldClear(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldSetAxis(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldSetSort(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldSetSubtotalTop(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldAddAggregation(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldClearAggregations(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldAddItem(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldClearItems(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldSetItemVisible(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldAddSubtotalFn(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldClearSubtotalFns(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldSetDateGroup(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldClearDateGroup(const Napi::CallbackInfo& info);
  Napi::Value PivotFieldSetNumberFormat(const Napi::CallbackInfo& info);
  Napi::Value PivotSetRowFieldOrder(const Napi::CallbackInfo& info);
  Napi::Value PivotSetColFieldOrder(const Napi::CallbackInfo& info);
  Napi::Value PivotDataFieldCount(const Napi::CallbackInfo& info);
  Napi::Value PivotDataFieldAdd(const Napi::CallbackInfo& info);
  Napi::Value PivotDataFieldClear(const Napi::CallbackInfo& info);
  Napi::Value PivotDataFieldSet(const Napi::CallbackInfo& info);
  Napi::Value PivotFilterCount(const Napi::CallbackInfo& info);
  Napi::Value PivotFilterAdd(const Napi::CallbackInfo& info);
  Napi::Value PivotFilterClear(const Napi::CallbackInfo& info);
  Napi::Value PivotFilterRemoveAt(const Napi::CallbackInfo& info);

  // Defined names.
  Napi::Value SetDefinedName(const Napi::CallbackInfo& info);

  // Conditional formatting.
  Napi::Value EvaluateCfRange(const Napi::CallbackInfo& info);

  // Sheet view / layout.
  Napi::Value GetSheetView(const Napi::CallbackInfo& info);
  Napi::Value SetSheetZoom(const Napi::CallbackInfo& info);
  Napi::Value SetSheetFreeze(const Napi::CallbackInfo& info);
  Napi::Value SetSheetTabHidden(const Napi::CallbackInfo& info);
  Napi::Value GetSheetColumns(const Napi::CallbackInfo& info);
  Napi::Value SetColumnWidth(const Napi::CallbackInfo& info);
  Napi::Value SetColumnHidden(const Napi::CallbackInfo& info);
  Napi::Value SetColumnOutline(const Napi::CallbackInfo& info);
  Napi::Value GetSheetRowOverrides(const Napi::CallbackInfo& info);
  Napi::Value SetRowHeight(const Napi::CallbackInfo& info);
  Napi::Value SetRowHidden(const Napi::CallbackInfo& info);
  Napi::Value SetRowOutline(const Napi::CallbackInfo& info);

  // Styles.
  Napi::Value GetCellXfIndex(const Napi::CallbackInfo& info);
  Napi::Value SetCellXfIndex(const Napi::CallbackInfo& info);
  Napi::Value GetCellXf(const Napi::CallbackInfo& info);
  Napi::Value GetFont(const Napi::CallbackInfo& info);
  Napi::Value GetFill(const Napi::CallbackInfo& info);
  Napi::Value GetBorder(const Napi::CallbackInfo& info);
  Napi::Value GetNumFmt(const Napi::CallbackInfo& info);
  Napi::Value AddFont(const Napi::CallbackInfo& info);
  Napi::Value AddFill(const Napi::CallbackInfo& info);
  Napi::Value AddBorder(const Napi::CallbackInfo& info);
  Napi::Value AddNumFmt(const Napi::CallbackInfo& info);
  Napi::Value AddXf(const Napi::CallbackInfo& info);
  Napi::Value FontCount(const Napi::CallbackInfo& info);
  Napi::Value FillCount(const Napi::CallbackInfo& info);
  Napi::Value BorderCount(const Napi::CallbackInfo& info);
  Napi::Value XfCount(const Napi::CallbackInfo& info);

  // Sheet UI features (merges, comments, hyperlinks, validations).
  Napi::Value AddMerge(const Napi::CallbackInfo& info);
  Napi::Value RemoveMerge(const Napi::CallbackInfo& info);
  Napi::Value RemoveMergeAt(const Napi::CallbackInfo& info);
  Napi::Value ClearMerges(const Napi::CallbackInfo& info);
  Napi::Value GetMerges(const Napi::CallbackInfo& info);
  Napi::Value GetComment(const Napi::CallbackInfo& info);
  Napi::Value SetComment(const Napi::CallbackInfo& info);
  Napi::Value AddHyperlink(const Napi::CallbackInfo& info);
  Napi::Value GetHyperlinks(const Napi::CallbackInfo& info);
  Napi::Value RemoveHyperlink(const Napi::CallbackInfo& info);
  Napi::Value RemoveHyperlinkAt(const Napi::CallbackInfo& info);
  Napi::Value ClearHyperlinks(const Napi::CallbackInfo& info);
  Napi::Value GetValidations(const Napi::CallbackInfo& info);
  Napi::Value AddValidation(const Napi::CallbackInfo& info);
  Napi::Value RemoveValidationAt(const Napi::CallbackInfo& info);
  Napi::Value ClearValidations(const Napi::CallbackInfo& info);

 private:
  /// Extracts a `uint32_t` argument or sets `*ok=false` and surfaces
  /// the conversion failure through the JS error-status envelope.
  static uint32_t ArgU32(const Napi::CallbackInfo& info, size_t idx);
  static double ArgDouble(const Napi::CallbackInfo& info, size_t idx);
  static std::string ArgString(const Napi::CallbackInfo& info, size_t idx);
  static bool ArgBool(const Napi::CallbackInfo& info, size_t idx);

  /// Builds an error-Status envelope when the wrapper has been
  /// finalized / destroyed but JS still holds a reference.
  Napi::Object NullHandleError(Napi::Env env) const { return MakeErrorStatus(env, kBindingNullPointer); }

  fm_workbook_t* handle_ = nullptr;
};

Workbook::Workbook(const Napi::CallbackInfo& info) : Napi::ObjectWrap<Workbook>(info) {
  // Default constructor used by the static factories. They populate
  // `handle_` after construction via `wb->handle_ = ...`.
  //
  // External JS callers should NOT invoke `new Workbook()` directly;
  // the JS-side index.mjs only re-exports the static factories.
  (void)info;
}

Workbook::~Workbook() {
  if (handle_ != nullptr) {
    // If this workbook had the global progress callback installed
    // we tear that down too -- otherwise the dangling C-callback
    // pointer would be invoked on a destroyed handle if recalc is
    // somehow re-driven against the same workbook. The slot itself
    // keeps the JS function reference alive across the addon's
    // lifetime; only the per-handle registration is unwound here.
    (void)fm_workbook_set_iterative_progress(handle_, nullptr, nullptr);
    fm_workbook_destroy(handle_);
    handle_ = nullptr;
  }
}

uint32_t Workbook::ArgU32(const Napi::CallbackInfo& info, size_t idx) {
  if (idx >= info.Length()) {
    return 0;
  }
  return info[idx].ToNumber().Uint32Value();
}

double Workbook::ArgDouble(const Napi::CallbackInfo& info, size_t idx) {
  if (idx >= info.Length()) {
    return 0.0;
  }
  return info[idx].ToNumber().DoubleValue();
}

std::string Workbook::ArgString(const Napi::CallbackInfo& info, size_t idx) {
  if (idx >= info.Length()) {
    return std::string();
  }
  return info[idx].ToString().Utf8Value();
}

bool Workbook::ArgBool(const Napi::CallbackInfo& info, size_t idx) {
  if (idx >= info.Length()) {
    return false;
  }
  return info[idx].ToBoolean().Value();
}

// ---- Static factories -----------------------------------------------

Napi::Value Workbook::CreateDefault(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Function ctor = GetClass(env);
  Napi::Object jsobj = ctor.New({});
  Workbook* wb = Napi::ObjectWrap<Workbook>::Unwrap(jsobj);
  fm_status_t rc = fm_workbook_create(&wb->handle_);
  if (rc != 0) {
    // Even on failure return the wrapper; the caller can inspect
    // `lastErrorMessage()` and the next operation will fail with
    // `kBindingNullPointer`. This matches embind's behaviour.
    wb->handle_ = nullptr;
  }
  return jsobj;
}

Napi::Value Workbook::CreateEmpty(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Function ctor = GetClass(env);
  Napi::Object jsobj = ctor.New({});
  Workbook* wb = Napi::ObjectWrap<Workbook>::Unwrap(jsobj);
  fm_status_t rc = fm_workbook_create_empty(&wb->handle_);
  if (rc != 0) {
    wb->handle_ = nullptr;
  }
  return jsobj;
}

Napi::Value Workbook::LoadBytes(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Function ctor = GetClass(env);
  Napi::Object jsobj = ctor.New({});
  Workbook* wb = Napi::ObjectWrap<Workbook>::Unwrap(jsobj);
  if (info.Length() < 1 || !info[0].IsTypedArray()) {
    wb->handle_ = nullptr;
    return jsobj;
  }
  Napi::TypedArray ta = info[0].As<Napi::TypedArray>();
  if (ta.TypedArrayType() != napi_uint8_array) {
    wb->handle_ = nullptr;
    return jsobj;
  }
  Napi::Uint8Array u8 = ta.As<Napi::Uint8Array>();
  const uint8_t* data = u8.Data();
  const std::size_t len = u8.ElementLength();
  if (data == nullptr || len == 0) {
    wb->handle_ = nullptr;
    return jsobj;
  }
  fm_status_t rc = fm_workbook_load(data, len, &wb->handle_);
  if (rc != 0) {
    wb->handle_ = nullptr;
  }
  return jsobj;
}

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

// ---- Recalc + save --------------------------------------------------

Napi::Value Workbook::Recalc(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  fm_status_t rc = fm_workbook_recalc(handle_);
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
  ProgressSlot& slot = js_progress_slot();
  // Passing null / undefined clears the callback. Anything else MUST be
  // a JS function -- we surface a 7000-band error if it is not.
  if (info.Length() < 1 || info[0].IsNull() || info[0].IsUndefined()) {
    if (slot.installed) {
      slot.fn.Reset();
      slot.installed = false;
    }
    fm_status_t rc = fm_workbook_set_iterative_progress(handle_, nullptr, nullptr);
    return MakeStatus(env, rc);
  }
  if (!info[0].IsFunction()) {
    return MakeErrorStatus(env, kBindingNullPointer);
  }
  // Persist the JS function in the module-global slot. A FunctionReference
  // with refcount=1 keeps the function alive against GC for as long as
  // the slot owns it. Replace any previous registration.
  if (slot.installed) {
    slot.fn.Reset();
  }
  slot.fn = Napi::Persistent(info[0].As<Napi::Function>());
  slot.fn.SuppressDestruct();
  slot.installed = true;
  fm_status_t rc = fm_workbook_set_iterative_progress(handle_, &IterativeProgressTrampoline, nullptr);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::IsValid(const Napi::CallbackInfo& info) {
  return Napi::Boolean::New(info.Env(), handle_ != nullptr);
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

// ---- Sheet operations -----------------------------------------------

Napi::Value Workbook::AddSheet(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::string name = ArgString(info, 0);
  fm_status_t rc = fm_workbook_add_sheet(handle_, name.c_str());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::RemoveSheet(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t idx = ArgU32(info, 0);
  fm_status_t rc = fm_workbook_remove_sheet(handle_, idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::RenameSheet(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t idx = ArgU32(info, 0);
  const std::string name = ArgString(info, 1);
  fm_status_t rc = fm_workbook_rename_sheet(handle_, idx, name.c_str());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::MoveSheet(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t from_idx = ArgU32(info, 0);
  const uint32_t to_idx = ArgU32(info, 1);
  fm_status_t rc = fm_workbook_move_sheet(handle_, from_idx, to_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SheetCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, static_cast<double>(fm_workbook_sheet_count(handle_)));
}

Napi::Value Workbook::SheetName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeStringFieldResult(env, NullHandleError(env), "value", "");
  }
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
  const char* name = nullptr;
  fm_status_t rc = fm_workbook_sheet_name(handle_, idx, &name);
  if (rc != 0) {
    return MakeStringFieldResult(env, MakeErrorStatus(env, rc), "value", "");
  }
  return MakeStringFieldResult(env, MakeOkStatus(env), "value", name);
}

// ---- Row / column structural edits ----------------------------------

Napi::Value Workbook::InsertRows(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t count = ArgU32(info, 2);
  fm_status_t rc = fm_workbook_insert_rows(handle_, sheet, row, count);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::DeleteRows(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t count = ArgU32(info, 2);
  fm_status_t rc = fm_workbook_delete_rows(handle_, sheet, row, count);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::InsertCols(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t col = ArgU32(info, 1);
  const uint32_t count = ArgU32(info, 2);
  fm_status_t rc = fm_workbook_insert_cols(handle_, sheet, col, count);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::DeleteCols(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t col = ArgU32(info, 1);
  const uint32_t count = ArgU32(info, 2);
  fm_status_t rc = fm_workbook_delete_cols(handle_, sheet, col, count);
  return MakeStatus(env, rc);
}

// ---- Iteration / metadata accessors ---------------------------------

Napi::Value Workbook::CellCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  std::size_t count = 0;
  if (fm_workbook_cell_count(handle_, sheet, &count) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, static_cast<double>(count));
}

Napi::Value Workbook::CellAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 1));
  uint32_t row = 0;
  uint32_t col = 0;
  const char* formula = nullptr;
  fm_value_t v{};
  fm_status_t rc = fm_workbook_cell_at(handle_, sheet, idx, &row, &col, &formula, &v);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("row", Napi::Number::New(env, row));
  out.Set("col", Napi::Number::New(env, col));
  if (formula != nullptr) {
    out.Set("formula", Napi::String::New(env, formula));
  } else {
    out.Set("formula", env.Null());
  }
  out.Set("value", TranslateValue(env, v));
  return out;
}

Napi::Value Workbook::DefinedNameCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, static_cast<double>(fm_workbook_defined_name_count(handle_)));
}

Napi::Value Workbook::DefinedNameAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
  const char* name = nullptr;
  const char* formula = nullptr;
  fm_status_t rc = fm_workbook_defined_name_at(handle_, idx, &name, &formula);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("name", Napi::String::New(env, name != nullptr ? name : ""));
  out.Set("formula", Napi::String::New(env, formula != nullptr ? formula : ""));
  return out;
}

Napi::Value Workbook::TableCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, static_cast<double>(fm_workbook_table_count(handle_)));
}

Napi::Value Workbook::TableAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
  const char* name = nullptr;
  const char* display = nullptr;
  const char* ref = nullptr;
  std::size_t sheet_index = 0;
  fm_status_t rc = fm_workbook_table_at(handle_, idx, &name, &display, &ref, &sheet_index);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("name", Napi::String::New(env, name != nullptr ? name : ""));
  out.Set("displayName", Napi::String::New(env, display != nullptr ? display : ""));
  out.Set("ref", Napi::String::New(env, ref != nullptr ? ref : ""));
  out.Set("sheetIndex", Napi::Number::New(env, static_cast<double>(sheet_index)));
  return out;
}

Napi::Value Workbook::PassthroughCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, static_cast<double>(fm_workbook_passthrough_count(handle_)));
}

Napi::Value Workbook::PassthroughAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
  const char* path = nullptr;
  fm_status_t rc = fm_workbook_passthrough_at(handle_, idx, &path);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("path", Napi::String::New(env, path != nullptr ? path : ""));
  return out;
}

Napi::Value Workbook::PivotCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  std::size_t count = 0;
  if (fm_workbook_pivot_count(handle_, sheet, &count) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, static_cast<double>(count));
}

Napi::Value Workbook::PivotLayout(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return EmptyPivotLayoutResult(env, NullHandleError(env));
  }

  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
  fm_pivot_cells_t* cells = nullptr;
  fm_status_t rc = fm_workbook_pivot_layout(handle_, sheet, pivot_index, &cells);
  if (rc != 0) {
    return EmptyPivotLayoutResult(env, MakeErrorStatus(env, rc));
  }

  uint32_t top = 0;
  uint32_t left = 0;
  uint32_t rows = 0;
  uint32_t cols = 0;
  rc = fm_pivot_cells_bounds(cells, &top, &left, &rows, &cols);
  if (rc != 0) {
    fm_pivot_cells_destroy(cells);
    return EmptyPivotLayoutResult(env, MakeErrorStatus(env, rc));
  }

  const std::size_t count = fm_pivot_cells_count(cells);
  Napi::Array arr = Napi::Array::New(env, count);
  for (std::size_t i = 0; i < count; ++i) {
    fm_pivot_cell_t cell{};
    if (fm_pivot_cells_at(cells, i, &cell) != 0) {
      continue;
    }
    arr.Set(static_cast<uint32_t>(i), TranslatePivotCell(env, cell));
  }
  fm_pivot_cells_destroy(cells);

  Napi::Object out = Napi::Object::New(env);
  out.Set("status", MakeOkStatus(env));
  out.Set("top", Napi::Number::New(env, top));
  out.Set("left", Napi::Number::New(env, left));
  out.Set("rows", Napi::Number::New(env, rows));
  out.Set("cols", Napi::Number::New(env, cols));
  out.Set("cells", arr);
  return out;
}

// ---- PivotCache mutation --------------------------------------------
//
// Each method below is a thin wrapper over `fm_workbook_pivot_cache_*`.
// The shape mirrors the embind binding: number / boolean / string args
// map straight through; the return shape is the JS Status envelope or
// an `{status, index}` pair (for entries that surface a freshly
// assigned id / index). See `c_api/formulon_c.h` "Pivot cache mutation"
// for the underlying contract.

namespace {

// Returns `{ status, index }` for the success path of a `*_create` /
// `*_add` style entry that hands back a newly assigned index.
Napi::Object MakeIndexResult(Napi::Env env, Napi::Object status, uint32_t index) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", status);
  out.Set("index", Napi::Number::New(env, index));
  return out;
}

// Pulls an optional int32 field from a JS spec object; returns `dflt`
// when the field is missing / undefined / null.
int32_t SpecPullInt32(const Napi::Object& spec, const char* key, int32_t dflt) {
  if (!spec.Has(key)) {
    return dflt;
  }
  Napi::Value v = spec.Get(key);
  if (v.IsUndefined() || v.IsNull()) {
    return dflt;
  }
  return v.As<Napi::Number>().Int32Value();
}

// Pulls an optional uint32 field; returns `dflt` when missing.
uint32_t SpecPullU32(const Napi::Object& spec, const char* key, uint32_t dflt) {
  if (!spec.Has(key)) {
    return dflt;
  }
  Napi::Value v = spec.Get(key);
  if (v.IsUndefined() || v.IsNull()) {
    return dflt;
  }
  return v.As<Napi::Number>().Uint32Value();
}

double SpecPullDouble(const Napi::Object& spec, const char* key, double dflt) {
  if (!spec.Has(key)) {
    return dflt;
  }
  Napi::Value v = spec.Get(key);
  if (v.IsUndefined() || v.IsNull()) {
    return dflt;
  }
  return v.As<Napi::Number>().DoubleValue();
}

bool SpecPullBool(const Napi::Object& spec, const char* key, bool dflt) {
  if (!spec.Has(key)) {
    return dflt;
  }
  Napi::Value v = spec.Get(key);
  if (v.IsUndefined() || v.IsNull()) {
    return dflt;
  }
  return v.ToBoolean().Value();
}

// Returns whether the spec carries a non-null entry for `key`. Used to
// decide whether a `const char*` field should be forwarded as nullptr.
bool SpecHas(const Napi::Object& spec, const char* key) {
  if (!spec.Has(key)) {
    return false;
  }
  Napi::Value v = spec.Get(key);
  return !v.IsUndefined() && !v.IsNull();
}

}  // namespace

Napi::Value Workbook::PivotCacheCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  std::size_t count = 0;
  if (fm_workbook_pivot_cache_count(handle_, &count) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, static_cast<double>(count));
}

Napi::Value Workbook::PivotCacheIdAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
  uint32_t out = 0;
  fm_status_t rc = fm_workbook_pivot_cache_id_at(handle_, idx, &out);
  if (rc != 0) {
    return MakeIndexResult(env, MakeErrorStatus(env, rc), 0);
  }
  return MakeIndexResult(env, MakeOkStatus(env), out);
}

Napi::Value Workbook::PivotCacheCreate(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const uint32_t requested = ArgU32(info, 0);
  uint32_t out = 0;
  fm_status_t rc = fm_workbook_pivot_cache_create(handle_, requested, &out);
  if (rc != 0) {
    return MakeIndexResult(env, MakeErrorStatus(env, rc), 0);
  }
  return MakeIndexResult(env, MakeOkStatus(env), out);
}

Napi::Value Workbook::PivotCacheRemove(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  fm_status_t rc = fm_workbook_pivot_cache_remove(handle_, cache_id);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheFieldCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  std::size_t count = 0;
  if (fm_workbook_pivot_cache_field_count(handle_, cache_id, &count) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, static_cast<double>(count));
}

Napi::Value Workbook::PivotCacheFieldName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    out.Set("value", Napi::String::New(env, ""));
    return out;
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const char* name = nullptr;
  fm_status_t rc = fm_workbook_pivot_cache_field_name(handle_, cache_id, field_idx, &name);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    out.Set("value", Napi::String::New(env, ""));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("value", Napi::String::New(env, name != nullptr ? name : ""));
  return out;
}

Napi::Value Workbook::PivotCacheFieldAdd(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::string name = ArgString(info, 1);
  std::size_t out = 0;
  fm_status_t rc = fm_workbook_pivot_cache_field_add(handle_, cache_id, name.c_str(), &out);
  if (rc != 0) {
    return MakeIndexResult(env, MakeErrorStatus(env, rc), 0);
  }
  return MakeIndexResult(env, MakeOkStatus(env), static_cast<uint32_t>(out));
}

Napi::Value Workbook::PivotCacheFieldClear(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  fm_status_t rc = fm_workbook_pivot_cache_field_clear(handle_, cache_id);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheFieldSharedItemCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 1));
  std::size_t count = 0;
  if (fm_workbook_pivot_cache_field_shared_item_count(handle_, cache_id, field_idx, &count) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, static_cast<double>(count));
}

Napi::Value Workbook::PivotCacheFieldAddSharedItemNumber(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const double value = ArgDouble(info, 2);
  fm_status_t rc = fm_workbook_pivot_cache_field_add_shared_item_number(handle_, cache_id, field_idx, value);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheFieldAddSharedItemText(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::string utf8 = ArgString(info, 2);
  fm_status_t rc = fm_workbook_pivot_cache_field_add_shared_item_text(handle_, cache_id, field_idx, utf8.c_str());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheFieldAddSharedItemBool(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const bool value = ArgBool(info, 2);
  fm_status_t rc = fm_workbook_pivot_cache_field_add_shared_item_bool(handle_, cache_id, field_idx, value ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheFieldAddSharedItemBlank(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 1));
  fm_status_t rc = fm_workbook_pivot_cache_field_add_shared_item_blank(handle_, cache_id, field_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheFieldClearSharedItems(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 1));
  fm_status_t rc = fm_workbook_pivot_cache_field_clear_shared_items(handle_, cache_id, field_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheRecordCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  std::size_t count = 0;
  if (fm_workbook_pivot_cache_record_count(handle_, cache_id, &count) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, static_cast<double>(count));
}

Napi::Value Workbook::PivotCacheRecordAdd(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  std::size_t out = 0;
  fm_status_t rc = fm_workbook_pivot_cache_record_add(handle_, cache_id, &out);
  if (rc != 0) {
    return MakeIndexResult(env, MakeErrorStatus(env, rc), 0);
  }
  return MakeIndexResult(env, MakeOkStatus(env), static_cast<uint32_t>(out));
}

Napi::Value Workbook::PivotCacheRecordClear(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  fm_status_t rc = fm_workbook_pivot_cache_record_clear(handle_, cache_id);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheRecordSetNumber(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t record_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const double value = ArgDouble(info, 3);
  fm_status_t rc = fm_workbook_pivot_cache_record_set_number(handle_, cache_id, record_idx, field_idx, value);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheRecordSetText(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t record_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const std::string utf8 = ArgString(info, 3);
  fm_status_t rc = fm_workbook_pivot_cache_record_set_text(handle_, cache_id, record_idx, field_idx, utf8.c_str());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheRecordSetBool(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t record_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const bool value = ArgBool(info, 3);
  fm_status_t rc = fm_workbook_pivot_cache_record_set_bool(handle_, cache_id, record_idx, field_idx, value ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheRecordSetBlank(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t record_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  fm_status_t rc = fm_workbook_pivot_cache_record_set_blank(handle_, cache_id, record_idx, field_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheRecordSetError(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t record_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const int32_t error_code = info.Length() > 3 ? info[3].As<Napi::Number>().Int32Value() : 0;
  fm_status_t rc = fm_workbook_pivot_cache_record_set_error(handle_, cache_id, record_idx, field_idx,
                                                            static_cast<fm_error_code_t>(error_code));
  return MakeStatus(env, rc);
}

// ---- PivotTable mutation --------------------------------------------

Napi::Value Workbook::PivotCreate(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::string name = ArgString(info, 1);
  const uint32_t cache_id = ArgU32(info, 2);
  const uint32_t anchor_row = ArgU32(info, 3);
  const uint32_t anchor_col = ArgU32(info, 4);
  std::size_t out = 0;
  fm_status_t rc = fm_workbook_pivot_create(handle_, sheet, name.c_str(), cache_id, anchor_row, anchor_col, &out);
  if (rc != 0) {
    return MakeIndexResult(env, MakeErrorStatus(env, rc), 0);
  }
  return MakeIndexResult(env, MakeOkStatus(env), static_cast<uint32_t>(out));
}

Napi::Value Workbook::PivotRemove(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  fm_status_t rc = fm_workbook_pivot_remove(handle_, sheet, pivot_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotSetName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::string name = ArgString(info, 2);
  fm_status_t rc = fm_workbook_pivot_set_name(handle_, sheet, pivot_idx, name.c_str());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotSetAnchor(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const uint32_t anchor_row = ArgU32(info, 2);
  const uint32_t anchor_col = ArgU32(info, 3);
  const uint32_t span_rows = ArgU32(info, 4);
  const uint32_t span_cols = ArgU32(info, 5);
  fm_status_t rc =
      fm_workbook_pivot_set_anchor(handle_, sheet, pivot_idx, anchor_row, anchor_col, span_rows, span_cols);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotSetGrandTotals(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const bool rows_enabled = ArgBool(info, 2);
  const bool cols_enabled = ArgBool(info, 3);
  fm_status_t rc =
      fm_workbook_pivot_set_grand_totals(handle_, sheet, pivot_idx, rows_enabled ? 1 : 0, cols_enabled ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  std::size_t count = 0;
  if (fm_workbook_pivot_field_count(handle_, sheet, pivot_idx, &count) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, static_cast<double>(count));
}

Napi::Value Workbook::PivotFieldAdd(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  if (info.Length() < 3 || !info[2].IsObject()) {
    return MakeIndexResult(env, MakeErrorStatus(env, kBindingNullPointer), 0);
  }
  Napi::Object spec = info[2].As<Napi::Object>();

  const bool has_custom = SpecHas(spec, "customName");
  const std::string source_name = spec.Get("sourceName").ToString().Utf8Value();
  const std::string custom_name = has_custom ? spec.Get("customName").ToString().Utf8Value() : std::string();
  const bool has_nfmt = SpecHas(spec, "numberFormat");
  const std::string number_format = has_nfmt ? spec.Get("numberFormat").ToString().Utf8Value() : std::string();

  fm_pivot_field_spec_t c_spec{};
  c_spec.source_name = source_name.c_str();
  c_spec.custom_name = has_custom ? custom_name.c_str() : nullptr;
  c_spec.axis = static_cast<fm_pivot_axis_t>(SpecPullU32(spec, "axis", 0U));
  c_spec.subtotal_top = SpecPullBool(spec, "subtotalTop", false) ? 1 : 0;
  c_spec.number_format = has_nfmt ? number_format.c_str() : nullptr;

  std::size_t out = 0;
  fm_status_t rc = fm_workbook_pivot_field_add(handle_, sheet, pivot_idx, &c_spec, &out);
  if (rc != 0) {
    return MakeIndexResult(env, MakeErrorStatus(env, rc), 0);
  }
  return MakeIndexResult(env, MakeOkStatus(env), static_cast<uint32_t>(out));
}

Napi::Value Workbook::PivotFieldClear(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  fm_status_t rc = fm_workbook_pivot_field_clear(handle_, sheet, pivot_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldSetAxis(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const uint32_t axis = ArgU32(info, 3);
  fm_status_t rc =
      fm_workbook_pivot_field_set_axis(handle_, sheet, pivot_idx, field_idx, static_cast<fm_pivot_axis_t>(axis));
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldSetSort(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const bool ascending = ArgBool(info, 3);
  const std::string by_field = ArgString(info, 4);
  const char* by = by_field.empty() ? nullptr : by_field.c_str();
  fm_status_t rc = fm_workbook_pivot_field_set_sort(handle_, sheet, pivot_idx, field_idx, ascending ? 1 : 0, by);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldSetSubtotalTop(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const bool top = ArgBool(info, 3);
  fm_status_t rc = fm_workbook_pivot_field_set_subtotal_top(handle_, sheet, pivot_idx, field_idx, top ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldAddAggregation(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const uint32_t agg = ArgU32(info, 3);
  fm_status_t rc = fm_workbook_pivot_field_add_aggregation(handle_, sheet, pivot_idx, field_idx,
                                                           static_cast<fm_pivot_aggregation_t>(agg));
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldClearAggregations(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  fm_status_t rc = fm_workbook_pivot_field_clear_aggregations(handle_, sheet, pivot_idx, field_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldAddItem(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const std::string name = ArgString(info, 3);
  const bool visible = ArgBool(info, 4);
  fm_status_t rc =
      fm_workbook_pivot_field_add_item(handle_, sheet, pivot_idx, field_idx, name.c_str(), visible ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldClearItems(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  fm_status_t rc = fm_workbook_pivot_field_clear_items(handle_, sheet, pivot_idx, field_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldSetItemVisible(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const std::size_t item_idx = static_cast<std::size_t>(ArgU32(info, 3));
  const bool visible = ArgBool(info, 4);
  fm_status_t rc =
      fm_workbook_pivot_field_set_item_visible(handle_, sheet, pivot_idx, field_idx, item_idx, visible ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldAddSubtotalFn(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const uint32_t agg = ArgU32(info, 3);
  fm_status_t rc = fm_workbook_pivot_field_add_subtotal_fn(handle_, sheet, pivot_idx, field_idx,
                                                           static_cast<fm_pivot_aggregation_t>(agg));
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldClearSubtotalFns(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  fm_status_t rc = fm_workbook_pivot_field_clear_subtotal_fns(handle_, sheet, pivot_idx, field_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldSetDateGroup(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const uint32_t granularity = ArgU32(info, 3);
  const uint32_t calendar = ArgU32(info, 4);
  const int32_t start_year = info.Length() > 5 ? info[5].As<Napi::Number>().Int32Value() : -1;
  const int32_t end_year = info.Length() > 6 ? info[6].As<Napi::Number>().Int32Value() : -1;
  fm_status_t rc = fm_workbook_pivot_field_set_date_group(
      handle_, sheet, pivot_idx, field_idx, static_cast<fm_pivot_date_grouping_t>(granularity),
      static_cast<fm_pivot_calendar_t>(calendar), start_year, end_year);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldClearDateGroup(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  fm_status_t rc = fm_workbook_pivot_field_clear_date_group(handle_, sheet, pivot_idx, field_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldSetNumberFormat(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const std::string utf8 = ArgString(info, 3);
  fm_status_t rc = fm_workbook_pivot_field_set_number_format(handle_, sheet, pivot_idx, field_idx, utf8.c_str());
  return MakeStatus(env, rc);
}

namespace {

// Pulls a numeric array from `info[idx]` into a `std::vector<uint32_t>`.
// Returns an empty vector when the argument is missing / undefined / null
// or not array-shaped.
std::vector<uint32_t> ReadU32Array(const Napi::CallbackInfo& info, size_t idx) {
  std::vector<uint32_t> out;
  if (idx >= info.Length()) {
    return out;
  }
  Napi::Value v = info[idx];
  if (!v.IsArray()) {
    return out;
  }
  Napi::Array arr = v.As<Napi::Array>();
  const uint32_t len = arr.Length();
  out.reserve(len);
  for (uint32_t i = 0; i < len; ++i) {
    out.push_back(arr.Get(i).As<Napi::Number>().Uint32Value());
  }
  return out;
}

}  // namespace

Napi::Value Workbook::PivotSetRowFieldOrder(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::vector<uint32_t> indices = ReadU32Array(info, 2);
  fm_status_t rc = fm_workbook_pivot_set_row_field_order(handle_, sheet, pivot_idx,
                                                         indices.empty() ? nullptr : indices.data(), indices.size());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotSetColFieldOrder(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::vector<uint32_t> indices = ReadU32Array(info, 2);
  fm_status_t rc = fm_workbook_pivot_set_col_field_order(handle_, sheet, pivot_idx,
                                                         indices.empty() ? nullptr : indices.data(), indices.size());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotDataFieldCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  std::size_t count = 0;
  if (fm_workbook_pivot_data_field_count(handle_, sheet, pivot_idx, &count) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, static_cast<double>(count));
}

namespace {

// Builds an `fm_pivot_data_field_spec_t` from a JS spec object. The
// `name_buf` / `nfmt_buf` strings keep the borrowed `const char*`
// pointers alive for the caller; `has_nfmt` is set when the spec
// carries a non-null `numberFormat`.
void BuildDataFieldSpec(const Napi::Object& spec, fm_pivot_data_field_spec_t& out, std::string& name_buf,
                        std::string& nfmt_buf, bool& has_nfmt) {
  name_buf = spec.Get("name").ToString().Utf8Value();
  has_nfmt = SpecHas(spec, "numberFormat");
  nfmt_buf = has_nfmt ? spec.Get("numberFormat").ToString().Utf8Value() : std::string();
  out.name = name_buf.c_str();
  out.field_index = SpecPullU32(spec, "fieldIndex", 0U);
  out.aggregation = static_cast<fm_pivot_aggregation_t>(SpecPullU32(spec, "aggregation", 0U));
  out.number_format = has_nfmt ? nfmt_buf.c_str() : nullptr;
  out.show_as = static_cast<fm_pivot_show_as_t>(SpecPullU32(spec, "showAs", 0U));
  out.show_as_base_field = SpecPullInt32(spec, "showAsBaseField", -1);
  out.show_as_base_item = SpecPullInt32(spec, "showAsBaseItem", -1);
}

}  // namespace

Napi::Value Workbook::PivotDataFieldAdd(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  if (info.Length() < 3 || !info[2].IsObject()) {
    return MakeIndexResult(env, MakeErrorStatus(env, kBindingNullPointer), 0);
  }
  Napi::Object spec = info[2].As<Napi::Object>();
  fm_pivot_data_field_spec_t c_spec{};
  std::string name_buf;
  std::string nfmt_buf;
  bool has_nfmt = false;
  BuildDataFieldSpec(spec, c_spec, name_buf, nfmt_buf, has_nfmt);
  std::size_t out = 0;
  fm_status_t rc = fm_workbook_pivot_data_field_add(handle_, sheet, pivot_idx, &c_spec, &out);
  if (rc != 0) {
    return MakeIndexResult(env, MakeErrorStatus(env, rc), 0);
  }
  return MakeIndexResult(env, MakeOkStatus(env), static_cast<uint32_t>(out));
}

Napi::Value Workbook::PivotDataFieldClear(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  fm_status_t rc = fm_workbook_pivot_data_field_clear(handle_, sheet, pivot_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotDataFieldSet(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t data_field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  if (info.Length() < 4 || !info[3].IsObject()) {
    return MakeErrorStatus(env, kBindingNullPointer);
  }
  Napi::Object spec = info[3].As<Napi::Object>();
  fm_pivot_data_field_spec_t c_spec{};
  std::string name_buf;
  std::string nfmt_buf;
  bool has_nfmt = false;
  BuildDataFieldSpec(spec, c_spec, name_buf, nfmt_buf, has_nfmt);
  fm_status_t rc = fm_workbook_pivot_data_field_set(handle_, sheet, pivot_idx, data_field_idx, &c_spec);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFilterCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  std::size_t count = 0;
  if (fm_workbook_pivot_filter_count(handle_, sheet, pivot_idx, &count) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, static_cast<double>(count));
}

Napi::Value Workbook::PivotFilterAdd(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  if (info.Length() < 3 || !info[2].IsObject()) {
    return MakeErrorStatus(env, kBindingNullPointer);
  }
  Napi::Object spec = info[2].As<Napi::Object>();

  const std::string field_name = spec.Get("fieldName").ToString().Utf8Value();
  const bool has_text = SpecHas(spec, "valueText");
  const std::string value_text = has_text ? spec.Get("valueText").ToString().Utf8Value() : std::string();

  fm_pivot_filter_spec_t c_spec{};
  c_spec.axis = static_cast<fm_pivot_axis_t>(SpecPullU32(spec, "axis", 0U));
  c_spec.field_name = field_name.c_str();
  c_spec.type = static_cast<fm_pivot_filter_type_t>(SpecPullU32(spec, "type", 0U));
  c_spec.value_kind = static_cast<fm_pivot_filter_value_kind_t>(SpecPullInt32(spec, "valueKind", -1));
  c_spec.value_int = SpecPullInt32(spec, "valueInt", 0);
  c_spec.value_double = SpecPullDouble(spec, "valueDouble", 0.0);
  c_spec.value_text = has_text ? value_text.c_str() : nullptr;
  c_spec.value_high_kind = static_cast<fm_pivot_filter_value_kind_t>(SpecPullInt32(spec, "valueHighKind", -1));
  c_spec.value_high_int = SpecPullInt32(spec, "valueHighInt", 0);
  c_spec.value_high_double = SpecPullDouble(spec, "valueHighDouble", 0.0);

  fm_status_t rc = fm_workbook_pivot_filter_add(handle_, sheet, pivot_idx, &c_spec);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFilterClear(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  fm_status_t rc = fm_workbook_pivot_filter_clear(handle_, sheet, pivot_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFilterRemoveAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t filter_idx = static_cast<std::size_t>(ArgU32(info, 2));
  fm_status_t rc = fm_workbook_pivot_filter_remove_at(handle_, sheet, pivot_idx, filter_idx);
  return MakeStatus(env, rc);
}

// ---- Defined names --------------------------------------------------

Napi::Value Workbook::SetDefinedName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::string name = ArgString(info, 0);
  const std::string formula = ArgString(info, 1);
  fm_status_t rc = fm_workbook_set_defined_name(handle_, name.c_str(), formula.c_str());
  return MakeStatus(env, rc);
}

// ---- Conditional formatting -----------------------------------------

Napi::Value Workbook::EvaluateCfRange(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    out.Set("cells", Napi::Array::New(env));
    return out;
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t first_row = ArgU32(info, 1);
  const uint32_t first_col = ArgU32(info, 2);
  const uint32_t last_row = ArgU32(info, 3);
  const uint32_t last_col = ArgU32(info, 4);
  const double today_serial = ArgDouble(info, 5);
  fm_cf_results_t* results = nullptr;
  fm_status_t rc =
      fm_workbook_cf_evaluate_range(handle_, sheet, first_row, first_col, last_row, last_col, today_serial, &results);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    out.Set("cells", Napi::Array::New(env));
    return out;
  }
  const std::size_t cell_count = fm_cf_results_cell_count(results);
  Napi::Array cells = Napi::Array::New(env, cell_count);
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < cell_count; ++i) {
    uint32_t row = 0;
    uint32_t col = 0;
    std::size_t match_count = 0;
    if (fm_cf_results_cell_at(results, i, &row, &col, &match_count) != 0) {
      // Defensive: skip entries the C ABI declines to materialise.
      continue;
    }
    Napi::Array matches = Napi::Array::New(env, match_count);
    std::size_t mj = 0;
    for (std::size_t j = 0; j < match_count; ++j) {
      fm_cf_match_t m{};
      if (fm_cf_results_match_at(results, i, j, &m) != 0) {
        continue;
      }
      matches.Set(static_cast<uint32_t>(mj), TranslateCfMatch(env, m));
      ++mj;
    }
    Napi::Object cell = Napi::Object::New(env);
    cell.Set("row", Napi::Number::New(env, row));
    cell.Set("col", Napi::Number::New(env, col));
    cell.Set("matches", matches);
    cells.Set(static_cast<uint32_t>(emitted), cell);
    ++emitted;
  }
  fm_cf_results_destroy(results);
  out.Set("status", MakeOkStatus(env));
  out.Set("cells", cells);
  return out;
}

// ---- Sheet view / layout --------------------------------------------

Napi::Value Workbook::GetSheetView(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    Napi::Object view = Napi::Object::New(env);
    view.Set("zoomScale", Napi::Number::New(env, 100));
    view.Set("freezeRows", Napi::Number::New(env, 0));
    view.Set("freezeCols", Napi::Number::New(env, 0));
    view.Set("tabHidden", Napi::Number::New(env, 0));
    out.Set("view", view);
    return out;
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  fm_sheet_view_t v{};
  fm_status_t rc = fm_sheet_get_view(handle_, sheet, &v);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    Napi::Object view = Napi::Object::New(env);
    view.Set("zoomScale", Napi::Number::New(env, 100));
    view.Set("freezeRows", Napi::Number::New(env, 0));
    view.Set("freezeCols", Napi::Number::New(env, 0));
    view.Set("tabHidden", Napi::Number::New(env, 0));
    out.Set("view", view);
    return out;
  }
  Napi::Object view = Napi::Object::New(env);
  view.Set("zoomScale", Napi::Number::New(env, v.zoom_scale));
  view.Set("freezeRows", Napi::Number::New(env, v.freeze_rows));
  view.Set("freezeCols", Napi::Number::New(env, v.freeze_cols));
  view.Set("tabHidden", Napi::Number::New(env, v.tab_hidden));
  out.Set("status", MakeOkStatus(env));
  out.Set("view", view);
  return out;
}

Napi::Value Workbook::SetSheetZoom(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t zoom = ArgU32(info, 1);
  fm_status_t rc = fm_sheet_set_zoom(handle_, sheet, zoom);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetSheetFreeze(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t freeze_rows = ArgU32(info, 1);
  const uint32_t freeze_cols = ArgU32(info, 2);
  fm_status_t rc = fm_sheet_set_freeze(handle_, sheet, freeze_rows, freeze_cols);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetSheetTabHidden(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const bool hidden = ArgBool(info, 1);
  fm_status_t rc = fm_sheet_set_tab_hidden(handle_, sheet, hidden ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::GetSheetColumns(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    out.Set("columns", Napi::Array::New(env));
    return out;
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  std::size_t count = 0;
  fm_status_t rc = fm_sheet_get_column_count(handle_, sheet, &count);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    out.Set("columns", Napi::Array::New(env));
    return out;
  }
  Napi::Array arr = Napi::Array::New(env, count);
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < count; ++i) {
    fm_column_layout_t entry{};
    if (fm_sheet_get_column(handle_, sheet, i, &entry) != 0) {
      continue;
    }
    Napi::Object col = Napi::Object::New(env);
    col.Set("first", Napi::Number::New(env, entry.first));
    col.Set("last", Napi::Number::New(env, entry.last));
    col.Set("width", Napi::Number::New(env, entry.width));
    col.Set("hidden", Napi::Number::New(env, entry.hidden));
    col.Set("outlineLevel", Napi::Number::New(env, static_cast<int32_t>(entry.outline_level)));
    arr.Set(static_cast<uint32_t>(emitted), col);
    ++emitted;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("columns", arr);
  return out;
}

Napi::Value Workbook::SetColumnWidth(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t first = ArgU32(info, 1);
  const uint32_t last = ArgU32(info, 2);
  const double width = ArgDouble(info, 3);
  fm_status_t rc = fm_sheet_set_column_width(handle_, sheet, first, last, width);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetColumnHidden(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t first = ArgU32(info, 1);
  const uint32_t last = ArgU32(info, 2);
  const bool hidden = ArgBool(info, 3);
  fm_status_t rc = fm_sheet_set_column_hidden(handle_, sheet, first, last, hidden ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetColumnOutline(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t first = ArgU32(info, 1);
  const uint32_t last = ArgU32(info, 2);
  uint32_t level = ArgU32(info, 3);
  if (level > 255U) {
    level = 255U;
  }
  fm_status_t rc = fm_sheet_set_column_outline(handle_, sheet, first, last, static_cast<uint8_t>(level));
  return MakeStatus(env, rc);
}

Napi::Value Workbook::GetSheetRowOverrides(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    out.Set("rows", Napi::Array::New(env));
    return out;
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  std::size_t count = 0;
  fm_status_t rc = fm_sheet_get_row_override_count(handle_, sheet, &count);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    out.Set("rows", Napi::Array::New(env));
    return out;
  }
  Napi::Array arr = Napi::Array::New(env, count);
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < count; ++i) {
    fm_row_layout_t entry{};
    if (fm_sheet_get_row_override(handle_, sheet, i, &entry) != 0) {
      continue;
    }
    Napi::Object row = Napi::Object::New(env);
    row.Set("row", Napi::Number::New(env, entry.row));
    row.Set("height", Napi::Number::New(env, entry.height));
    row.Set("hidden", Napi::Number::New(env, entry.hidden));
    row.Set("outlineLevel", Napi::Number::New(env, static_cast<int32_t>(entry.outline_level)));
    arr.Set(static_cast<uint32_t>(emitted), row);
    ++emitted;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("rows", arr);
  return out;
}

Napi::Value Workbook::SetRowHeight(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const double height = ArgDouble(info, 2);
  fm_status_t rc = fm_sheet_set_row_height(handle_, sheet, row, height);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetRowHidden(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const bool hidden = ArgBool(info, 2);
  fm_status_t rc = fm_sheet_set_row_hidden(handle_, sheet, row, hidden ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetRowOutline(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  uint32_t level = ArgU32(info, 2);
  if (level > 255U) {
    level = 255U;
  }
  fm_status_t rc = fm_sheet_set_row_outline(handle_, sheet, row, static_cast<uint8_t>(level));
  return MakeStatus(env, rc);
}

// ---- Styles ---------------------------------------------------------

Napi::Value Workbook::GetCellXfIndex(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "xfIndex", 0);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  uint32_t xf = 0;
  fm_status_t rc = fm_cell_get_xf_index(handle_, sheet, row, col, &xf);
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "xfIndex", 0);
  }
  return MakeNumberFieldResult(env, MakeOkStatus(env), "xfIndex", xf);
}

Napi::Value Workbook::SetCellXfIndex(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  const uint32_t xf = ArgU32(info, 3);
  fm_status_t rc = fm_cell_set_xf_index(handle_, sheet, row, col, xf);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::GetCellXf(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint32_t xf_index = ArgU32(info, 0);
  fm_cell_xf xf{};
  fm_status_t rc = fm_styles_get_cell_xf(handle_, xf_index, &xf);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("fontIndex", Napi::Number::New(env, xf.font_index));
  out.Set("fillIndex", Napi::Number::New(env, xf.fill_index));
  out.Set("borderIndex", Napi::Number::New(env, xf.border_index));
  out.Set("numFmtId", Napi::Number::New(env, static_cast<uint32_t>(xf.num_fmt_id)));
  out.Set("horizontalAlign", Napi::Number::New(env, static_cast<uint32_t>(xf.horizontal_align)));
  out.Set("verticalAlign", Napi::Number::New(env, static_cast<uint32_t>(xf.vertical_align)));
  out.Set("wrapText", Napi::Boolean::New(env, xf.wrap_text != 0));
  return out;
}

Napi::Value Workbook::GetFont(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint32_t font_index = ArgU32(info, 0);
  fm_font_record f{};
  fm_status_t rc = fm_styles_get_font(handle_, font_index, &f);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("name", Napi::String::New(env, f.name != nullptr ? f.name : ""));
  out.Set("size", Napi::Number::New(env, f.size));
  out.Set("colorArgb", Napi::Number::New(env, f.color_argb));
  out.Set("bold", Napi::Boolean::New(env, f.bold != 0));
  out.Set("italic", Napi::Boolean::New(env, f.italic != 0));
  out.Set("strike", Napi::Boolean::New(env, f.strike != 0));
  out.Set("underline", Napi::Number::New(env, static_cast<uint32_t>(f.underline)));
  return out;
}

Napi::Value Workbook::GetFill(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint32_t fill_index = ArgU32(info, 0);
  fm_fill_record f{};
  fm_status_t rc = fm_styles_get_fill(handle_, fill_index, &f);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("pattern", Napi::Number::New(env, static_cast<uint32_t>(f.pattern)));
  out.Set("fgArgb", Napi::Number::New(env, f.fg_argb));
  out.Set("bgArgb", Napi::Number::New(env, f.bg_argb));
  return out;
}

Napi::Value Workbook::GetBorder(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint32_t border_index = ArgU32(info, 0);
  fm_border_record b{};
  fm_status_t rc = fm_styles_get_border(handle_, border_index, &b);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  auto side_obj = [&](const fm_border_side& s) {
    Napi::Object so = Napi::Object::New(env);
    so.Set("style", Napi::Number::New(env, static_cast<uint32_t>(s.style)));
    so.Set("colorArgb", Napi::Number::New(env, s.color_argb));
    return so;
  };
  out.Set("status", MakeOkStatus(env));
  out.Set("left", side_obj(b.left));
  out.Set("right", side_obj(b.right));
  out.Set("top", side_obj(b.top));
  out.Set("bottom", side_obj(b.bottom));
  out.Set("diagonal", side_obj(b.diagonal));
  out.Set("diagonalUp", Napi::Boolean::New(env, b.diagonal_up != 0));
  out.Set("diagonalDown", Napi::Boolean::New(env, b.diagonal_down != 0));
  return out;
}

Napi::Value Workbook::GetNumFmt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint32_t num_fmt_id = ArgU32(info, 0);
  const char* s = nullptr;
  fm_status_t rc = fm_styles_get_num_fmt_string(handle_, static_cast<uint16_t>(num_fmt_id), &s);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("numFmtId", Napi::Number::New(env, num_fmt_id));
  out.Set("formatCode", Napi::String::New(env, s != nullptr ? s : ""));
  return out;
}

Napi::Value Workbook::AddFont(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "index", 0);
  }
  Napi::Object record = (info.Length() > 0 && info[0].IsObject()) ? info[0].As<Napi::Object>() : Napi::Object::New(env);
  std::string name;
  if (record.Has("name") && !record.Get("name").IsUndefined() && !record.Get("name").IsNull()) {
    name = record.Get("name").ToString().Utf8Value();
  }
  fm_font_record fr{};
  fr.name = name.c_str();
  fr.size = record.Has("size") ? record.Get("size").ToNumber().DoubleValue() : 11.0;
  fr.bold = (record.Has("bold") && record.Get("bold").ToBoolean().Value()) ? 1 : 0;
  fr.italic = (record.Has("italic") && record.Get("italic").ToBoolean().Value()) ? 1 : 0;
  fr.strike = (record.Has("strike") && record.Get("strike").ToBoolean().Value()) ? 1 : 0;
  fr.underline = record.Has("underline") ? static_cast<uint8_t>(record.Get("underline").ToNumber().Uint32Value()) : 0U;
  fr.color_argb = record.Has("colorArgb") ? record.Get("colorArgb").ToNumber().Uint32Value() : 0xFF000000U;
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_font(handle_, fr, &idx);
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "index", 0);
  }
  return MakeNumberFieldResult(env, MakeOkStatus(env), "index", idx);
}

Napi::Value Workbook::AddFill(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "index", 0);
  }
  Napi::Object record = (info.Length() > 0 && info[0].IsObject()) ? info[0].As<Napi::Object>() : Napi::Object::New(env);
  fm_fill_record fr{};
  fr.pattern = record.Has("pattern") ? static_cast<uint8_t>(record.Get("pattern").ToNumber().Uint32Value()) : 0U;
  fr.fg_argb = record.Has("fgArgb") ? record.Get("fgArgb").ToNumber().Uint32Value() : 0U;
  fr.bg_argb = record.Has("bgArgb") ? record.Get("bgArgb").ToNumber().Uint32Value() : 0U;
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_fill(handle_, fr, &idx);
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "index", 0);
  }
  return MakeNumberFieldResult(env, MakeOkStatus(env), "index", idx);
}

Napi::Value Workbook::AddBorder(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "index", 0);
  }
  Napi::Object record = (info.Length() > 0 && info[0].IsObject()) ? info[0].As<Napi::Object>() : Napi::Object::New(env);
  auto pull_side = [&](const char* key) {
    fm_border_side s{};
    if (record.Has(key) && record.Get(key).IsObject()) {
      Napi::Object so = record.Get(key).As<Napi::Object>();
      if (so.Has("style")) {
        s.style = static_cast<uint8_t>(so.Get("style").ToNumber().Uint32Value());
      }
      if (so.Has("colorArgb")) {
        s.color_argb = so.Get("colorArgb").ToNumber().Uint32Value();
      }
    }
    return s;
  };
  fm_border_record br{};
  br.left = pull_side("left");
  br.right = pull_side("right");
  br.top = pull_side("top");
  br.bottom = pull_side("bottom");
  br.diagonal = pull_side("diagonal");
  br.diagonal_up = (record.Has("diagonalUp") && record.Get("diagonalUp").ToBoolean().Value()) ? 1 : 0;
  br.diagonal_down = (record.Has("diagonalDown") && record.Get("diagonalDown").ToBoolean().Value()) ? 1 : 0;
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_border(handle_, br, &idx);
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "index", 0);
  }
  return MakeNumberFieldResult(env, MakeOkStatus(env), "index", idx);
}

Napi::Value Workbook::AddNumFmt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "numFmtId", 0);
  }
  const std::string code = ArgString(info, 0);
  uint16_t id = 0;
  fm_status_t rc = fm_styles_add_num_fmt(handle_, code.c_str(), &id);
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "numFmtId", 0);
  }
  return MakeNumberFieldResult(env, MakeOkStatus(env), "numFmtId", static_cast<uint32_t>(id));
}

Napi::Value Workbook::AddXf(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "index", 0);
  }
  Napi::Object record = (info.Length() > 0 && info[0].IsObject()) ? info[0].As<Napi::Object>() : Napi::Object::New(env);
  fm_cell_xf xf{};
  xf.font_index = record.Has("fontIndex") ? record.Get("fontIndex").ToNumber().Uint32Value() : 0U;
  xf.fill_index = record.Has("fillIndex") ? record.Get("fillIndex").ToNumber().Uint32Value() : 0U;
  xf.border_index = record.Has("borderIndex") ? record.Get("borderIndex").ToNumber().Uint32Value() : 0U;
  xf.num_fmt_id = record.Has("numFmtId") ? static_cast<uint16_t>(record.Get("numFmtId").ToNumber().Uint32Value()) : 0U;
  xf.horizontal_align =
      record.Has("horizontalAlign") ? static_cast<uint8_t>(record.Get("horizontalAlign").ToNumber().Uint32Value()) : 0U;
  xf.vertical_align =
      record.Has("verticalAlign") ? static_cast<uint8_t>(record.Get("verticalAlign").ToNumber().Uint32Value()) : 0U;
  xf.wrap_text = (record.Has("wrapText") && record.Get("wrapText").ToBoolean().Value()) ? 1 : 0;
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_cell_xf(handle_, xf, &idx);
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "index", 0);
  }
  return MakeNumberFieldResult(env, MakeOkStatus(env), "index", idx);
}

Napi::Value Workbook::FontCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  uint32_t n = 0;
  if (fm_styles_get_font_count(handle_, &n) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, n);
}

Napi::Value Workbook::FillCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  uint32_t n = 0;
  if (fm_styles_get_fill_count(handle_, &n) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, n);
}

Napi::Value Workbook::BorderCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  uint32_t n = 0;
  if (fm_styles_get_border_count(handle_, &n) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, n);
}

Napi::Value Workbook::XfCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  uint32_t n = 0;
  if (fm_styles_get_cell_xf_count(handle_, &n) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, n);
}

// ---- Sheet UI features (merges, comments, hyperlinks, validations) --

Napi::Value Workbook::AddMerge(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  fm_merge_range m{};
  if (info.Length() > 1 && info[1].IsObject()) {
    Napi::Object range = info[1].As<Napi::Object>();
    m.first_row = range.Get("firstRow").ToNumber().Uint32Value();
    m.last_row = range.Get("lastRow").ToNumber().Uint32Value();
    m.first_col = range.Get("firstCol").ToNumber().Uint32Value();
    m.last_col = range.Get("lastCol").ToNumber().Uint32Value();
  }
  fm_status_t rc = fm_sheet_add_merge(handle_, sheet, m);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::RemoveMerge(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  fm_merge_range m{};
  if (info.Length() > 1 && info[1].IsObject()) {
    Napi::Object range = info[1].As<Napi::Object>();
    m.first_row = range.Get("firstRow").ToNumber().Uint32Value();
    m.last_row = range.Get("lastRow").ToNumber().Uint32Value();
    m.first_col = range.Get("firstCol").ToNumber().Uint32Value();
    m.last_col = range.Get("lastCol").ToNumber().Uint32Value();
  }
  fm_status_t rc = fm_sheet_remove_merge(handle_, sheet, m);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::RemoveMergeAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t index = ArgU32(info, 1);
  fm_status_t rc = fm_sheet_remove_merge_at(handle_, sheet, index);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::ClearMerges(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  fm_status_t rc = fm_sheet_clear_merges(handle_, sheet);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::GetMerges(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return arr;
  }
  const uint32_t sheet = ArgU32(info, 0);
  uint32_t count = 0;
  if (fm_sheet_get_merge_count(handle_, sheet, &count) != 0) {
    return arr;
  }
  std::size_t emitted = 0;
  for (uint32_t i = 0; i < count; ++i) {
    fm_merge_range m{};
    if (fm_sheet_get_merge_at(handle_, sheet, i, &m) != 0) {
      continue;
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("firstRow", Napi::Number::New(env, m.first_row));
    item.Set("lastRow", Napi::Number::New(env, m.last_row));
    item.Set("firstCol", Napi::Number::New(env, m.first_col));
    item.Set("lastCol", Napi::Number::New(env, m.last_col));
    arr.Set(static_cast<uint32_t>(emitted), item);
    ++emitted;
  }
  return arr;
}

Napi::Value Workbook::GetComment(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return env.Null();
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  fm_comment c{};
  if (fm_sheet_get_comment_at(handle_, sheet, row, col, &c) != 0) {
    return env.Null();
  }
  Napi::Object o = Napi::Object::New(env);
  o.Set("author", Napi::String::New(env, c.author != nullptr ? c.author : ""));
  o.Set("text", Napi::String::New(env, c.text != nullptr ? c.text : ""));
  return o;
}

Napi::Value Workbook::SetComment(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  const std::string author = ArgString(info, 3);
  const std::string text = ArgString(info, 4);
  // Empty strings turn into NULL in the C ABI to remove an entry,
  // matching the embind binding's contract.
  const char* author_c = author.empty() ? nullptr : author.c_str();
  const char* text_c = text.empty() ? nullptr : text.c_str();
  fm_status_t rc = fm_sheet_set_comment(handle_, sheet, row, col, author_c, text_c);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::AddHyperlink(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  // Frontend signature: addHyperlink(sheet, row, col, target, display, tooltip).
  // The `location` field is omitted and forwarded as NULL; the writer
  // mints a fresh `rId` on save.
  if (info.Length() < 6 || !info[0].IsNumber() || !info[1].IsNumber() || !info[2].IsNumber() || !info[3].IsString() ||
      !info[4].IsString() || !info[5].IsString()) {
    Napi::TypeError::New(env,
                         "addHyperlink expects (sheet:number, row:number, col:number, "
                         "target:string, display:string, tooltip:string)")
        .ThrowAsJavaScriptException();
    return env.Undefined();
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  // Keep the std::string buffers alive until after the C ABI call so the
  // borrowed `const char*` pointers in `fm_hyperlink` stay valid.
  const std::string target = ArgString(info, 3);
  const std::string display = ArgString(info, 4);
  const std::string tooltip = ArgString(info, 5);
  fm_hyperlink hl{};
  hl.row = row;
  hl.col = col;
  hl.target = target.empty() ? nullptr : target.c_str();
  hl.location = nullptr;
  hl.display = display.empty() ? nullptr : display.c_str();
  hl.tooltip = tooltip.empty() ? nullptr : tooltip.c_str();
  fm_status_t rc = fm_sheet_add_hyperlink(handle_, sheet, hl);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::GetHyperlinks(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return arr;
  }
  const uint32_t sheet = ArgU32(info, 0);
  uint32_t count = 0;
  if (fm_sheet_get_hyperlink_count(handle_, sheet, &count) != 0) {
    return arr;
  }
  std::size_t emitted = 0;
  for (uint32_t i = 0; i < count; ++i) {
    fm_hyperlink h{};
    if (fm_sheet_get_hyperlink_at(handle_, sheet, i, &h) != 0) {
      continue;
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("row", Napi::Number::New(env, h.row));
    item.Set("col", Napi::Number::New(env, h.col));
    item.Set("target", Napi::String::New(env, h.target != nullptr ? h.target : ""));
    item.Set("display", Napi::String::New(env, h.display != nullptr ? h.display : ""));
    item.Set("tooltip", Napi::String::New(env, h.tooltip != nullptr ? h.tooltip : ""));
    arr.Set(static_cast<uint32_t>(emitted), item);
    ++emitted;
  }
  return arr;
}

Napi::Value Workbook::RemoveHyperlink(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  fm_status_t rc = fm_sheet_remove_hyperlink(handle_, sheet, row, col);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::RemoveHyperlinkAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t index = ArgU32(info, 1);
  fm_status_t rc = fm_sheet_remove_hyperlink_at(handle_, sheet, index);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::ClearHyperlinks(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  fm_status_t rc = fm_sheet_clear_hyperlinks(handle_, sheet);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::GetValidations(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return arr;
  }
  const uint32_t sheet = ArgU32(info, 0);
  uint32_t count = 0;
  if (fm_sheet_get_validation_count(handle_, sheet, &count) != 0) {
    return arr;
  }
  std::size_t emitted = 0;
  for (uint32_t i = 0; i < count; ++i) {
    fm_data_validation v{};
    if (fm_sheet_get_validation_at(handle_, sheet, i, &v) != 0) {
      continue;
    }
    Napi::Object item = Napi::Object::New(env);
    Napi::Array ranges = Napi::Array::New(env);
    for (uint32_t r = 0; r < v.range_count; ++r) {
      Napi::Object rng = Napi::Object::New(env);
      rng.Set("firstRow", Napi::Number::New(env, v.ranges[r].first_row));
      rng.Set("lastRow", Napi::Number::New(env, v.ranges[r].last_row));
      rng.Set("firstCol", Napi::Number::New(env, v.ranges[r].first_col));
      rng.Set("lastCol", Napi::Number::New(env, v.ranges[r].last_col));
      ranges.Set(r, rng);
    }
    item.Set("ranges", ranges);
    item.Set("type", Napi::Number::New(env, v.type));
    item.Set("op", Napi::Number::New(env, v.op));
    item.Set("errorStyle", Napi::Number::New(env, v.error_style));
    item.Set("allowBlank", Napi::Boolean::New(env, v.allow_blank != 0));
    item.Set("showInputMessage", Napi::Boolean::New(env, v.show_input_message != 0));
    item.Set("showErrorMessage", Napi::Boolean::New(env, v.show_error_message != 0));
    item.Set("formula1", Napi::String::New(env, v.formula1 != nullptr ? v.formula1 : ""));
    item.Set("formula2", Napi::String::New(env, v.formula2 != nullptr ? v.formula2 : ""));
    item.Set("errorTitle", Napi::String::New(env, v.error_title != nullptr ? v.error_title : ""));
    item.Set("errorMessage", Napi::String::New(env, v.error_message != nullptr ? v.error_message : ""));
    item.Set("promptTitle", Napi::String::New(env, v.prompt_title != nullptr ? v.prompt_title : ""));
    item.Set("promptMessage", Napi::String::New(env, v.prompt_message != nullptr ? v.prompt_message : ""));
    arr.Set(static_cast<uint32_t>(emitted), item);
    ++emitted;
  }
  return arr;
}

Napi::Value Workbook::AddValidation(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  if (info.Length() < 2 || !info[0].IsNumber() || !info[1].IsObject()) {
    Napi::TypeError::New(env, "addValidation expects (sheet:number, validation:object)").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  const uint32_t sheet = ArgU32(info, 0);
  Napi::Object v = info[1].As<Napi::Object>();

  // Pull every JS field into local storage first; the C ABI receives
  // borrowed `const char*` views that must stay valid until
  // `fm_sheet_add_validation` returns.
  std::vector<fm_merge_range> ranges_buf;
  if (v.Has("ranges")) {
    Napi::Value ranges_js = v.Get("ranges");
    if (ranges_js.IsArray()) {
      Napi::Array ranges_arr = ranges_js.As<Napi::Array>();
      const uint32_t n = ranges_arr.Length();
      ranges_buf.reserve(n);
      for (uint32_t i = 0; i < n; ++i) {
        Napi::Value rng_v = ranges_arr.Get(i);
        if (!rng_v.IsObject()) {
          continue;
        }
        Napi::Object rng = rng_v.As<Napi::Object>();
        fm_merge_range m{};
        m.first_row = rng.Get("firstRow").ToNumber().Uint32Value();
        m.last_row = rng.Get("lastRow").ToNumber().Uint32Value();
        m.first_col = rng.Get("firstCol").ToNumber().Uint32Value();
        m.last_col = rng.Get("lastCol").ToNumber().Uint32Value();
        ranges_buf.push_back(m);
      }
    }
  }
  auto pull_string = [&](const char* key) -> std::string {
    if (!v.Has(key)) {
      return std::string();
    }
    Napi::Value f = v.Get(key);
    if (f.IsUndefined() || f.IsNull()) {
      return std::string();
    }
    return f.ToString().Utf8Value();
  };
  auto pull_u8 = [&](const char* key) -> uint8_t {
    if (!v.Has(key)) {
      return 0;
    }
    Napi::Value f = v.Get(key);
    if (f.IsUndefined() || f.IsNull()) {
      return 0;
    }
    return static_cast<uint8_t>(f.ToNumber().Uint32Value() & 0xFFU);
  };
  auto pull_bool = [&](const char* key, bool dflt) -> bool {
    if (!v.Has(key)) {
      return dflt;
    }
    Napi::Value f = v.Get(key);
    if (f.IsUndefined() || f.IsNull()) {
      return dflt;
    }
    return f.ToBoolean().Value();
  };
  const std::string formula1 = pull_string("formula1");
  const std::string formula2 = pull_string("formula2");
  const std::string error_title = pull_string("errorTitle");
  const std::string error_message = pull_string("errorMessage");
  const std::string prompt_title = pull_string("promptTitle");
  const std::string prompt_message = pull_string("promptMessage");

  fm_data_validation dv{};
  dv.ranges = ranges_buf.empty() ? nullptr : ranges_buf.data();
  dv.range_count = static_cast<uint32_t>(ranges_buf.size());
  dv.type = pull_u8("type");
  dv.op = pull_u8("op");
  dv.error_style = pull_u8("errorStyle");
  dv.allow_blank = pull_bool("allowBlank", true) ? 1 : 0;
  dv.show_input_message = pull_bool("showInputMessage", false) ? 1 : 0;
  dv.show_error_message = pull_bool("showErrorMessage", false) ? 1 : 0;
  dv.formula1 = formula1.empty() ? nullptr : formula1.c_str();
  dv.formula2 = formula2.empty() ? nullptr : formula2.c_str();
  dv.error_title = error_title.empty() ? nullptr : error_title.c_str();
  dv.error_message = error_message.empty() ? nullptr : error_message.c_str();
  dv.prompt_title = prompt_title.empty() ? nullptr : prompt_title.c_str();
  dv.prompt_message = prompt_message.empty() ? nullptr : prompt_message.c_str();
  fm_status_t rc = fm_sheet_add_validation(handle_, sheet, dv);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::RemoveValidationAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t index = ArgU32(info, 1);
  fm_status_t rc = fm_sheet_remove_validation_at(handle_, sheet, index);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::ClearValidations(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  fm_status_t rc = fm_sheet_clear_validations(handle_, sheet);
  return MakeStatus(env, rc);
}

// ---- Class registration ---------------------------------------------

Napi::Function Workbook::GetClass(Napi::Env env) {
  return DefineClass(
      env, "Workbook",
      {
          StaticMethod<&Workbook::CreateDefault>("createDefault"),
          StaticMethod<&Workbook::CreateEmpty>("createEmpty"),
          StaticMethod<&Workbook::LoadBytes>("loadBytes"),
          InstanceMethod<&Workbook::AddBorder>("addBorder"),
          InstanceMethod<&Workbook::AddFill>("addFill"),
          InstanceMethod<&Workbook::AddFont>("addFont"),
          InstanceMethod<&Workbook::AddHyperlink>("addHyperlink"),
          InstanceMethod<&Workbook::AddMerge>("addMerge"),
          InstanceMethod<&Workbook::AddNumFmt>("addNumFmt"),
          InstanceMethod<&Workbook::AddSheet>("addSheet"),
          InstanceMethod<&Workbook::AddValidation>("addValidation"),
          InstanceMethod<&Workbook::AddXf>("addXf"),
          InstanceMethod<&Workbook::BorderCount>("borderCount"),
          InstanceMethod<&Workbook::CellAt>("cellAt"),
          InstanceMethod<&Workbook::CellCount>("cellCount"),
          InstanceMethod<&Workbook::ClearHyperlinks>("clearHyperlinks"),
          InstanceMethod<&Workbook::ClearMerges>("clearMerges"),
          InstanceMethod<&Workbook::ClearValidations>("clearValidations"),
          InstanceMethod<&Workbook::DefinedNameAt>("definedNameAt"),
          InstanceMethod<&Workbook::DefinedNameCount>("definedNameCount"),
          InstanceMethod<&Workbook::DeleteCols>("deleteCols"),
          InstanceMethod<&Workbook::DeleteRows>("deleteRows"),
          InstanceMethod<&Workbook::EvaluateCfRange>("evaluateCfRange"),
          InstanceMethod<&Workbook::FillCount>("fillCount"),
          InstanceMethod<&Workbook::FontCount>("fontCount"),
          InstanceMethod<&Workbook::GetBorder>("getBorder"),
          InstanceMethod<&Workbook::GetCellXf>("getCellXf"),
          InstanceMethod<&Workbook::GetCellXfIndex>("getCellXfIndex"),
          InstanceMethod<&Workbook::GetComment>("getComment"),
          InstanceMethod<&Workbook::GetFill>("getFill"),
          InstanceMethod<&Workbook::GetFont>("getFont"),
          InstanceMethod<&Workbook::GetHyperlinks>("getHyperlinks"),
          InstanceMethod<&Workbook::GetMerges>("getMerges"),
          InstanceMethod<&Workbook::GetNumFmt>("getNumFmt"),
          InstanceMethod<&Workbook::GetSheetColumns>("getSheetColumns"),
          InstanceMethod<&Workbook::GetSheetRowOverrides>("getSheetRowOverrides"),
          InstanceMethod<&Workbook::GetSheetView>("getSheetView"),
          InstanceMethod<&Workbook::GetValidations>("getValidations"),
          InstanceMethod<&Workbook::GetValue>("getValue"),
          InstanceMethod<&Workbook::InsertCols>("insertCols"),
          InstanceMethod<&Workbook::InsertRows>("insertRows"),
          InstanceMethod<&Workbook::IsValid>("isValid"),
          InstanceMethod<&Workbook::MoveSheet>("moveSheet"),
          InstanceMethod<&Workbook::PartialRecalc>("partialRecalc"),
          InstanceMethod<&Workbook::PassthroughAt>("passthroughAt"),
          InstanceMethod<&Workbook::PassthroughCount>("passthroughCount"),
          InstanceMethod<&Workbook::PivotCacheCount>("pivotCacheCount"),
          InstanceMethod<&Workbook::PivotCacheCreate>("pivotCacheCreate"),
          InstanceMethod<&Workbook::PivotCacheFieldAdd>("pivotCacheFieldAdd"),
          InstanceMethod<&Workbook::PivotCacheFieldAddSharedItemBlank>("pivotCacheFieldAddSharedItemBlank"),
          InstanceMethod<&Workbook::PivotCacheFieldAddSharedItemBool>("pivotCacheFieldAddSharedItemBool"),
          InstanceMethod<&Workbook::PivotCacheFieldAddSharedItemNumber>("pivotCacheFieldAddSharedItemNumber"),
          InstanceMethod<&Workbook::PivotCacheFieldAddSharedItemText>("pivotCacheFieldAddSharedItemText"),
          InstanceMethod<&Workbook::PivotCacheFieldClear>("pivotCacheFieldClear"),
          InstanceMethod<&Workbook::PivotCacheFieldClearSharedItems>("pivotCacheFieldClearSharedItems"),
          InstanceMethod<&Workbook::PivotCacheFieldCount>("pivotCacheFieldCount"),
          InstanceMethod<&Workbook::PivotCacheFieldName>("pivotCacheFieldName"),
          InstanceMethod<&Workbook::PivotCacheFieldSharedItemCount>("pivotCacheFieldSharedItemCount"),
          InstanceMethod<&Workbook::PivotCacheIdAt>("pivotCacheIdAt"),
          InstanceMethod<&Workbook::PivotCacheRecordAdd>("pivotCacheRecordAdd"),
          InstanceMethod<&Workbook::PivotCacheRecordClear>("pivotCacheRecordClear"),
          InstanceMethod<&Workbook::PivotCacheRecordCount>("pivotCacheRecordCount"),
          InstanceMethod<&Workbook::PivotCacheRecordSetBlank>("pivotCacheRecordSetBlank"),
          InstanceMethod<&Workbook::PivotCacheRecordSetBool>("pivotCacheRecordSetBool"),
          InstanceMethod<&Workbook::PivotCacheRecordSetError>("pivotCacheRecordSetError"),
          InstanceMethod<&Workbook::PivotCacheRecordSetNumber>("pivotCacheRecordSetNumber"),
          InstanceMethod<&Workbook::PivotCacheRecordSetText>("pivotCacheRecordSetText"),
          InstanceMethod<&Workbook::PivotCacheRemove>("pivotCacheRemove"),
          InstanceMethod<&Workbook::PivotCount>("pivotCount"),
          InstanceMethod<&Workbook::PivotCreate>("pivotCreate"),
          InstanceMethod<&Workbook::PivotDataFieldAdd>("pivotDataFieldAdd"),
          InstanceMethod<&Workbook::PivotDataFieldClear>("pivotDataFieldClear"),
          InstanceMethod<&Workbook::PivotDataFieldCount>("pivotDataFieldCount"),
          InstanceMethod<&Workbook::PivotDataFieldSet>("pivotDataFieldSet"),
          InstanceMethod<&Workbook::PivotFieldAdd>("pivotFieldAdd"),
          InstanceMethod<&Workbook::PivotFieldAddAggregation>("pivotFieldAddAggregation"),
          InstanceMethod<&Workbook::PivotFieldAddItem>("pivotFieldAddItem"),
          InstanceMethod<&Workbook::PivotFieldAddSubtotalFn>("pivotFieldAddSubtotalFn"),
          InstanceMethod<&Workbook::PivotFieldClear>("pivotFieldClear"),
          InstanceMethod<&Workbook::PivotFieldClearAggregations>("pivotFieldClearAggregations"),
          InstanceMethod<&Workbook::PivotFieldClearDateGroup>("pivotFieldClearDateGroup"),
          InstanceMethod<&Workbook::PivotFieldClearItems>("pivotFieldClearItems"),
          InstanceMethod<&Workbook::PivotFieldClearSubtotalFns>("pivotFieldClearSubtotalFns"),
          InstanceMethod<&Workbook::PivotFieldCount>("pivotFieldCount"),
          InstanceMethod<&Workbook::PivotFieldSetAxis>("pivotFieldSetAxis"),
          InstanceMethod<&Workbook::PivotFieldSetDateGroup>("pivotFieldSetDateGroup"),
          InstanceMethod<&Workbook::PivotFieldSetItemVisible>("pivotFieldSetItemVisible"),
          InstanceMethod<&Workbook::PivotFieldSetNumberFormat>("pivotFieldSetNumberFormat"),
          InstanceMethod<&Workbook::PivotFieldSetSort>("pivotFieldSetSort"),
          InstanceMethod<&Workbook::PivotFieldSetSubtotalTop>("pivotFieldSetSubtotalTop"),
          InstanceMethod<&Workbook::PivotFilterAdd>("pivotFilterAdd"),
          InstanceMethod<&Workbook::PivotFilterClear>("pivotFilterClear"),
          InstanceMethod<&Workbook::PivotFilterCount>("pivotFilterCount"),
          InstanceMethod<&Workbook::PivotFilterRemoveAt>("pivotFilterRemoveAt"),
          InstanceMethod<&Workbook::PivotLayout>("pivotLayout"),
          InstanceMethod<&Workbook::PivotRemove>("pivotRemove"),
          InstanceMethod<&Workbook::PivotSetAnchor>("pivotSetAnchor"),
          InstanceMethod<&Workbook::PivotSetColFieldOrder>("pivotSetColFieldOrder"),
          InstanceMethod<&Workbook::PivotSetGrandTotals>("pivotSetGrandTotals"),
          InstanceMethod<&Workbook::PivotSetName>("pivotSetName"),
          InstanceMethod<&Workbook::PivotSetRowFieldOrder>("pivotSetRowFieldOrder"),
          InstanceMethod<&Workbook::Recalc>("recalc"),
          InstanceMethod<&Workbook::RemoveHyperlink>("removeHyperlink"),
          InstanceMethod<&Workbook::RemoveHyperlinkAt>("removeHyperlinkAt"),
          InstanceMethod<&Workbook::RemoveMerge>("removeMerge"),
          InstanceMethod<&Workbook::RemoveMergeAt>("removeMergeAt"),
          InstanceMethod<&Workbook::RemoveSheet>("removeSheet"),
          InstanceMethod<&Workbook::RemoveValidationAt>("removeValidationAt"),
          InstanceMethod<&Workbook::RenameSheet>("renameSheet"),
          InstanceMethod<&Workbook::Save>("save"),
          InstanceMethod<&Workbook::SetBlank>("setBlank"),
          InstanceMethod<&Workbook::SetBool>("setBool"),
          InstanceMethod<&Workbook::SetCellXfIndex>("setCellXfIndex"),
          InstanceMethod<&Workbook::SetColumnHidden>("setColumnHidden"),
          InstanceMethod<&Workbook::SetColumnOutline>("setColumnOutline"),
          InstanceMethod<&Workbook::SetColumnWidth>("setColumnWidth"),
          InstanceMethod<&Workbook::SetComment>("setComment"),
          InstanceMethod<&Workbook::SetDefinedName>("setDefinedName"),
          InstanceMethod<&Workbook::SetFormula>("setFormula"),
          InstanceMethod<&Workbook::SetIterative>("setIterative"),
          InstanceMethod<&Workbook::SetIterativeProgress>("setIterativeProgress"),
          InstanceMethod<&Workbook::SetNumber>("setNumber"),
          InstanceMethod<&Workbook::SetRowHeight>("setRowHeight"),
          InstanceMethod<&Workbook::SetRowHidden>("setRowHidden"),
          InstanceMethod<&Workbook::SetRowOutline>("setRowOutline"),
          InstanceMethod<&Workbook::SetSheetFreeze>("setSheetFreeze"),
          InstanceMethod<&Workbook::SetSheetTabHidden>("setSheetTabHidden"),
          InstanceMethod<&Workbook::SetSheetZoom>("setSheetZoom"),
          InstanceMethod<&Workbook::SetText>("setText"),
          InstanceMethod<&Workbook::SheetCount>("sheetCount"),
          InstanceMethod<&Workbook::SheetName>("sheetName"),
          InstanceMethod<&Workbook::TableAt>("tableAt"),
          InstanceMethod<&Workbook::TableCount>("tableCount"),
          InstanceMethod<&Workbook::XfCount>("xfCount"),
      });
}

// ---------------------------------------------------------------------
// Free functions
// ---------------------------------------------------------------------

/// `evalFormula(formula)`: convenience that mirrors the embind variant.
/// Spins up an empty workbook, places the formula at A1, recalcs, and
/// returns `{ status, value }`.
Napi::Value EvalFormula(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  std::string formula;
  if (info.Length() > 0) {
    formula = info[0].ToString().Utf8Value();
  }

  fm_workbook_t* wb = nullptr;
  fm_status_t rc = fm_workbook_create(&wb);
  if (rc != 0) {
    return MakeEmptyValueResult(env, MakeErrorStatus(env, rc));
  }
  rc = fm_workbook_set_formula(wb, 0, 0, 0, formula.c_str());
  if (rc != 0) {
    Napi::Object out = MakeEmptyValueResult(env, MakeErrorStatus(env, rc));
    fm_workbook_destroy(wb);
    return out;
  }
  rc = fm_workbook_recalc(wb);
  if (rc != 0) {
    Napi::Object out = MakeEmptyValueResult(env, MakeErrorStatus(env, rc));
    fm_workbook_destroy(wb);
    return out;
  }
  fm_value_t v{};
  rc = fm_workbook_get_value(wb, 0, 0, 0, &v);
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

// ---------------------------------------------------------------------
// Module init
// ---------------------------------------------------------------------

Napi::Object Init(Napi::Env env, Napi::Object exports) {
  exports.Set("Workbook", Workbook::GetClass(env));
  exports.Set("evalFormula", Napi::Function::New(env, &EvalFormula, "evalFormula"));
  exports.Set("version", Napi::Function::New(env, &Version, "version"));
  exports.Set("lastErrorMessage", Napi::Function::New(env, &LastErrorMessage, "lastErrorMessage"));
  exports.Set("lastErrorContext", Napi::Function::New(env, &LastErrorContext, "lastErrorContext"));
  exports.Set("statusString", Napi::Function::New(env, &StatusString, "statusString"));
  return exports;
}

}  // namespace

NODE_API_MODULE(formulon, Init)
