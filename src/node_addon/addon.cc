// Node.js N-API addon entry point for the Formulon engine.
//
// This TU is intentionally tiny: it wires the per-area implementation
// units under `src/node_addon/parts/` into the module exports and
// declares the `NODE_API_MODULE` macro. The bindings themselves live
// in:
//
//   * `parts/addon_common.{h,cc}`  -- shared translation helpers and
//     JS-spec pullers.
//   * `parts/workbook_class.{h,cc}` -- `Workbook` ObjectWrap definition,
//     ctor / dtor, static factories, argument helpers, and the
//     `DefineClass` registration table.
//   * `parts/cells.cc`             -- cell mutation / read, phonetic
//     metadata, ad-hoc formula evaluation, lambda text.
//   * `parts/lifecycle.cc`         -- calc policy, recalc, save,
//     iterative-solver registration, diagnostics, `disposed`.
//   * `parts/sheet.cc`             -- sheet operations, row/col edits,
//     metadata iteration, defined-name mutation.
//   * `parts/pivot_cache.cc`       -- PivotCache mutation surface.
//   * `parts/pivot_table.cc`       -- PivotTable mutation surface.
//   * `parts/styles.cc`            -- styles.
//   * `parts/conditional_format.cc` -- conditional-format surface +
//     EvaluateCfRange.
//   * `parts/sheet_view.cc`        -- view / protection / column / row
//     layout.
//   * `parts/sheet_ui.cc`         -- merges / comments / hyperlinks /
//     validations / threaded comments and persons.
//   * `parts/auto_filter.cc`      -- typed AutoFilter for sheets and
//     tables.
//   * `parts/drawing.cc`          -- drawing images.
//   * `parts/reference.cc`         -- parseReference / formatReference /
//     getSheetIndex.
//   * `parts/free_funcs.{h,cc}`    -- evalFormula / getVersion /
//     getStatusName / getErrorDisplayName / the process-wide
//     structured-log controls / `__setErrorClass`.
//
// ## Design notes
//
//   * The whole engine is built `-fno-exceptions -fno-rtti`. node-addon-api
//     would normally throw C++ exceptions across the boundary; we
//     disable that via `NODE_ADDON_API_DISABLE_CPP_EXCEPTIONS` (and
//     `NAPI_DISABLE_CPP_EXCEPTIONS`) before any napi header include,
//     which the FormulonNodeAddon.cmake target propagates.
//
//   * A binding entry returns its payload on success, the same JS shape
//     as the embind binding. A failed C-ABI call throws the `FormulonError`
//     the package loader injected through `__setErrorClass`, carrying the
//     status code and the thread-local `fm_last_error_*` message and
//     context. A `Value` carries `kind` and only that kind's field.
//
//   * An exception a JS callback throws stays pending: the C call is
//     aborted and the outer entry returns that exception unchanged.
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
//   * `setIterativeProgress` keeps its `Napi::FunctionReference` on the
//     owning Workbook and passes that wrapper as C ABI user-data. The
//     solver invokes it synchronously inside `recalc()`, so no
//     thread-safe-function plumbing is required.
//
//   * `setLogSink` cannot make that assumption: the engine may emit a
//     record from any thread it owns, so the sink goes through a
//     `Napi::ThreadSafeFunction` and is delivered asynchronously. It is
//     process-wide state, hence a module-level function rather than a
//     `Workbook` method.

#include "node_addon/parts/free_funcs.h"
#include "node_addon/parts/workbook_class.h"

namespace {

Napi::Object Init(Napi::Env env, Napi::Object exports) {
  exports.Set("Workbook", formulon_node::Workbook::GetClass(env));
  exports.Set("evalFormula", Napi::Function::New(env, &formulon_node::EvalFormula, "evalFormula"));
  exports.Set("getVersion", Napi::Function::New(env, &formulon_node::Version, "getVersion"));
  exports.Set("getStatusName", Napi::Function::New(env, &formulon_node::StatusString, "getStatusName"));
  exports.Set("getErrorDisplayName", Napi::Function::New(env, &formulon_node::ErrorDisplayName, "getErrorDisplayName"));
  exports.Set("setLogMinLevel", Napi::Function::New(env, &formulon_node::SetLogMinLevel, "setLogMinLevel"));
  exports.Set("setLogSink", Napi::Function::New(env, &formulon_node::SetLogSink, "setLogSink"));
  exports.Set("__setErrorClass", Napi::Function::New(env, &formulon_node::SetErrorClass, "__setErrorClass"));
  return exports;
}

}  // namespace

NODE_API_MODULE(formulon, Init)
