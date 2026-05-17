// Copyright 2026 libraz. Licensed under the MIT License.
//
// Node.js N-API addon entry point for the Formulon engine.
//
// This TU is intentionally tiny: it wires the per-area implementation
// units under `src/node_addon/parts/` into the module exports and
// declares the `NODE_API_MODULE` macro. The bindings themselves live
// in:
//
//   * `parts/addon_common.{h,cc}`  -- shared translation helpers,
//     module-global iterative-progress slot, JS-spec pullers.
//   * `parts/workbook_class.{h,cc}` -- `Workbook` ObjectWrap definition,
//     ctor / dtor, static factories, argument helpers, and the
//     `DefineClass` registration table.
//   * `parts/lifecycle.cc`         -- cell mutation / read, recalc,
//     save, iterative-solver registration, isValid.
//   * `parts/sheet.cc`             -- sheet operations, row/col edits,
//     metadata iteration, defined-name mutation.
//   * `parts/pivot_cache.cc`       -- PivotCache mutation surface.
//   * `parts/pivot_table.cc`       -- PivotTable mutation surface.
//   * `parts/styles.cc`            -- styles + EvaluateCfRange.
//   * `parts/sheet_view.cc`        -- view / column / row layout +
//     merges / comments / hyperlinks / validations.
//   * `parts/free_funcs.{h,cc}`    -- evalFormula / version /
//     lastError* / statusString.
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

#include "node_addon/parts/free_funcs.h"
#include "node_addon/parts/workbook_class.h"

namespace {

Napi::Object Init(Napi::Env env, Napi::Object exports) {
  exports.Set("Workbook", formulon_node::Workbook::GetClass(env));
  exports.Set("evalFormula", Napi::Function::New(env, &formulon_node::EvalFormula, "evalFormula"));
  exports.Set("version", Napi::Function::New(env, &formulon_node::Version, "version"));
  exports.Set("lastErrorMessage", Napi::Function::New(env, &formulon_node::LastErrorMessage, "lastErrorMessage"));
  exports.Set("lastErrorContext", Napi::Function::New(env, &formulon_node::LastErrorContext, "lastErrorContext"));
  exports.Set("statusString", Napi::Function::New(env, &formulon_node::StatusString, "statusString"));
  return exports;
}

}  // namespace

NODE_API_MODULE(formulon, Init)
