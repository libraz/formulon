// Free (non-method) bindings exposed on the module exports object:
// `evalFormula`, `version` (aliased as `versionString` to match the WASM
// binding), `lastErrorMessage`, `lastErrorContext`, and `statusString`.
// Declared here so `addon.cc::Init` can attach them without leaking the
// per-area TU layout.

#ifndef FORMULON_NODE_ADDON_PARTS_FREE_FUNCS_H_
#define FORMULON_NODE_ADDON_PARTS_FREE_FUNCS_H_

#include "node_addon/parts/addon_common.h"

namespace formulon_node {

/// `evalFormula(formula)`: convenience that mirrors the embind variant.
/// Spins up an empty workbook, places the formula at A1, recalcs, and
/// returns `{ status, value }`.
Napi::Value EvalFormula(const Napi::CallbackInfo& info);

/// Returns the engine version string (`fm_version_string`).
Napi::Value Version(const Napi::CallbackInfo& info);

/// Returns the thread-local last-error message snapshot.
Napi::Value LastErrorMessage(const Napi::CallbackInfo& info);

/// Returns the thread-local last-error context snapshot.
Napi::Value LastErrorContext(const Napi::CallbackInfo& info);

/// Returns the human-readable description of a numeric status code.
Napi::Value StatusString(const Napi::CallbackInfo& info);

}  // namespace formulon_node

#endif  // FORMULON_NODE_ADDON_PARTS_FREE_FUNCS_H_
