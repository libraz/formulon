// Free (non-method) bindings exposed on the module exports object:
// `evalFormula`, `version` (aliased as `versionString` to match the WASM
// binding), `lastErrorMessage`, `lastErrorContext`, `statusString`,
// `errorDisplayName`, `setLogMinLevel` and `setLogSink`.
// Declared here so `addon.cc::Init` can attach them without leaking the
// per-area TU layout.

#ifndef FORMULON_NODE_ADDON_PARTS_FREE_FUNCS_H_
#define FORMULON_NODE_ADDON_PARTS_FREE_FUNCS_H_

#include "node_addon/parts/addon_common.h"

namespace formulon_node {

/// `evalFormula(formula)`: convenience that mirrors the embind variant.
/// Evaluates `formula` read-only against a fresh single-sheet workbook,
/// anchored at `Sheet1!A1`, and returns `{ status, value }`. Nothing is
/// written and no recalc runs, so an anchor-referencing formula such as
/// `=A1` reads the (blank) anchor rather than becoming a self-reference.
/// An array result is reduced to its top-left element.
Napi::Value EvalFormula(const Napi::CallbackInfo& info);

/// Returns the engine version string (`fm_version_string`).
Napi::Value Version(const Napi::CallbackInfo& info);

/// Returns the thread-local last-error message snapshot.
Napi::Value LastErrorMessage(const Napi::CallbackInfo& info);

/// Returns the thread-local last-error context snapshot.
Napi::Value LastErrorContext(const Napi::CallbackInfo& info);

/// Returns the human-readable description of a numeric status code.
Napi::Value StatusString(const Napi::CallbackInfo& info);

/// Returns an Excel literal such as `#DIV/0!` for a numeric cell error code.
Napi::Value ErrorDisplayName(const Napi::CallbackInfo& info);

/// `setLogMinLevel(level)`: sets the engine's minimum structured-log
/// severity. Process-wide, not per workbook handle. The default is
/// `FM_LOG_LEVEL_OFF` (4), under which the engine writes nothing.
Napi::Value SetLogMinLevel(const Napi::CallbackInfo& info);

/// `setLogSink(cb | null)`: routes structured-log records to `cb`, or
/// restores the (silent at the default threshold) stderr fallback.
/// Process-wide, not per workbook handle.
Napi::Value SetLogSink(const Napi::CallbackInfo& info);

}  // namespace formulon_node

#endif  // FORMULON_NODE_ADDON_PARTS_FREE_FUNCS_H_
