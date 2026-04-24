// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's host-service built-ins as deterministic stubs:
//
//   * IMAGE(source, [alt_text], [sizing], [height], [width]) -> #VALUE!
//   * RTD(progID, server, topic1, [topic2], ...)             -> #N/A
//   * TRANSLATE(text, source_lang, target_lang)              -> #NAME?
//   * DETECTLANGUAGE(text)                                    -> #NAME?
//   * COPILOT(prompt, [context], ...)                         -> #NAME?
//
// These functions each require a runtime capability a pure calculation
// engine does not provide — image rendering, an RTD server, a
// translation service, or a Copilot backend. The stubs mirror the
// WEBSERVICE / PY pattern in `src/eval/builtins/web.cpp`: arguments are
// still evaluated via the eager dispatcher so error args short-circuit
// normally before the fixed return fires. Their presence keeps the
// catalog complete and surfaces a deterministic Excel-visible error
// rather than `#NAME?` from an unknown-function lookup.

#ifndef FORMULON_EVAL_BUILTINS_SERVICE_STUBS_H_
#define FORMULON_EVAL_BUILTINS_SERVICE_STUBS_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the service-stub built-ins (IMAGE, RTD, TRANSLATE,
/// DETECTLANGUAGE, COPILOT) into `registry`. Intended to be invoked
/// from `register_builtins`.
void register_service_stub_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_SERVICE_STUBS_H_
