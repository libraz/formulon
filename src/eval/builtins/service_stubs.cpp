// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Formulon's host-service built-in stubs:
//
//   * IMAGE           -> fixed #VALUE! (inline image embedding requires a
//                        rendering host; Formulon is a calc engine).
//   * RTD             -> fixed #N/A    (Real-Time-Data server integration
//                        is out of scope; #N/A matches Excel's behaviour
//                        when the RTD provider is unavailable).
//   * TRANSLATE       -> fixed #NAME?  (translation service is out of
//                        scope; #NAME? matches the function-unknown
//                        surface Excel emits when the cloud translate
//                        service is offline).
//   * DETECTLANGUAGE  -> fixed #NAME?  (language-detection service is out
//                        of scope; same rationale as TRANSLATE).
//   * COPILOT         -> fixed #NAME?  (Copilot integration requires an
//                        external API; #NAME? matches the offline-service
//                        surface).
//
// All five stubs ride the eager dispatch path (`accepts_ranges = false`,
// default `propagate_errors = true`) so an error argument short-circuits
// before the fixed return fires -- this matches the WEBSERVICE / PY
// stubs in `src/eval/builtins/web.cpp` and keeps the surface consistent
// with real functions for formulas that propagate errors through
// service calls.

#include "eval/builtins/service_stubs.h"

#include <cstdint>

#include "eval/function_registry.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Inline image embedding requires a rendering host; Formulon is a pure
// calculation engine, so we always return #VALUE!. Same pattern as
// WEBSERVICE in `src/eval/builtins/web.cpp`.
Value Image(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::error(ErrorCode::Value);
}

// RTD dispatches to an external Real-Time-Data COM server; Formulon has
// no provider runtime, so the stub returns #N/A. Excel itself emits
// #N/A when the RTD server is unreachable, so this matches the offline
// surface for formulas that gracefully handle a missing feed.
Value Rtd(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::error(ErrorCode::NA);
}

// Cloud translation service; Formulon does no network I/O, so the stub
// returns #NAME? to mirror Excel's own behaviour when the translate
// service is unavailable (function-unknown surface). Same pattern as PY
// in `src/eval/builtins/web.cpp`.
Value Translate(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::error(ErrorCode::Name);
}

// Language-detection service shares Translate's transport, so the stub
// surfaces #NAME? for the same reason.
Value DetectLanguage(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::error(ErrorCode::Name);
}

// Copilot dispatches prompts to an external LLM backend; Formulon does
// not embed one. #NAME? matches Excel's offline-service behaviour.
Value Copilot(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::error(ErrorCode::Name);
}

}  // namespace

void register_service_stub_builtins(FunctionRegistry& registry) {
  // IMAGE(source, [alt_text], [sizing], [height], [width]) -- arity 1..5.
  registry.register_function(FunctionDef{"IMAGE", 1u, 5u, &Image});
  // RTD(progID, server, topic1, [topic2], ...) -- Excel caps total args
  // at 255 after the two required scalars (the usual variadic ceiling
  // inside the engine; kVariadic is the explicit sentinel for no cap).
  registry.register_function(FunctionDef{"RTD", 3u, 255u, &Rtd});
  // TRANSLATE(text, source_lang, target_lang) -- Mac Excel 365 accepts
  // the source language as optional (auto-detect), so min_arity = 2.
  registry.register_function(FunctionDef{"TRANSLATE", 2u, 3u, &Translate});
  registry.register_function(FunctionDef{"DETECTLANGUAGE", 1u, 1u, &DetectLanguage});
  // COPILOT(prompt, [context...]) -- variadic context cells.
  registry.register_function(FunctionDef{"COPILOT", 1u, 255u, &Copilot});
}

}  // namespace eval
}  // namespace formulon
