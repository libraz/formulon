// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Formulon's host-service and no-infrastructure built-in
// stubs. Two groups live together because both classes return a deterministic
// Excel-visible surface in lieu of a feature the engine does not yet provide:
//
// Host-service stubs (require an external runtime Formulon does not embed):
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
// No-infrastructure stubs (the supporting Formulon subsystem isn't built
// out yet, so any answer the function might give is structurally absent):
//
//   * PHONETIC        -> input text passthrough / #N/A on non-text.
//                        Mac Excel reads the IME-typed kana from the
//                        OOXML <rPh> annotation block and falls back to
//                        the surface text when no annotation is present.
//                        Formulon does not yet parse <rPh> nor surface
//                        phonetic metadata through `Value`, so every
//                        cell is "unannotated" by construction and we
//                        always return the input text. Non-text inputs
//                        return #N/A to match Mac's strict-text rule.
//   * GETPIVOTDATA    -> fixed #REF!   (no pivot tables yet, so no
//                        field/item lookup can ever resolve; #REF!
//                        matches Mac when the lookup target is invalid).
//
// All host-service stubs and PHONETIC / GETPIVOTDATA ride the eager dispatch
// path (`accepts_ranges = false`, default `propagate_errors = true`) so an
// error argument short-circuits before the fixed return fires -- this matches
// the WEBSERVICE / PY stubs in `src/eval/builtins/web.cpp` and keeps the
// surface consistent with real functions for formulas that propagate errors
// through service calls.
//
// ISOMITTED is NOT a stub. It is a real lazy special form implemented in
// `eval_isomitted_lazy` (see `special_forms_lazy.cpp`) and registered through
// `tree_walker.cpp`'s lazy dispatch table. The eager registry cannot see the
// argument's AST shape or the active `NameEnv`'s omitted flag, both of which
// the implementation needs.

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

// PHONETIC reads the IME-typed kana from the OOXML <rPh> annotation
// block on the cell's rich text. Formulon's OOXML reader does not yet
// surface <rPh> annotations into the `Value` model, so every cell is
// effectively unannotated. Mac Excel returns the surface text unchanged
// for unannotated text cells, returns #N/A for non-text references
// (numbers, bools, dates), and returns the empty string for blanks.
// We match that surface today; once <rPh> is parsed and threaded
// through the value model, this stub gets replaced with a real lookup
// against the annotation table.
Value Phonetic(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  const Value& v = args[0];
  if (v.is_text()) {
    return v;
  }
  if (v.is_blank()) {
    return Value::text(arena.intern(""));
  }
  // Number / Bool / Date / Array / Ref / Lambda all surface #N/A on
  // Mac, matching the strict-text rule. Errors are short-circuited by
  // the dispatcher (`propagate_errors = true`) and never reach here.
  return Value::error(ErrorCode::NA);
}

// GETPIVOTDATA looks up a measure value inside a pivot table by
// field/item key. Formulon does not yet implement pivot tables, so no
// lookup can ever resolve, and #REF! is the surface Mac Excel returns
// when the field/item path doesn't address a valid pivot cell. The
// dispatcher's eager error propagation means an error in any argument
// (data field, pivot anchor, field/item slots) short-circuits before
// this body runs.
Value GetPivotData(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::error(ErrorCode::Ref);
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
  // PHONETIC(reference) -- exact arity 1. Default `propagate_errors = true`
  // so an error argument short-circuits before the text/non-text branch.
  registry.register_function(FunctionDef{"PHONETIC", 1u, 1u, &Phonetic});
  // GETPIVOTDATA(data_field, pivot_table, [field1, item1, ...]) -- min 2
  // (data_field + pivot anchor required), max kVariadic (field/item pairs
  // are optional and arbitrary in count). Default `propagate_errors = true`
  // so an error in any argument surfaces instead of the fixed #REF!.
  registry.register_function(FunctionDef{"GETPIVOTDATA", 2u, kVariadic, &GetPivotData});
  // ISOMITTED(argument) is registered through the lazy dispatch table in
  // `tree_walker.cpp` (entry: ISOMITTED -> eval_isomitted_lazy). The lazy
  // path inspects the argument's AST shape and the active `NameEnv`'s
  // omitted flag, which the eager registry cannot see.
}

}  // namespace eval
}  // namespace formulon
