// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Registers Excel's host-service built-ins as deterministic stubs:
//
//     * IMAGE(source, [alt_text], [sizing], [height], [width]) -> #VALUE!
//     * RTD(progID, server, topic1, [topic2], ...)             -> #N/A
//     * TRANSLATE(text, source_lang, target_lang)              -> #NAME?
//     * DETECTLANGUAGE(text)                                    -> #NAME?
//     * COPILOT(prompt, [context], ...)                         -> #NAME?
//
// Each stub requires a runtime capability a pure calculation engine
// does not provide -- image rendering, an RTD server, a translation
// service, or a Copilot backend. All stubs ride the eager dispatcher
// (so error args short-circuit before the fixed return fires). Their
// presence keeps the catalog complete and surfaces a deterministic
// Excel-visible result rather than `#NAME?` from an unknown-function
// lookup.
//
// GETPIVOTDATA is intentionally NOT registered here. It is a real lazy
// form (see `eval_getpivotdata_lazy` in `getpivotdata_lazy.cpp`) that
// recovers the (sheet, row, col) anchor of the second argument from the
// un-evaluated Reference AST and walks the resolved `PivotTable` /
// `PivotResult` to return the addressed cell. The eager registry would
// flatten the anchor to a Value before the impl could see the
// reference, so GETPIVOTDATA is wired through `tree_walker.cpp`'s lazy
// dispatch table.
//
// PHONETIC is intentionally NOT registered here. It is a real lazy form
// (see `eval_phonetic_lazy` in `phonetic_lazy.cpp`) that reads the
// referenced cell's `<rPh>` annotation directly off the un-evaluated
// Reference AST. The eager registry would flatten the argument to a
// Value before the impl could consult `Cell::phonetic_text`, so PHONETIC
// is wired through `tree_walker.cpp`'s lazy dispatch table.
//
// ISOMITTED is intentionally NOT registered here. It is a real lazy
// special form (see `eval_isomitted_lazy` in `special_forms_lazy.cpp`)
// that inspects the argument's AST shape and the active `NameEnv`'s
// omitted flag. The eager registry cannot see either, so ISOMITTED is
// wired through `tree_walker.cpp`'s lazy dispatch table.

#ifndef FORMULON_EVAL_BUILTINS_SERVICE_STUBS_H_
#define FORMULON_EVAL_BUILTINS_SERVICE_STUBS_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the service-stub built-ins (IMAGE, RTD, TRANSLATE,
/// DETECTLANGUAGE, COPILOT) into `registry`. Intended to be invoked
/// from `register_builtins`. GETPIVOTDATA, PHONETIC, and ISOMITTED are
/// registered separately as lazy forms; see the file-level comment.
void register_service_stub_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_SERVICE_STUBS_H_
