// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's host-service and no-infrastructure built-ins as
// deterministic stubs:
//
//   Host-service stubs (require an external runtime Formulon does not embed):
//
//     * IMAGE(source, [alt_text], [sizing], [height], [width]) -> #VALUE!
//     * RTD(progID, server, topic1, [topic2], ...)             -> #N/A
//     * TRANSLATE(text, source_lang, target_lang)              -> #NAME?
//     * DETECTLANGUAGE(text)                                    -> #NAME?
//     * COPILOT(prompt, [context], ...)                         -> #NAME?
//
//   No-infrastructure stubs (the supporting subsystem isn't built yet):
//
//     * PHONETIC(reference)                                     -> input
//                                                                  text /
//                                                                  #N/A on
//                                                                  non-text
//     * GETPIVOTDATA(data_field, pivot_table, [field1, item1, ...]) -> #REF!
//
// The host-service stubs each require a runtime capability a pure
// calculation engine does not provide — image rendering, an RTD server,
// a translation service, or a Copilot backend. The no-infrastructure
// stubs each require a Formulon subsystem that is not yet built out —
// OOXML <rPh> annotation parsing for PHONETIC and pivot tables for
// GETPIVOTDATA. All stubs ride the eager dispatcher (so error args
// short-circuit before the fixed return fires). Their presence keeps
// the catalog complete and surfaces a deterministic Excel-visible
// result rather than `#NAME?` from an unknown-function lookup.
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

/// Registers the service-stub and no-infrastructure built-ins (IMAGE, RTD,
/// TRANSLATE, DETECTLANGUAGE, COPILOT, PHONETIC, GETPIVOTDATA) into
/// `registry`. Intended to be invoked from `register_builtins`. ISOMITTED
/// is registered separately as a lazy special form; see the file-level
/// comment.
void register_service_stub_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_SERVICE_STUBS_H_
