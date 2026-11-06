//
// Implementation of Formulon's host-service built-in stubs. Each member
// of this family returns a deterministic Excel-visible surface in lieu
// of a feature the engine does not yet provide:
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
// All host-service stubs ride the eager dispatch path
// (`accepts_ranges = false`, default `propagate_errors = true`) so an
// error argument short-circuits before the fixed return fires -- this
// matches the WEBSERVICE / PY stubs in `src/eval/builtins/web.cpp` and
// keeps the surface consistent with real functions for formulas that
// propagate errors through service calls.
//
// GETPIVOTDATA is NOT a stub. It is a real lazy form implemented in
// `eval_getpivotdata_lazy` (see `getpivotdata_lazy.cpp`) and wired
// through `tree_walker.cpp`'s lazy dispatch table. Its second argument
// must reach the impl as an un-evaluated `Ref` AST so the resolver can
// recover the (sheet, row, col) anchor; the eager dispatcher would
// flatten the argument to a Value before that happened. Mirrors the
// PHONETIC pattern.
//
// PHONETIC is NOT a stub. It is a real lazy form implemented in
// `eval_phonetic_lazy` (see `phonetic_lazy.cpp`) and wired through
// `tree_walker.cpp`'s lazy dispatch table. The lazy form reads the
// referenced cell's `<rPh>` annotation directly off the un-evaluated
// Reference AST; the eager dispatcher would flatten the argument to a
// Value before the impl could consult `Cell::phonetic_text`.
//
// ISOMITTED is NOT a stub. It is a real lazy special form implemented in
// `eval_isomitted_lazy` (see `special_forms_lazy.cpp`) and registered through
// `tree_walker.cpp`'s lazy dispatch table. The eager registry cannot see the
// argument's AST shape or the active `NameEnv`'s omitted flag, both of which
// the implementation needs.

#include "eval/builtins/service_stubs.h"

#include <cstdint>

#include "eval/builtins/registration_helpers.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
void register_service_stub_builtins(FunctionRegistry& registry) {
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
      // IMAGE(source, [alt_text], [sizing], [height], [width]) -- arity 1..5.
      {"IMAGE", 1u, 5u, &builtins_detail::fixed_error_stub<ErrorCode::Value>},
      // RTD(progID, server, topic1, [topic2], ...).
      {"RTD", 3u, kVariadic, &builtins_detail::fixed_error_stub<ErrorCode::NA>},
      // TRANSLATE(text, [source_lang], [target_lang]) -- both languages are
      // optional at entry in Excel 365.
      {"TRANSLATE", 1u, 3u, &builtins_detail::fixed_error_stub<ErrorCode::Name>},
      {"DETECTLANGUAGE", 1u, 1u, &builtins_detail::fixed_error_stub<ErrorCode::Name>},
      // COPILOT(prompt, [context...]) -- variadic context cells.
      {"COPILOT", 1u, kVariadic, &builtins_detail::fixed_error_stub<ErrorCode::Name>},
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));
}

}  // namespace eval
}  // namespace formulon
