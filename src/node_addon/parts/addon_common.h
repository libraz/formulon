// Shared translation helpers and module-global state for the Node.js
// N-API addon. Per-area implementation TUs (`parts/*.cc`) include this
// header to obtain the small kit of `Make*Status`, `TranslateValue`,
// and JS-shape builders that the bindings collectively rely on.
//
// All helpers translate the C ABI types declared in
// `c_api/formulon_c.h` (`fm_value_t`, `fm_cf_match_t`, `fm_pivot_*`,
// etc.) into the JS object shapes documented in `addon.cc`. The shape
// MUST stay byte-identical with the WASM/embind binding so JS callers
// can swap packages transparently.

#ifndef FORMULON_NODE_ADDON_PARTS_ADDON_COMMON_H_
#define FORMULON_NODE_ADDON_PARTS_ADDON_COMMON_H_

// NOTE: `NODE_ADDON_API_DISABLE_CPP_EXCEPTIONS` and
// `NAPI_DISABLE_CPP_EXCEPTIONS` are defined on the command line by
// `cmake/FormulonNodeAddon.cmake` so they apply uniformly to every TU
// that includes `napi.h`. They are required for the addon to compile
// under the project's `-fno-exceptions` policy.
//
// The compiler driver assigns the implicit replacement list `1` to a
// `-D X` flag, but `napi.h` later does an unconditional `#define X`
// (empty replacement list). To keep the build `-Werror`-clean we undef
// the command-line versions first; the napi.h `#ifdef NAPI_DISABLE_*`
// blocks immediately afterwards re-establish the same macros with the
// expected empty replacement list.
#ifdef NAPI_DISABLE_CPP_EXCEPTIONS
#undef NAPI_DISABLE_CPP_EXCEPTIONS
#endif
#ifdef NODE_ADDON_API_DISABLE_CPP_EXCEPTIONS
#undef NODE_ADDON_API_DISABLE_CPP_EXCEPTIONS
#endif
#define NAPI_DISABLE_CPP_EXCEPTIONS
#define NODE_ADDON_API_DISABLE_CPP_EXCEPTIONS

// NOLINTNEXTLINE(misc-include-cleaner): napi.h is the canonical entry point.
#include <napi.h>

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"

namespace formulon_node {

/// `kBindingNullPointer` ordinal mirrors `formulon::FormulonErrorCode`
/// in the 7000-7999 range allocated to bindings (see CLAUDE.md error
/// code table). The C ABI itself returns this code when a NULL pointer
/// crosses the boundary; we emit the same code from the JS side when
/// the wrapper is asked to operate on a destroyed handle.
constexpr fm_status_t kBindingNullPointer = 7000;

// ---------------------------------------------------------------------
// Status / Value envelope builders
// ---------------------------------------------------------------------

/// Builds an `ok` Status envelope:
///   { ok: true, status: 0, message: "", context: "" }
Napi::Object MakeOkStatus(Napi::Env env);

/// Builds an error Status envelope, copying the thread-local
/// diagnostics surfaced by the most recent C-ABI call.
Napi::Object MakeErrorStatus(Napi::Env env, fm_status_t code);

/// Converts a C-ABI status code into the shared JS Status envelope.
Napi::Object MakeStatus(Napi::Env env, fm_status_t code);

/// Translates an `fm_value_t` into the JS Value shape.
Napi::Object TranslateValue(Napi::Env env, const fm_value_t& v);

/// Builds `{ status, <field>: value }`.
Napi::Object MakeFieldResult(Napi::Env env, Napi::Object status, const char* field, Napi::Value value);

/// Builds `{ status, <field>: number }`.
Napi::Object MakeNumberFieldResult(Napi::Env env, Napi::Object status, const char* field, double value);

/// Builds `{ status, <field>: string }`. NULL `value` becomes "".
Napi::Object MakeStringFieldResult(Napi::Env env, Napi::Object status, const char* field, const char* value);

/// Builds `{ status, value: TranslateValue(value) }`.
Napi::Object MakeValueResult(Napi::Env env, Napi::Object status, const fm_value_t& value);

/// Builds `{ status, value: TranslateValue(blank) }`.
Napi::Object MakeEmptyValueResult(Napi::Env env, Napi::Object status);

/// Builds `{ status, index }` (used by `*_create` / `*_add` style entries).
Napi::Object MakeIndexResult(Napi::Env env, Napi::Object status, uint32_t index);

/// Builds `{ status, top, left, rows, cols, cells: [] }` placeholder for
/// failure paths in `PivotLayout`.
Napi::Object EmptyPivotLayoutResult(Napi::Env env, Napi::Object status);

/// Translates an `fm_pivot_cell_t` into the JS shape.
Napi::Object TranslatePivotCell(Napi::Env env, const fm_pivot_cell_t& cell);

/// Translates an `fm_cf_color_t` into the JS shape used by embind.
Napi::Object TranslateCfColor(Napi::Env env, const fm_cf_color_t& c);

/// Translates an `fm_cf_match_t` into the JS shape used by embind.
Napi::Object TranslateCfMatch(Napi::Env env, const fm_cf_match_t& m);

// ---------------------------------------------------------------------
// JS-spec pullers (read optional fields out of JS objects)
// ---------------------------------------------------------------------

/// Pulls an optional int32 field from a JS spec object; returns `dflt`
/// when the field is missing / undefined / null.
int32_t SpecPullInt32(const Napi::Object& spec, const char* key, int32_t dflt);

/// Pulls an optional uint32 field; returns `dflt` when missing.
uint32_t SpecPullU32(const Napi::Object& spec, const char* key, uint32_t dflt);

/// Pulls an optional double field; returns `dflt` when missing.
double SpecPullDouble(const Napi::Object& spec, const char* key, double dflt);

/// Pulls an optional bool field; returns `dflt` when missing.
bool SpecPullBool(const Napi::Object& spec, const char* key, bool dflt);

/// Returns whether the spec carries a non-null entry for `key`. Used to
/// decide whether a `const char*` field should be forwarded as nullptr.
bool SpecHas(const Napi::Object& spec, const char* key);

/// Pulls a numeric array from `info[idx]` into a `std::vector<uint32_t>`.
/// Returns an empty vector when the argument is missing / undefined /
/// null or not array-shaped.
std::vector<uint32_t> ReadU32Array(const Napi::CallbackInfo& info, size_t idx);

/// Builds an `fm_pivot_data_field_spec_t` from a JS spec object. The
/// `name_buf` / `nfmt_buf` strings keep the borrowed `const char*`
/// pointers alive for the caller; `has_nfmt` is set when the spec
/// carries a non-null `numberFormat`.
void BuildDataFieldSpec(const Napi::Object& spec, fm_pivot_data_field_spec_t& out, std::string& name_buf,
                        std::string& nfmt_buf, bool& has_nfmt);

// ---------------------------------------------------------------------
// JS-side iterative-progress callback slot
// ---------------------------------------------------------------------
//
// The C ABI's iterative solver is synchronous: it invokes the
// registered C callback inline from `fm_workbook_recalc` /
// `fm_workbook_partial_recalc` on the calling thread. That means the
// JS function we hold here is always invoked on the same thread that
// drove recalc, and we can safely call it through the standard
// `Napi::FunctionReference::Call` API (no thread-safe-function plumbing
// is required).
//
// Mirrors the embind binding's single-slot policy: there is one JS
// callback for the whole module, installing a new one displaces the
// previous, and clearing it (passing `null`) reverts to the default
// "always continue" behaviour.
struct ProgressSlot {
  Napi::FunctionReference fn;
  bool installed = false;
};

/// Function-local static keeps the slot alive for the addon's lifetime
/// without needing eager static initialisation of a Napi::Reference.
ProgressSlot& js_progress_slot();

/// C-ABI compatible trampoline that forwards into the held JS callback.
/// Returning `false` from the JS side aborts the iterative solve.
bool IterativeProgressTrampoline(uint32_t iteration, double max_residual, uint32_t max_iterations, void* user_data);

}  // namespace formulon_node

#endif  // FORMULON_NODE_ADDON_PARTS_ADDON_COMMON_H_
