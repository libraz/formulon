// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// TU-agnostic marshalling leaves shared by the three binding translation
// units (`src/c_api/formulon_c.cpp`, `src/wasm/embind.cpp`,
// `src/node_addon/addon.cc`). Each binding still owns its own output
// emission (the `fm_value_t` struct fill, embind `val`, Napi `Object`),
// but every helper here is a small, total function over `formulon::Value`
// or `formulon::ErrorCode` that the bindings previously inlined three
// slightly different ways.
//
// Scope deliberately narrow: pure leaf inspection / re-export only. No
// memory ownership, no `TextStore`, no JS-engine types. Everything is
// `noexcept`.

#ifndef FORMULON_C_API_INTERNAL_VALUE_MARSHAL_H_
#define FORMULON_C_API_INTERNAL_VALUE_MARSHAL_H_

#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "value.h"

namespace formulon {
namespace c_api {
namespace internal {

/// Mirror of `fm_value_kind_t` in strongly-typed form.
///
/// The numeric values are load-bearing: each enumerator equals the
/// corresponding `FM_VAL_*` constant exposed by `formulon_c.h`, which in
/// turn equals the `formulon::ValueKind` ordinal. Bindings rely on this
/// chain of identities to `static_cast` between the three types without
/// any translation table.
///
/// Compile-time `static_assert`s in `value_marshal.cpp` enforce the
/// parity. Reordering this enum without updating `FM_VAL_*` (or vice
/// versa) is a build break.
enum class FmValueTag : std::int32_t {
  Blank = FM_VAL_BLANK,
  Number = FM_VAL_NUMBER,
  Bool = FM_VAL_BOOL,
  Text = FM_VAL_TEXT,
  Error = FM_VAL_ERROR,
  Array = FM_VAL_ARRAY,
  Ref = FM_VAL_REF,
  Lambda = FM_VAL_LAMBDA,
};

/// Returns the value-tag matching `v.kind()`.
///
/// Total over every `ValueKind` enumerator; never traps. The mapping is
/// 1:1 with `ValueKind`, so this is effectively a no-op cast wrapped in
/// type safety.
FmValueTag value_tag(const Value& v) noexcept;

/// Returns the OOXML wire code for `ec` (e.g. `7` for `#DIV/0!`).
///
/// Thin wrapper around `formulon::ooxml_code` re-exported in the binding
/// namespace so all three TUs reach the single source of truth via the
/// same symbol. Returns a `std::int32_t` to match the C ABI
/// (`fm_value_t::u.error_code`) without further casts at the call site.
std::int32_t error_to_fm_code(ErrorCode ec) noexcept;

/// Returns the tokenised Excel display name for `ec` (e.g. `"#DIV/0!"`).
///
/// Thin wrapper around `formulon::display_name`. The returned pointer
/// references a static string literal with program lifetime, identical
/// to the underlying call.
const char* error_to_fm_text(ErrorCode ec) noexcept;

/// Shape descriptor for `ArrayValue`. `empty` is true when the underlying
/// pointer is null or either dimension is zero, which collapses the
/// "empty array" check that all three bindings spell slightly
/// differently.
struct ArrayShape {
  std::uint32_t rows;
  std::uint32_t cols;
  bool empty;
};

/// Inspects the array payload of an `Array`-kind `Value`.
///
/// Precondition: `v.kind() == ValueKind::Array`. Returns the row/column
/// counts from the arena-backed `ArrayValue` and a derived `empty` flag
/// (true for null payload, 0-row, or 0-column matrices). Bindings use
/// this to short-circuit the "report kind but no payload" passthrough
/// path uniformly.
ArrayShape inspect_array(const Value& v) noexcept;

/// Returns `v.as_number()` when `v` is `Number`, otherwise `0.0`.
///
/// Collapses the `is_number() ? as_number() : 0.0` ternary that the
/// bindings repeat at every Value -> JS scalar conversion site. Never
/// traps regardless of `v.kind()`.
double value_as_number_or_zero(const Value& v) noexcept;

/// Returns `v.as_boolean()` when `v` is `Bool`, otherwise `false`.
///
/// Collapses the `is_boolean() ? as_boolean() : false` ternary that the
/// bindings repeat at every Value -> JS scalar conversion site. Never
/// traps regardless of `v.kind()`.
bool value_as_bool_or_false(const Value& v) noexcept;

/// Returns a pointer to a freshly constructed `std::string` owned by
/// `storage` and reflecting the text payload of `v`, or `nullptr` when
/// `v` is not `Text`.
///
/// The signature returns a pointer into `storage` rather than a value so
/// callers can distinguish "no text" from "empty text" without sentinels.
/// `storage` is constructed inside the function and its address is only
/// stable for the lifetime of `storage` at the call site.
const std::string* value_as_text_or_null(const Value& v, std::string& storage) noexcept;

}  // namespace internal
}  // namespace c_api
}  // namespace formulon

#endif  // FORMULON_C_API_INTERNAL_VALUE_MARSHAL_H_
